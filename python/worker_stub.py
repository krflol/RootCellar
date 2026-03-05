"""RootCellar Python macro worker for process-isolated script execution.

Reads one JSON request from stdin and writes one JSON response to stdout.
"""

from __future__ import annotations

import contextlib
import io
import importlib.util
import json
import re
import sys
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    import openpyxl  # type: ignore
except Exception as exc:
    openpyxl = None
    OPENPYXL_IMPORT_ERROR = str(exc)
else:
    OPENPYXL_IMPORT_ERROR = None


class ScriptWorkerError(RuntimeError):
    """Worker-facing error with optional additional details."""

    def __init__(self, message: str, *, details: Optional[str] = None) -> None:
        super().__init__(message)
        self.message = message
        self.details = details


@dataclass
class ScriptRequest:
    command: str
    trace_id: str
    script_path: str
    macro_name: str
    workbook_path: str
    permissions: List[str]
    args: Dict[str, str]

    @classmethod
    def from_json(cls, payload: Dict[str, Any]) -> "ScriptRequest":
        return cls(
            command=str(payload["command"]),
            trace_id=str(payload["trace_id"]),
            script_path=str(payload["script_path"]),
            macro_name=str(payload["macro_name"]),
            workbook_path=str(payload["workbook_path"]),
            permissions=[str(permission) for permission in payload.get("permissions", [])],
            args={str(k): str(v) for k, v in (payload.get("args", {}) or {}).items()},
        )


@dataclass
class ScriptPermissionEvent:
    event_name: str
    permission: str
    allowed: bool
    reason: str

    def to_payload(self) -> Dict[str, Any]:
        return {
            "event_name": self.event_name,
            "permission": self.permission,
            "allowed": self.allowed,
            "reason": self.reason,
        }


@dataclass
class ScriptRuntimeEvent:
    event_name: str
    payload: Any
    severity: Optional[str] = None

    def to_payload(self) -> Dict[str, Any]:
        return {
            "event_name": self.event_name,
            "payload": self.payload,
            "severity": self.severity,
        }


@dataclass
class ScriptMutation:
    op: str
    payload: Dict[str, Any]

    @classmethod
    def set_cell_value(cls, sheet: str, cell: str, value: Any) -> "ScriptMutation":
        return cls("set_cell_value", {"sheet": sheet, "cell": cell, "value": value})

    @classmethod
    def set_cell_formula(cls, sheet: str, cell: str, formula: str) -> "ScriptMutation":
        return cls("set_cell_formula", {"sheet": sheet, "cell": cell, "formula": formula})

    @classmethod
    def set_cell_range_value(cls, sheet: str, start: str, end: str, value: Any) -> "ScriptMutation":
        return cls("set_cell_range_value", {"sheet": sheet, "start": start, "end": end, "value": value})

    @classmethod
    def set_cell_range_formula(
        cls,
        sheet: str,
        start: str,
        end: str,
        formula: str,
    ) -> "ScriptMutation":
        return cls(
            "set_cell_range_formula",
            {"sheet": sheet, "start": start, "end": end, "formula": formula},
        )

    def to_payload(self) -> Dict[str, Any]:
        if self.op in {"set_cell_value", "set_cell_formula"}:
            payload = {
                "op": self.op,
                "sheet": self.payload["sheet"],
                "cell": self.payload.get("cell"),
            }
            if self.op == "set_cell_value":
                payload["value"] = serialize_cell_value(self.payload.get("value"))
            else:
                payload["formula"] = self.payload.get("formula")
            return payload

        payload = {
            "op": self.op,
            "sheet": self.payload["sheet"],
            "start": self.payload["start"],
            "end": self.payload["end"],
        }
        if self.op == "set_cell_range_value":
            payload["value"] = serialize_cell_value(self.payload.get("value"))
        else:
            payload["formula"] = self.payload.get("formula")
        return payload


def normalize_formula(formula: str) -> str:
    text = str(formula)
    if text.startswith("="):
        return text
    return f"={text}"


def serialize_cell_value(value: Any) -> Dict[str, Any]:
    if value is None:
        return {"kind": "empty"}
    if isinstance(value, bool):
        return {"kind": "bool", "value": value}
    if isinstance(value, (int, float)):
        return {"kind": "number", "value": float(value)}
    if isinstance(value, str):
        if value.startswith("#"):
            return {"kind": "error", "value": value}
        return {"kind": "text", "value": value}
    return {"kind": "text", "value": str(value)}


def parse_a1(cell_ref: str) -> Tuple[int, int]:
    match = re.fullmatch(r"\s*([A-Za-z]+)(\d+)\s*", cell_ref or "")
    if not match:
        raise ScriptWorkerError(f"invalid cell reference: {cell_ref}")

    col_text = match.group(1).upper()
    row = int(match.group(2))
    if row < 1:
        raise ScriptWorkerError(f"invalid row in cell reference: {cell_ref}")

    col = 0
    for char in col_text:
        col = (col * 26) + (ord(char) - ord("A") + 1)
    if col < 1:
        raise ScriptWorkerError(f"invalid column in cell reference: {cell_ref}")
    return row, col


def parse_range(range_ref: str) -> Tuple[str, str, int, int, int, int]:
    parts = [part.strip() for part in (range_ref or "").split(":")]
    if len(parts) != 2:
        raise ScriptWorkerError(f"invalid range reference: {range_ref}")
    start_cell, end_cell = parts
    start_row, start_col = parse_a1(start_cell)
    end_row, end_col = parse_a1(end_cell)
    return (
        start_cell.upper(),
        end_cell.upper(),
        min(start_row, end_row),
        min(start_col, end_col),
        max(start_row, end_row),
        max(start_col, end_col),
    )


def normalize_a1(row: int, col: int) -> str:
    if row < 1 or col < 1:
        raise ScriptWorkerError(f"invalid index to A1 conversion: row={row}, col={col}")
    letters: List[str] = []
    cursor = col
    while cursor > 0:
        cursor -= 1
        letters.append(chr(ord("A") + (cursor % 26)))
        cursor //= 26
    return "".join(reversed(letters)) + str(row)


def iter_range(
    start_row: int, start_col: int, end_row: int, end_col: int
) -> Iterable[Tuple[str, int, int]]:
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            yield normalize_a1(row, col), row, col


class PermissionState:
    def __init__(self, granted_permissions: Iterable[str]) -> None:
        self.granted = {permission.strip() for permission in granted_permissions if permission.strip()}
        self.events: List[ScriptPermissionEvent] = []

    def check(self, permission: str, purpose: str) -> bool:
        allowed = permission in self.granted
        self.events.append(
            ScriptPermissionEvent(
                event_name=("script.permission.granted" if allowed else "script.permission.denied"),
                permission=permission,
                allowed=allowed,
                reason=(f"{permission} requested for {purpose}") if allowed else "permission denied",
            )
        )
        return allowed


class RuntimeEventSink:
    def __init__(self) -> None:
        self.events: List[ScriptRuntimeEvent] = []

    def emit(self, event_name: str, payload: Any, *, severity: Optional[str] = None) -> None:
        self.events.append(
            ScriptRuntimeEvent(
                event_name=event_name,
                payload=_normalize_runtime_payload(payload),
                severity=severity,
            )
        )


class WorkbookRuntime:
    def __init__(self, path: str, permissions: PermissionState):
        if openpyxl is None:
            raise ScriptWorkerError(
                "openpyxl is not available for runtime workbook mutation",
                details=OPENPYXL_IMPORT_ERROR,
            )
        try:
            self.workbook = openpyxl.load_workbook(Path(path))
        except Exception as exc:
            raise ScriptWorkerError(
                f"failed to load workbook '{path}'",
                details=str(exc),
            ) from exc
        self.permissions = permissions
        self.mutations: List[ScriptMutation] = []

    def _ensure_sheet(self, sheet: str):
        try:
            return self.workbook[sheet]
        except KeyError as exc:
            raise ScriptWorkerError(f"sheet '{sheet}' does not exist in workbook") from exc

    @staticmethod
    def _coerce_value(value: Any) -> Any:
        if isinstance(value, (str, bool, int, float)) or value is None:
            return value
        return str(value)

    def get_cell(self, sheet: str, cell: str) -> Any:
        if not self.permissions.check("fs.read", f"get_cell({sheet},{cell})"):
            raise ScriptWorkerError("permission denied for fs.read")
        return self._ensure_sheet(sheet)[cell].value

    def set_cell_value(self, sheet: str, cell: str, value: Any) -> None:
        if not self.permissions.check("fs.write", f"set_cell_value({sheet},{cell})"):
            raise ScriptWorkerError("permission denied for fs.write")
        self._ensure_sheet(sheet)[cell] = self._coerce_value(value)
        self.mutations.append(ScriptMutation.set_cell_value(sheet, cell.upper(), value))

    def set_cell_formula(self, sheet: str, cell: str, formula: str) -> None:
        if not self.permissions.check("fs.write", f"set_cell_formula({sheet},{cell})"):
            raise ScriptWorkerError("permission denied for fs.write")
        normalized = normalize_formula(formula)
        self._ensure_sheet(sheet)[cell] = normalized
        self.mutations.append(
            ScriptMutation.set_cell_formula(sheet, cell.upper(), normalized),
        )

    def set_range_values(self, sheet: str, range_ref: str, value: Any) -> None:
        if ":" not in range_ref:
            self.set_cell_value(sheet, range_ref, value)
            return
        start_cell, end_cell, start_row, start_col, end_row, end_col = parse_range(range_ref)
        if not self.permissions.check("fs.write", f"set_range_values({sheet},{range_ref})"):
            raise ScriptWorkerError("permission denied for fs.write")
        worksheet = self._ensure_sheet(sheet)
        value = self._coerce_value(value)
        for coord, _, _ in iter_range(start_row, start_col, end_row, end_col):
            worksheet[coord] = value
        self.mutations.append(
            ScriptMutation.set_cell_range_value(sheet, start_cell, end_cell, value),
        )

    def set_range_formulas(self, sheet: str, range_ref: str, formula: str) -> None:
        if ":" not in range_ref:
            self.set_cell_formula(sheet, range_ref, formula)
            return
        start_cell, end_cell, start_row, start_col, end_row, end_col = parse_range(range_ref)
        if not self.permissions.check("fs.write", f"set_range_formulas({sheet},{range_ref})"):
            raise ScriptWorkerError("permission denied for fs.write")
        normalized = normalize_formula(formula)
        worksheet = self._ensure_sheet(sheet)
        for coord, _, _ in iter_range(start_row, start_col, end_row, end_col):
            worksheet[coord] = normalized
        self.mutations.append(
            ScriptMutation.set_cell_range_formula(sheet, start_cell, end_cell, normalized),
        )

    def read_range(self, sheet: str, range_ref: str) -> List[List[Any]]:
        if ":" not in range_ref:
            raise ScriptWorkerError(f"invalid range reference: {range_ref}")
        if not self.permissions.check("fs.read", f"read_range({sheet},{range_ref})"):
            raise ScriptWorkerError("permission denied for fs.read")
        worksheet = self._ensure_sheet(sheet)
        _, _, start_row, start_col, end_row, end_col = parse_range(range_ref)
        values: List[List[Any]] = []
        for row in range(start_row, end_row + 1):
            row_values = []
            for col in range(start_col, end_col + 1):
                row_values.append(worksheet[f"{normalize_a1(row, col)}"].value)
            values.append(row_values)
        return values

    def snapshot_mutations(self) -> List[Dict[str, Any]]:
        return [mutation.to_payload() for mutation in self.mutations]


class ScriptObject:
    def __init__(
        self,
        workbook: WorkbookRuntime,
        module: Any,
        permissions: PermissionState,
        args: Dict[str, str],
        runtime_events: RuntimeEventSink,
    ):
        self._workbook = workbook
        self.io = ScriptIO(permissions)
        self.net = ScriptNetwork(permissions)
        self.process = ScriptProcess(permissions)
        self.udf = ScriptUDF(module, permissions)
        self.events = ScriptEvents(permissions, runtime_events)
        self.args = dict(args)

    def cell(self, sheet: str, cell: str) -> Any:
        return self._workbook.get_cell(sheet, cell)

    def range_values(self, sheet: str, range_ref: str) -> List[List[Any]]:
        return self._workbook.read_range(sheet, range_ref)

    def set_value(self, sheet: str, cell: str, value: Any) -> None:
        self._workbook.set_cell_value(sheet, cell, value)

    def set_formula(self, sheet: str, cell: str, formula: str) -> None:
        self._workbook.set_cell_formula(sheet, cell, formula)

    def set_range_values(self, sheet: str, range_ref: str, value: Any) -> None:
        self._workbook.set_range_values(sheet, range_ref, value)

    def set_range_formulas(self, sheet: str, range_ref: str, formula: str) -> None:
        self._workbook.set_range_formulas(sheet, range_ref, formula)


class ScriptIO:
    def __init__(self, permissions: PermissionState):
        self.permissions = permissions

    def read_text(self, path: str) -> str:
        if not self.permissions.check("fs.read", f"read_text({path})"):
            raise ScriptWorkerError("permission denied for fs.read")
        return Path(path).read_text(encoding="utf-8")

    def write_text(self, path: str, contents: str) -> None:
        if not self.permissions.check("fs.write", f"write_text({path})"):
            raise ScriptWorkerError("permission denied for fs.write")
        Path(path).write_text(str(contents), encoding="utf-8")

    def read_json(self, path: str) -> Any:
        if not self.permissions.check("fs.read", f"read_json({path})"):
            raise ScriptWorkerError("permission denied for fs.read")
        with Path(path).open("r", encoding="utf-8") as handle:
            return json.load(handle)

    def write_json(self, path: str, value: Any) -> None:
        if not self.permissions.check("fs.write", f"write_json({path})"):
            raise ScriptWorkerError("permission denied for fs.write")
        with Path(path).open("w", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, ensure_ascii=False)


class ScriptNetwork:
    def __init__(self, permissions: PermissionState):
        self.permissions = permissions

    def get(self, url: str) -> str:
        if not self.permissions.check("net.http", f"get({url})"):
            raise ScriptWorkerError("permission denied for net.http")
        import urllib.request

        with urllib.request.urlopen(url) as response:
            return response.read().decode("utf-8")


class ScriptProcess:
    def __init__(self, permissions: PermissionState):
        self.permissions = permissions

    def run(self, command: str) -> int:
        if not self.permissions.check("process.exec", f"run({command})"):
            raise ScriptWorkerError("permission denied for process.exec")
        process = __import__("subprocess")
        return int(process.run(command, shell=True).returncode or 0)


class ScriptUDF:
    def __init__(self, module: Any, permissions: PermissionState):
        self.module = module
        self.permissions = permissions

    def __call__(self, function_name: str, *args: Any, **kwargs: Any) -> Any:
        return self.invoke(function_name, *args, **kwargs)

    def invoke(self, function_name: str, *args: Any, **kwargs: Any) -> Any:
        if not self.permissions.check("udf", f"invoke({function_name})"):
            raise ScriptWorkerError("permission denied for udf")
        if not function_name:
            raise ScriptWorkerError("udf function name must be a non-empty string")
        if not hasattr(self.module, function_name):
            raise ScriptWorkerError(f"udf '{function_name}' is not defined in script module")
        function = getattr(self.module, function_name)
        if not callable(function):
            raise ScriptWorkerError(f"udf '{function_name}' is not callable")
        return function(*args, **kwargs)

    def call(self, function_name: str, *args: Any, **kwargs: Any) -> Any:
        return self.invoke(function_name, *args, **kwargs)


class ScriptEvents:
    def __init__(self, permissions: PermissionState, runtime_events: RuntimeEventSink):
        self.permissions = permissions
        self.runtime_events = runtime_events

    def emit(
        self,
        event_name: str,
        payload: Any = None,
        *,
        severity: Optional[str] = None,
    ) -> None:
        if not event_name:
            raise ScriptWorkerError("event name must be a non-empty string")
        if not self.permissions.check("events.emit", f"emit({event_name})"):
            raise ScriptWorkerError("permission denied for events.emit")
        normalized = _normalize_runtime_payload(payload)
        self.runtime_events.emit(event_name, normalized, severity=severity)


def _normalize_runtime_payload(payload: Any) -> Any:
    if payload is None:
        return {}
    try:
        return json.loads(json.dumps(payload))
    except TypeError as error:
        raise ScriptWorkerError(f"runtime event payload must be JSON-serializable: {error}") from error


def build_response(
    status: str,
    *,
    message: Optional[str] = None,
    stdout: Optional[str] = None,
    stderr: Optional[str] = None,
    permission_events: Optional[List[ScriptPermissionEvent]] = None,
    mutations: Optional[List[Dict[str, Any]]] = None,
    runtime_events: Optional[List[Dict[str, Any]]] = None,
    result: Optional[Any] = None,
) -> Dict[str, Any]:
    return {
        "status": status,
        "message": message,
        "stdout": stdout,
        "stderr": stderr,
        "permission_events": [event.to_payload() for event in (permission_events or [])],
        "runtime_events": runtime_events or [],
        "mutations": mutations or [],
        "result": result,
    }


def build_normalized_result(value: Any) -> Any:
    try:
        json.loads(json.dumps(value))
        return value
    except TypeError:
        return str(value)


def call_macro(context: ScriptObject, macro_name: str, module: Any, args: Dict[str, str]) -> Any:
    macro = getattr(module, macro_name, None)
    if macro is None:
        raise ScriptWorkerError(f"macro '{macro_name}' not found in script")
    if not callable(macro):
        raise ScriptWorkerError(f"macro '{macro_name}' is not callable")

    for shape in ((context, args), (context,), tuple()):
        try:
            return macro(*shape)
        except TypeError as exc:
            continue
    # Fall back to no-arg invocation if type-checker suggests one.
    try:
        return macro()
    except Exception as exc:
        raise ScriptWorkerError(f"macro invocation failed: {exc}") from exc


def load_script_module(script_path: Path) -> Any:
    spec = importlib.util.spec_from_file_location("rootcellar_user_script", script_path)
    if spec is None or spec.loader is None:
        raise ScriptWorkerError(f"unable to load macro from '{script_path}'")
    if script_path.parent:
        parent = str(script_path.parent.resolve())
        if parent not in sys.path:
            sys.path.insert(0, parent)
    module = importlib.util.module_from_spec(spec)
    sys.modules["rootcellar_user_script"] = module
    spec.loader.exec_module(module)
    return module


def main() -> int:
    stdout_capture = io.StringIO()
    stderr_capture = io.StringIO()
    permission_state = PermissionState([])
    runtime_events = RuntimeEventSink()

    try:
        request_text = sys.stdin.read()
        if not request_text.strip():
            raise ScriptWorkerError("worker request payload was empty")

        payload = json.loads(request_text)
        request = ScriptRequest.from_json(payload)

        if request.command != "macro.run":
            raise ScriptWorkerError(f"unsupported command: {request.command}")

        permission_state = PermissionState(request.permissions)
        script_path = Path(request.script_path)
        if not script_path.exists():
            raise ScriptWorkerError(f"macro script does not exist: {script_path}")

        workbook = WorkbookRuntime(request.workbook_path, permission_state)
        module = load_script_module(script_path)
        context = ScriptObject(
            workbook,
            module,
            permission_state,
            request.args,
            runtime_events,
        )

        with contextlib.redirect_stdout(stdout_capture), contextlib.redirect_stderr(stderr_capture):
            result = call_macro(context, request.macro_name, module, request.args)

        response = build_response(
            "ok",
            mutations=workbook.snapshot_mutations(),
            permission_events=permission_state.events,
            runtime_events=[event.to_payload() for event in runtime_events.events],
            result=build_normalized_result(result),
            stdout=stdout_capture.getvalue(),
            stderr=stderr_capture.getvalue(),
        )
    except ScriptWorkerError as error:
        response = build_response(
            "error",
            message=error.message,
            stderr=error.details or stderr_capture.getvalue(),
            permission_events=permission_state.events,
        )
    except Exception as exc:
        response = build_response(
            "error",
            message=str(exc),
            stderr=f"{''.join(traceback.format_exception(type(exc), exc, exc.__traceback__))}\n{stderr_capture.getvalue()}",
            permission_events=permission_state.events,
        )

    if not response.get("stdout"):
        response["stdout"] = None
    print(json.dumps(response, sort_keys=True))
    return 0 if response["status"] == "ok" else 1


if __name__ == "__main__":
    raise SystemExit(main())
