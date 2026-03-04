export type TraceOutputArtifactRef = {
  artifactId: string;
  artifactType: string;
  relation: string;
  path?: string;
  description?: string;
};

export type TraceOutputMetadata = {
  traceId: string;
  traceRootId: string;
  commandStatus: string;
  spanId: string;
  parentSpanId: string | null;
  sessionId: string | null;
  uiCommandId?: string;
  uiCommandName?: string;
  eventLogPath?: string;
  artifactIndexPath?: string;
  durationMs?: number;
  linkedArtifactIds: string[];
  artifactRefs: TraceOutputArtifactRef[];
};

type TraceOutputMetadataLike = {
  traceId: string;
  traceRootId?: string;
  commandStatus?: string;
  spanId: string;
  parentSpanId: string | null;
  sessionId: string | null;
  uiCommandId?: string;
  uiCommandName?: string;
  eventLogPath?: string | null;
  artifactIndexPath?: string | null;
  durationMs?: number;
  linkedArtifactIds?: string[];
  artifactRefs?: TraceOutputArtifactRef[];
};

export const DESKTOP_TRACE_OUTPUT_PREFIX = "Latest trace:";

export function formatTraceHeader(
  trace: TraceOutputMetadataLike,
  commandId?: string,
): string {
  const parts: string[] = commandId || trace.uiCommandId ? [`ui_command_id=${commandId ?? trace.uiCommandId}`] : [];
  const resolvedCommandStatus = trace.commandStatus ?? "unknown";
  const artifactRefs = trace.artifactRefs ?? [];
  const linkedArtifactIds = trace.linkedArtifactIds ?? [];

  if (trace.uiCommandName) {
    parts.push(`ui_command_name=${trace.uiCommandName}`);
  }
  if (trace.eventLogPath) {
    parts.push(`event_log_path=${trace.eventLogPath}`);
  }
  if (trace.artifactIndexPath) {
    parts.push(`artifact_index_path=${trace.artifactIndexPath}`);
  }
  parts.push(
    `trace_id=${trace.traceId}`,
    `trace_root_id=${trace.traceRootId ?? trace.traceId}`,
    `command_status=${resolvedCommandStatus}`,
    `span_id=${trace.spanId}`,
    `parent_span_id=${trace.parentSpanId ?? "null"}`,
    `session_id=${trace.sessionId ?? "null"}`,
  );
  if (trace.durationMs !== undefined) {
    parts.push(`duration_ms=${trace.durationMs}`);
  }
  if (linkedArtifactIds.length > 0) {
    parts.push(`linked_artifact_ids=${linkedArtifactIds.join(",")}`);
  }
  if (artifactRefs.length > 0) {
    const refSummary = artifactRefs
      .map(
        (artifact) =>
          `${artifact.artifactType}|${artifact.relation}|${artifact.artifactId}${artifact.path ? `|${artifact.path}` : ""}`,
      )
      .join(";");
    parts.push(`artifact_refs=${refSummary}`);
  }
  return parts.join(" | ");
}

export function jsonWithTraceHeader(
  trace: TraceOutputMetadataLike,
  body: string,
  commandId?: string,
): string {
  return `${DESKTOP_TRACE_OUTPUT_PREFIX} ${formatTraceHeader(trace, commandId)}\n\n${body}`;
}

export function parseTraceHeaderLine(line: string): TraceOutputMetadata | null {
  const markerIndex = line.indexOf(DESKTOP_TRACE_OUTPUT_PREFIX);
  if (markerIndex === -1) {
    return null;
  }

  const payload = line
    .slice(markerIndex + DESKTOP_TRACE_OUTPUT_PREFIX.length)
    .trim()
    .replace(/^:\s*/, "");
  const segments = payload.split(" | ");

  const result: {
    traceId?: string;
    traceRootId?: string;
    commandStatus?: string;
    spanId?: string;
    parentSpanId?: string | null;
    sessionId?: string | null;
    uiCommandId?: string;
    uiCommandName?: string;
    eventLogPath?: string;
    artifactIndexPath?: string;
    durationMs?: number;
    linkedArtifactIds?: string[];
    artifactRefs?: TraceOutputArtifactRef[];
  } = {};

  for (const segment of segments) {
    const trimmedSegment = segment.trim();
    if (trimmedSegment.length === 0) {
      continue;
    }
    const separatorIndex = trimmedSegment.indexOf("=");
    if (separatorIndex < 0) {
      continue;
    }
    const key = trimmedSegment.slice(0, separatorIndex);
    const value = trimmedSegment.slice(separatorIndex + 1);
    switch (key) {
      case "trace_id": {
        result.traceId = value;
        break;
      }
      case "trace_root_id": {
        result.traceRootId = value;
        break;
      }
      case "command_status": {
        result.commandStatus = value;
        break;
      }
      case "span_id": {
        result.spanId = value;
        break;
      }
      case "parent_span_id": {
        result.parentSpanId = value === "null" ? null : value;
        break;
      }
      case "session_id": {
        result.sessionId = value === "null" ? null : value;
        break;
      }
      case "ui_command_id": {
        result.uiCommandId = value;
        break;
      }
      case "ui_command_name": {
        result.uiCommandName = value;
        break;
      }
      case "event_log_path": {
        result.eventLogPath = value;
        break;
      }
      case "artifact_index_path": {
        result.artifactIndexPath = value;
        break;
      }
      case "duration_ms": {
        result.durationMs = Number(value);
        break;
      }
      case "linked_artifact_ids": {
        result.linkedArtifactIds = value.length > 0 ? value.split(",") : [];
        break;
      }
      case "artifact_refs": {
        const parsed = value
          .split(";")
          .filter((raw) => raw.length > 0)
          .map((raw) => {
            const [artifactType, relation, artifactId, ...remainingPath] = raw.split("|");
            if (!artifactType || !relation || !artifactId) {
              return null;
            }
            const ref: TraceOutputArtifactRef = {
              artifactType,
              relation,
              artifactId,
            };
            if (remainingPath.length > 0) {
              ref.path = remainingPath.join("|");
            }
            return ref;
          });
        result.artifactRefs = parsed.filter(
          (entry): entry is TraceOutputArtifactRef => entry !== null,
        );
        break;
      }
      default:
        break;
    }
  }

  if (!result.traceId || !result.traceRootId || !result.commandStatus || !result.spanId) {
    return null;
  }

  return {
    traceId: result.traceId,
    traceRootId: result.traceRootId,
    commandStatus: result.commandStatus,
    spanId: result.spanId,
    parentSpanId: result.parentSpanId ?? null,
    sessionId: result.sessionId ?? null,
    uiCommandId: result.uiCommandId,
    uiCommandName: result.uiCommandName,
    eventLogPath: result.eventLogPath,
    artifactIndexPath: result.artifactIndexPath,
    durationMs: result.durationMs,
    linkedArtifactIds: result.linkedArtifactIds ?? [],
    artifactRefs: result.artifactRefs ?? [],
  };
}

export type SchemaValidationIssue = {
  path: string;
  message: string;
};

export type SchemaValidationResult = {
  ok: boolean;
  issues: SchemaValidationIssue[];
};

const UUID_V4_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function pushIssue(
  issues: SchemaValidationIssue[],
  path: string,
  message: string,
): void {
  issues.push({ path, message });
}

export function validateTraceOutputMetadata(metadata: unknown): SchemaValidationResult {
  const issues: SchemaValidationIssue[] = [];
  if (typeof metadata !== "object" || metadata === null) {
    return { ok: false, issues: [{ path: "$", message: "not an object" }] };
  }

  const candidate = metadata as Record<string, unknown>;

  const getValue = (names: readonly string[]): unknown | undefined => {
    for (const name of names) {
      if (Object.prototype.hasOwnProperty.call(candidate, name)) {
        return candidate[name];
      }
    }
    return undefined;
  };

  const requiredStringFields: Array<{
    path: string;
    candidates: readonly string[];
  }> = [
    { path: "trace_id", candidates: ["trace_id", "traceId"] },
    { path: "trace_root_id", candidates: ["trace_root_id", "traceRootId"] },
    { path: "command_status", candidates: ["command_status", "commandStatus"] },
    { path: "span_id", candidates: ["span_id", "spanId"] },
    { path: "session_id", candidates: ["session_id", "sessionId"] },
    { path: "parent_span_id", candidates: ["parent_span_id", "parentSpanId"] },
  ];

  for (const field of requiredStringFields) {
    const value = getValue(field.candidates);
    if (
      value === undefined ||
      (field.path !== "parent_span_id" && field.path !== "session_id" && value === null) ||
      (typeof value === "string" && value.length === 0)
    ) {
      issues.push({
        path: field.path,
        message: "missing required value",
      });
    }
  }

  const optionalUuidLikeFields: [string, string][] = [
    ["trace_id", "traceId"],
    ["trace_root_id", "traceRootId"],
    ["span_id", "spanId"],
    ["parent_span_id", "parentSpanId"],
    ["session_id", "sessionId"],
    ["ui_command_id", "uiCommandId"],
  ];
  for (const [path, camelCaseKey] of optionalUuidLikeFields) {
    const value = getValue([path, camelCaseKey]);
    if (value === null) {
      continue;
    }
    if (value !== undefined && (typeof value !== "string" || !UUID_V4_PATTERN.test(value))) {
      pushIssue(issues, path, "invalid UUID format");
    }
  }

  const durationValue = getValue(["duration_ms", "durationMs"]);
  if (durationValue !== undefined && durationValue !== null) {
    if (typeof durationValue !== "number" || !Number.isFinite(durationValue)) {
      pushIssue(issues, "duration_ms", "must be a finite number");
    }
  }

  const commandStatus = getValue(["command_status", "commandStatus"]);
  if (typeof commandStatus === "string" && commandStatus.trim().length === 0) {
    pushIssue(issues, "command_status", "must not be empty");
  } else if (commandStatus === undefined) {
    pushIssue(issues, "command_status", "missing required value");
  }

  const artifactIds = getValue(["linked_artifact_ids", "linkedArtifactIds"]);
  const artifactRefsValue = getValue(["artifact_refs", "artifactRefs"]);

  for (const field of ["linked_artifact_ids", "artifact_refs"] as const) {
    const value = field === "linked_artifact_ids" ? artifactIds : artifactRefsValue;
    if (value === undefined) {
      continue;
    }
    if (!Array.isArray(value)) {
      pushIssue(issues, field, "must be an array");
      continue;
    }
  }

  if (Array.isArray(artifactIds)) {
    artifactIds.forEach((artifactId, index) => {
      if (typeof artifactId !== "string" || artifactId.trim().length === 0) {
        pushIssue(
          issues,
          `linked_artifact_ids[${index}]`,
          "artifact ID must be a non-empty string",
        );
      }
    });
  }

  if (Array.isArray(artifactRefsValue)) {
    artifactRefsValue.forEach((entry, index) => {
      if (typeof entry !== "object" || entry === null) {
        pushIssue(issues, `artifact_refs[${index}]`, "must be an object");
        return;
      }
      const artifact = entry as Record<string, unknown>;
      const getArtifactValue = (names: readonly string[]): unknown => {
        for (const name of names) {
          if (Object.prototype.hasOwnProperty.call(artifact, name)) {
            return artifact[name];
          }
        }
        return undefined;
      };

      const artifactId = getArtifactValue(["artifact_id", "artifactId"]);
      const artifactType = getArtifactValue(["artifact_type", "artifactType"]);
      const relation = getArtifactValue(["relation"]);
      const path = getArtifactValue(["path"]);
      const description = getArtifactValue(["description"]);

      if (typeof artifactId !== "string" || artifactId.trim().length === 0) {
        pushIssue(issues, `artifact_refs[${index}].artifact_id`, "missing artifact_id");
      }
      if (typeof artifactType !== "string" || artifactType.trim().length === 0) {
        pushIssue(issues, `artifact_refs[${index}].artifact_type`, "missing artifact_type");
      }
      if (typeof relation !== "string" || relation.trim().length === 0) {
        pushIssue(issues, `artifact_refs[${index}].relation`, "missing relation");
      }
      if (
        path !== undefined &&
        path !== null &&
        typeof path !== "string"
      ) {
        pushIssue(issues, `artifact_refs[${index}].path`, "path must be a string");
      }
      if (
        description !== undefined &&
        description !== null &&
        typeof description !== "string"
      ) {
        pushIssue(
          issues,
          `artifact_refs[${index}].description`,
          "description must be a string",
        );
      }
    });
  }

  return {
    ok: issues.length === 0,
    issues,
  };
}
