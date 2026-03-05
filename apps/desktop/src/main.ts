import { invoke } from "@tauri-apps/api/core";
import { open as openDialog, save as saveDialog } from "@tauri-apps/plugin-dialog";
import { buildPresetRange } from "./editRangePresets";
import { bindFormulaBarApplyHandlers, computeNextPreviewSelection } from "./previewInteractions";
import {
  PRESET_DIMENSIONS_STORAGE_KEY,
  parsePresetDimensions,
  serializePresetDimensions,
  type PresetDimensions,
} from "./presetReuse";
import { renderLastCustomPresetButtonView } from "./presetReuseView";
import {
  markRecalcCompletedState,
  markRecalcPendingState,
  resetRecalcStatusState,
  type RecalcStatusState,
} from "./recalcFreshness";
import { renderRecalcStatusView } from "./recalcFreshnessView";
import { jsonWithTraceHeader, formatTraceHeader } from "./desktopTraceOutput";
import "./styles.css";

type TraceInput = {
  traceId: string;
  spanId: string;
  parentSpanId: string | null;
  sessionId: string | null;
  commandId?: string;
  commandName?: string;
  eventLogPath?: string;
  artifactIndexPath?: string | null;
};

type TraceArtifactRef = {
  artifactId: string;
  artifactType: string;
  relation: string;
  path?: string;
  description?: string;
};

type AppStatusResponse = {
  app: string;
  uiReady: boolean;
  engineReady: boolean;
  interopReady: boolean;
};

type TraceEcho = {
  traceId: string;
  spanId: string;
  parentSpanId: string | null;
  sessionId: string | null;
  uiCommandId?: string;
  uiCommandName?: string;
  traceRootId?: string;
  commandStatus?: string;
  durationMs?: number;
  eventLogPath?: string | null;
  artifactIndexPath?: string | null;
  linkedArtifactIds?: string[];
  artifactRefs?: TraceArtifactRef[];
};

function isTraceEcho(
  traceContext: TraceInput | TraceEcho,
): traceContext is TraceEcho & { uiCommandId?: string; uiCommandName?: string } {
  return "uiCommandId" in traceContext || "uiCommandName" in traceContext;
}

function toTraceInput(traceContext: TraceInput | TraceEcho): TraceInput {
  if (isTraceEcho(traceContext)) {
    return {
      traceId: traceContext.traceId,
      spanId: traceContext.spanId,
      parentSpanId: traceContext.parentSpanId,
      sessionId: traceContext.sessionId,
        commandId: traceContext.uiCommandId,
      commandName: traceContext.uiCommandName,
      eventLogPath:
        traceContext.eventLogPath === null ? undefined : traceContext.eventLogPath,
      artifactIndexPath:
        traceContext.artifactIndexPath === null ? undefined : traceContext.artifactIndexPath,
    };
  }

  return {
    traceId: traceContext.traceId,
    spanId: traceContext.spanId,
    parentSpanId: traceContext.parentSpanId,
    sessionId: traceContext.sessionId,
    commandId: traceContext.commandId,
    commandName: traceContext.commandName,
    eventLogPath: traceContext.eventLogPath,
    artifactIndexPath:
      traceContext.artifactIndexPath === null ? undefined : traceContext.artifactIndexPath,
  };
}

type UiCommandTrace = {
  traceId: string;
  spanId: string;
  parentSpanId: string | null;
  sessionId: string | null;
  commandId: string;
  commandName: string;
};

type CellValue =
  | { number: number }
  | { text: string }
  | { bool: boolean }
  | { error: string }
  | "empty";

type EngineRoundTripResponse = {
  sheet: string;
  formulaCell: string;
  value: CellValue;
  evaluatedCells: number;
  cycleCount: number;
  parseErrorCount: number;
  workbookId: string;
  trace: TraceEcho;
};

type CompatibilityStatus =
  | "supported"
  | "partially_supported"
  | "preserved_only"
  | "not_supported";

type CompatibilityIssue = {
  code: string;
  title: string;
  status: CompatibilityStatus;
  details: string;
};

type PartGraphSummary = {
  nodeCount: number;
  edgeCount: number;
  danglingEdgeCount: number;
  externalEdgeCount: number;
  unknownPartCount: number;
};

type InteropSessionStatusResponse = {
  loaded: boolean;
  inputPath: string | null;
  workbookId: string | null;
  sheetCount: number;
  cellCount: number;
  issueCount: number;
  unknownPartCount: number;
  dirtySheetCount: number;
  dirtySheets: string[];
  undoCount: number;
  redoCount: number;
  sheets: string[];
};

type InteropOpenResponse = {
  inputPath: string;
  workbookId: string;
  sheetCount: number;
  cellCount: number;
  sheets: string[];
  featureScore: number;
  issueCount: number;
  unknownPartCount: number;
  issues: CompatibilityIssue[];
  unknownParts: string[];
  partGraph: PartGraphSummary;
  trace: TraceEcho;
};

type InteropSaveMode = "preserve" | "normalize";

type SavePartGraphFlags = {
  strategy: string;
  sourceGraphReused: boolean;
  relationshipsPreserved: boolean;
  unknownPartsPreserved: boolean;
};

type InteropSaveResponse = {
  inputPath: string;
  outputPath: string;
  workbookId: string;
  mode: InteropSaveMode;
  sheetCount: number;
  cellCount: number;
  copiedBytes: number;
  partGraph: PartGraphSummary;
  partGraphFlags: SavePartGraphFlags;
  trace: TraceEcho;
};

type RecalcReport = {
  sheet: string;
  evaluatedCells: number;
  cycleCount: number;
  parseErrorCount: number;
};

type InteropRecalcResponse = {
  workbookId: string;
  reports: RecalcReport[];
  trace: TraceEcho;
};

type InteropEditMode = "value" | "formula";

type InteropCellEditResponse = {
  workbookId: string;
  sheet: string;
  cell: string;
  anchorCell: string;
  appliedCellCount: number;
  mode: InteropEditMode;
  value: CellValue;
  formula: string | null;
  dirtySheetCount: number;
  dirtySheets: string[];
  trace: TraceEcho;
};

type InteropPreviewCell = {
  cell: string;
  row: number;
  col: number;
  value: CellValue;
  formula: string | null;
};

type InteropSheetPreviewResponse = {
  workbookId: string;
  sheet: string;
  totalCells: number;
  shownCells: number;
  truncated: boolean;
  cells: InteropPreviewCell[];
  trace: TraceEcho;
};

type InteropUndoRedoResponse = {
  action: "undo" | "redo";
  workbookId: string;
  dirtySheetCount: number;
  dirtySheets: string[];
  undoCount: number;
  redoCount: number;
  trace: TraceEcho;
};

type InteropMacroPermissionConfig = {
  fsRead: boolean;
  fsWrite: boolean;
  netHttp: boolean;
  clipboard: boolean;
  processExec: boolean;
  udf: boolean;
  eventsEmit: boolean;
};

type MacroPermissionPolicySource = "stored" | "fresh";

const MACRO_PERMISSION_POLICY_STORAGE_KEY = "rootcellar.macro.permission.policy.v1";

type StoredMacroPermissionPolicy = {
  version: 1;
  permissions: InteropMacroPermissionConfig;
  createdAt: string;
  lastUsedAt: string;
  scriptPath: string;
};

type MacroPermissionPolicyStore = Record<string, StoredMacroPermissionPolicy>;

type InteropScriptPermissionEvent = {
  eventName: string;
  permission: string;
  allowed: boolean;
  reason: string;
};

type InteropScriptRuntimeEvent = {
  eventName: string;
  payload: unknown;
  severity?: string;
};

type InteropMacroTrustProvenance = {
  mode: string;
  manifestPath?: string;
  manifestName?: string;
  manifestVersion?: string;
  publisher?: string;
  apiMinVersion?: number;
  permissionsRequired: string[];
  permissionsDeclared: string[];
  runtimeApiVersion: number;
  signaturePresent: boolean;
  signatureVerified?: boolean;
  fingerprint: string;
  trusted: boolean;
};

type InteropMacroMutationPreview = {
  sheet: string;
  cell: string;
  kind: "value" | "formula";
  value?: CellValue;
  formula?: string;
};

type InteropRunMacroResponse = {
  workbookId: string;
  scriptPath: string;
  macroName: string;
  scriptFingerprint?: string;
  requestedPermissions: string[];
  permissionEvents: InteropScriptPermissionEvent[];
  trust?: InteropMacroTrustProvenance;
  runtimeEvents: InteropScriptRuntimeEvent[];
  permissionGranted: number;
  permissionDenied: number;
  mutationCount: number;
  changedSheets: string[];
  mutations: InteropMacroMutationPreview[];
  recalcReports: RecalcReport[];
  policySource?: MacroPermissionPolicySource;
  stdout?: string;
  stderr?: string;
  trace: TraceEcho;
};

type SelectedPreviewCell = {
  sheet: string;
  cell: InteropPreviewCell;
};

type EditActionSource = "edit-form" | "formula-bar";
type UiCaptureRecalcState = "pending" | "fresh" | "stale";
type UiCaptureSection = "default" | "edit-cell" | "save-recalc";
type EditLifecyclePhase = "start" | "success" | "error";
type EditLifecycleEntry = {
  ts: string;
  command: string;
  phase: EditLifecyclePhase;
  message: string;
  durationMs?: number;
  traceId?: string;
  uiCommandId?: string;
  commandStatus?: string;
  eventLogPath?: string;
  artifactIndexPath?: string;
  linkedArtifactIds?: string[];
  error?: string;
};

function formatCellValue(value: CellValue): string {
  if (typeof value === "string") {
    return value;
  }

  const [kind, raw] = Object.entries(value)[0];
  return `${kind}=${raw}`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function toMacroPolicyKey(scriptPath: string): string {
  return scriptPath.trim().toLowerCase();
}

function isInteropMacroPermissionConfig(value: unknown): value is InteropMacroPermissionConfig {
  if (!value || typeof value !== "object") {
    return false;
  }

  const candidate = value as {
    fsRead?: unknown;
    fsWrite?: unknown;
    netHttp?: unknown;
    clipboard?: unknown;
    processExec?: unknown;
    udf?: unknown;
    eventsEmit?: unknown;
  };
  return (
    typeof candidate.fsRead === "boolean" &&
    typeof candidate.fsWrite === "boolean" &&
    typeof candidate.netHttp === "boolean" &&
    typeof candidate.clipboard === "boolean" &&
    typeof candidate.processExec === "boolean" &&
    typeof candidate.udf === "boolean" &&
    typeof candidate.eventsEmit === "boolean"
  );
}

function isStoredMacroPermissionPolicy(
  value: unknown,
): value is StoredMacroPermissionPolicy {
  if (!value || typeof value !== "object") {
    return false;
  }

  const candidate = value as {
    version?: unknown;
    permissions?: unknown;
    createdAt?: unknown;
    lastUsedAt?: unknown;
    scriptPath?: unknown;
  };
  return (
    candidate.version === 1 &&
    typeof candidate.createdAt === "string" &&
    typeof candidate.lastUsedAt === "string" &&
    typeof candidate.scriptPath === "string" &&
    isInteropMacroPermissionConfig(candidate.permissions)
  );
}

function loadMacroPermissionPolicyStore(): MacroPermissionPolicyStore {
  let raw: string | null = null;
  try {
    raw = localStorage.getItem(MACRO_PERMISSION_POLICY_STORAGE_KEY);
  } catch {
    return {};
  }
  if (!raw) {
    return {};
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return {};
  }

  if (!parsed || typeof parsed !== "object") {
    return {};
  }

  const safeStore: MacroPermissionPolicyStore = {};
  const candidate = parsed as Record<string, unknown>;
  for (const [scriptPath, entry] of Object.entries(candidate)) {
    if (!isStoredMacroPermissionPolicy(entry)) {
      continue;
    }
    safeStore[scriptPath] = entry;
  }
  return safeStore;
}

function saveMacroPermissionPolicyStore(store: MacroPermissionPolicyStore): void {
  try {
    localStorage.setItem(
      MACRO_PERMISSION_POLICY_STORAGE_KEY,
      JSON.stringify(store),
    );
  } catch {
    return;
  }
}

function loadMacroPermissionPolicy(scriptPath: string): InteropMacroPermissionConfig | null {
  const policyKey = toMacroPolicyKey(scriptPath);
  if (policyKey.length === 0) {
    return null;
  }

  const store = loadMacroPermissionPolicyStore();
  const entry = store[policyKey];
  if (!entry || !isInteropMacroPermissionConfig(entry.permissions)) {
    return null;
  }
  return entry.permissions;
}

function loadMacroPermissionPolicyMetadata(
  scriptPath: string,
): StoredMacroPermissionPolicy | null {
  const policyKey = toMacroPolicyKey(scriptPath);
  if (policyKey.length === 0) {
    return null;
  }

  const store = loadMacroPermissionPolicyStore();
  const entry = store[policyKey];
  return isStoredMacroPermissionPolicy(entry) ? entry : null;
}

function saveMacroPermissionPolicy(
  scriptPath: string,
  permissions: InteropMacroPermissionConfig,
): void {
  const policyKey = toMacroPolicyKey(scriptPath);
  if (policyKey.length === 0) {
    return;
  }

  const store = loadMacroPermissionPolicyStore();
  const previous = store[policyKey];
  const now = new Date().toISOString();
  store[policyKey] = {
    version: 1,
    scriptPath,
    permissions,
    createdAt: previous?.createdAt ?? now,
    lastUsedAt: now,
  };
  saveMacroPermissionPolicyStore(store);
}

function deleteMacroPermissionPolicy(scriptPath: string): void {
  const policyKey = toMacroPolicyKey(scriptPath);
  if (policyKey.length === 0) {
    return;
  }
  const store = loadMacroPermissionPolicyStore();
  if (Object.prototype.hasOwnProperty.call(store, policyKey)) {
    delete store[policyKey];
    saveMacroPermissionPolicyStore(store);
  }
}

function permissionsEqual(left: InteropMacroPermissionConfig, right: InteropMacroPermissionConfig): boolean {
  return (
    left.fsRead === right.fsRead &&
    left.fsWrite === right.fsWrite &&
    left.netHttp === right.netHttp &&
    left.clipboard === right.clipboard &&
    left.processExec === right.processExec &&
    left.udf === right.udf &&
    left.eventsEmit === right.eventsEmit
  );
}

function hasElevatedMacroPermissions(config: InteropMacroPermissionConfig): boolean {
  return (
    config.fsWrite ||
    config.netHttp ||
    config.clipboard ||
    config.processExec ||
    config.udf ||
    config.eventsEmit
  );
}

function newTrace(): TraceInput {
  return {
    traceId: crypto.randomUUID(),
    spanId: crypto.randomUUID(),
    parentSpanId: null,
    sessionId: crypto.randomUUID(),
  };
}

function startUiCommand(commandName: string): UiCommandTrace {
  return {
    ...newTrace(),
    commandId: crypto.randomUUID(),
    commandName,
  };
}

const LIFECYCLE_LOG_LIMIT = 24;
const editLifecycleEntries: EditLifecycleEntry[] = [];

export function getEditLifecycleEntriesForTests(): EditLifecycleEntry[] {
  return [...editLifecycleEntries];
}

export function clearEditLifecycleEntriesForTests(): void {
  editLifecycleEntries.length = 0;
  renderEditLifecycleOutput();
}

export function logEditLifecycleEvent(
  command: string,
  phase: EditLifecyclePhase,
  message: string,
  options?: {
    trace?: TraceEcho;
    error?: unknown;
  },
): void {
  const now = new Date().toISOString();
  const trace = options?.trace;
  const errorText = options?.error ? String(options.error) : undefined;
  const entry: EditLifecycleEntry = {
    ts: now,
    command,
    phase,
    message,
    durationMs: trace?.durationMs,
    traceId: trace?.traceId,
    uiCommandId: trace?.uiCommandId,
    commandStatus: trace?.commandStatus,
    eventLogPath: trace?.eventLogPath ?? undefined,
    artifactIndexPath: trace?.artifactIndexPath ?? undefined,
    linkedArtifactIds: trace?.linkedArtifactIds,
    error: errorText,
  };
  editLifecycleEntries.unshift(entry);
  if (editLifecycleEntries.length > LIFECYCLE_LOG_LIMIT) {
    editLifecycleEntries.length = LIFECYCLE_LOG_LIMIT;
  }

  renderEditLifecycleOutput();

  if (phase === "error") {
    announceToScreenReader(`${command} failed: ${message}`, true);
  } else if (phase === "success") {
    announceToScreenReader(`${command} ${message}`);
  }
}

function renderEditLifecycleOutput(): void {
  if (editLifecycleEntries.length === 0) {
    editLifecycleOutputEl.textContent = "No edit lifecycle events yet.";
    return;
  }

  editLifecycleOutputEl.textContent = editLifecycleEntries
    .slice(0, LIFECYCLE_LOG_LIMIT)
    .map((entry) =>
      JSON.stringify(
        {
          ts: entry.ts,
          command: entry.command,
          phase: entry.phase,
          message: entry.message,
          durationMs: entry.durationMs,
          traceId: entry.traceId,
          uiCommandId: entry.uiCommandId,
          commandStatus: entry.commandStatus,
          eventLogPath: entry.eventLogPath,
          artifactIndexPath: entry.artifactIndexPath,
          linkedArtifactIds: entry.linkedArtifactIds,
          error: entry.error,
        },
        null,
        2,
      ),
    )
    .join("\n\n");
}

function announceToScreenReader(message: string, assertive = false): void {
  const live = assertive ? srAnnouncerAssertiveEl : srAnnouncerEl;
  if (!live) {
    return;
  }
  // Reset first so screen readers re-read repeated strings.
  live.textContent = "";
  requestAnimationFrame(() => {
    live.textContent = message;
  });
}

function normalizeMode(value: string): InteropSaveMode {
  return value === "normalize" ? "normalize" : "preserve";
}

function normalizeEditMode(value: string): InteropEditMode {
  return value === "formula" ? "formula" : "value";
}

function suggestOutputPath(inputPath: string, mode: InteropSaveMode): string {
  const suffix = `.rootcellar.${mode}.xlsx`;
  if (/\.xlsx$/i.test(inputPath)) {
    return inputPath.replace(/\.xlsx$/i, suffix);
  }
  return `${inputPath}${suffix}`;
}

function columnToA1(col: number): string {
  let value = col;
  let result = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result || "?";
}

function formatCompatibility(report: InteropOpenResponse): string {
  const lines: string[] = [
    `Feature score: ${report.featureScore}`,
    `Issues: ${report.issueCount}`,
    `Unknown parts: ${report.unknownPartCount}`,
    `Part graph: nodes=${report.partGraph.nodeCount}, edges=${report.partGraph.edgeCount}, dangling=${report.partGraph.danglingEdgeCount}, external=${report.partGraph.externalEdgeCount}`,
  ];

  if (report.issues.length > 0) {
    lines.push("");
    lines.push("Compatibility Issues:");
    for (const issue of report.issues) {
      lines.push(`- [${issue.status}] ${issue.code}: ${issue.title}`);
    }
  }

  if (report.unknownParts.length > 0) {
    lines.push("");
    lines.push("Unknown Parts (first 15):");
    for (const part of report.unknownParts.slice(0, 15)) {
      lines.push(`- ${part}`);
    }
    if (report.unknownParts.length > 15) {
      lines.push(`... ${report.unknownParts.length - 15} more`);
    }
  }

  return lines.join("\n");
}

function formatPreviewCell(cell: InteropPreviewCell): string {
  const value = formatCellValue(cell.value);
  return cell.formula ? `${cell.formula} => ${value}` : value;
}

function focusCaptureSection(section: UiCaptureSection): void {
  const id =
    section === "edit-cell"
      ? "capture-section-edit-cell"
      : section === "save-recalc"
        ? "capture-section-save-recalc"
        : null;
  if (!id) {
    return;
  }

  const sectionElement = document.getElementById(id);
  if (!sectionElement) {
    return;
  }

  sectionElement.scrollIntoView({ behavior: "auto", block: "start", inline: "nearest" });
}

const captureParams = new URLSearchParams(window.location.search);
const uiCaptureMode = captureParams.get("ui_capture") === "1";
const uiCaptureRecalcState = (() => {
  const value = (captureParams.get("capture_state") ?? "stale").trim().toLowerCase();
  if (value === "pending" || value === "fresh" || value === "stale") {
    return value as UiCaptureRecalcState;
  }
  return "stale";
})();
const uiCaptureSection = (() => {
  const value = (captureParams.get("capture_section") ?? "default").trim().toLowerCase();
  if (value === "edit-cell" || value === "save-recalc" || value === "default") {
    return value as UiCaptureSection;
  }
  return "default";
})();
const maxPresetRows = 1_048_576;
const maxPresetCols = 16_384;

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) {
  throw new Error("missing #app container");
}

app.innerHTML = `
  <main class="shell">
    <div id="sr-announcer" class="sr-only" role="status" aria-live="polite" aria-atomic="true"></div>
    <div id="sr-announcer-assertive" class="sr-only" role="status" aria-live="assertive" aria-atomic="true"></div>
    <section class="panel hero">
      <p class="eyebrow">RootCellar Desktop Shell</p>
      <h1>UI + Engine Vertical Slice</h1>
      <p class="summary">
        Build UI and core in tandem with a compatibility-first interop path.
      </p>
      <div class="actions">
        <button id="refresh-status" class="btn secondary">Refresh Status</button>
        <button id="run-round-trip" class="btn primary">Run Engine Round-Trip</button>
      </div>
    </section>

    <section class="panel">
      <h2>Interop Session</h2>
      <p class="summary">
        Open an Excel-authored <code>.xlsx</code>, inspect compatibility, and keep preserve-mode
        round-trips as the default path.
      </p>
      <label class="field">
        <span>Workbook Path</span>
        <input id="open-path" placeholder="C:\\path\\to\\workbook.xlsx" />
      </label>
      <div class="actions">
        <button id="pick-open-path" class="btn secondary">Choose Workbook...</button>
        <button id="open-workbook" class="btn primary">Open + Inspect</button>
        <button id="refresh-session" class="btn secondary">Refresh Session</button>
      </div>
      <pre id="session-output" class="output">No workbook loaded.</pre>
    </section>

    <section class="panel">
      <h2>Compatibility Panel</h2>
      <pre id="compat-output" class="output">No workbook loaded.</pre>
    </section>

    <section class="panel">
      <h2>Sheet Preview</h2>
      <div class="row">
        <label class="field compact">
          <span>Sheet</span>
          <select id="preview-sheet"></select>
        </label>
        <label class="field compact">
          <span>Cell Limit</span>
          <input id="preview-limit" value="120" />
        </label>
      </div>
      <div class="actions preview-actions">
        <button id="refresh-preview" class="btn secondary">Refresh Preview</button>
        <button id="jump-last-edited" class="btn secondary">Jump To Last Edit</button>
        <button id="copy-preview-a1" class="btn secondary">Copy A1</button>
        <button id="copy-preview-value" class="btn secondary">Copy Value</button>
        <button id="copy-preview-formula" class="btn secondary">Copy Formula</button>
        <button id="paste-clipboard-cell" class="btn secondary">Paste Into Cell</button>
      </div>
      <div class="formula-bar">
        <label class="field compact">
          <span>Selected</span>
          <input id="formula-bar-cell" readonly value="-" />
        </label>
        <label class="field compact">
          <span>Bar Mode</span>
          <select id="formula-bar-mode">
            <option value="value">Value</option>
            <option value="formula">Formula</option>
          </select>
        </label>
        <label class="field compact wide">
          <span>Formula Bar Input</span>
          <input id="formula-bar-input" placeholder="Type value or formula for selected cell" />
        </label>
        <button id="formula-bar-apply" class="btn primary">Apply From Bar</button>
      </div>
      <pre id="preview-output" class="output compact">No preview loaded.</pre>
      <pre id="preview-selected-output" class="output compact" role="status" aria-live="polite">No cell selected.</pre>
      <pre id="preview-copy-status" class="output compact" role="status" aria-live="polite">No copy action yet.</pre>
      <div id="preview-wrap" class="table-wrap" tabindex="0">
        <table id="preview-table" class="preview-table" role="grid" aria-label="Spreadsheet preview grid"></table>
      </div>
    </section>

    <section class="panel" id="capture-section-edit-cell">
      <h2>Edit Cell</h2>
      <div class="row three">
        <label class="field compact">
          <span>Sheet</span>
          <select id="edit-sheet"></select>
        </label>
        <label class="field compact">
          <span>Cell (A1 or A1:B3)</span>
          <input id="edit-cell" placeholder="A1 or A1:B3" value="A1" />
        </label>
        <label class="field compact">
          <span>Edit Mode</span>
          <select id="edit-mode">
            <option value="value">Value</option>
            <option value="formula">Formula</option>
          </select>
        </label>
      </div>
      <label class="field">
        <span>Input</span>
        <input id="edit-input" placeholder="10, true, text, or =A1+B1" />
      </label>
      <div class="actions compact-actions">
        <button id="preset-row-3" class="btn secondary">Preset Row x3</button>
        <button id="preset-col-3" class="btn secondary">Preset Col x3</button>
        <button id="preset-block-2x2" class="btn secondary">Preset Block 2x2</button>
        <button id="undo-last-edit" class="btn secondary" disabled>Undo</button>
        <button id="redo-last-edit" class="btn secondary" disabled>Redo</button>
      </div>
      <div class="row three">
        <label class="field compact">
          <span>Preset Rows</span>
          <input id="preset-rows" type="number" min="1" max="1048576" step="1" value="2" />
        </label>
        <label class="field compact">
          <span>Preset Cols</span>
          <input id="preset-cols" type="number" min="1" max="16384" step="1" value="2" />
        </label>
        <div class="field compact">
          <span>Apply Custom</span>
          <button id="preset-apply-custom" class="btn secondary">Apply N x M</button>
          <button id="preset-apply-last" class="btn secondary" disabled>Apply Last Custom</button>
        </div>
      </div>
      <p class="summary hint">
        Presets anchor at the selected preview cell when available; otherwise they use the current
        edit cell. Custom presets support configurable <code>N x M</code> ranges.
      </p>
      <div class="actions">
        <button id="apply-edit" class="btn primary">Apply Cell Edit</button>
      </div>
      <pre id="edit-output" class="output">No edits yet.</pre>
    </section>

    <section class="panel">
      <h2>Run Macro</h2>
      <label class="field">
        <span>Python Script Path</span>
        <input id="macro-script-path" placeholder="C:\\path\\to\\macro.py" />
      </label>
      <label class="field">
        <span>Macro Name</span>
        <input id="macro-name" value="main" />
      </label>
      <label class="field">
        <span>Macro Args (one per line, key=value)</span>
        <textarea id="macro-args" rows="4" placeholder="region=North\nrunId=monthly"></textarea>
      </label>
      <div class="permission-grid">
        <label class="toggle">
          <input id="macro-permission-fs-read" type="checkbox" />
          <span>fs.read</span>
        </label>
        <label class="toggle">
          <input id="macro-permission-fs-write" type="checkbox" />
          <span>fs.write</span>
        </label>
        <label class="toggle">
          <input id="macro-permission-net-http" type="checkbox" />
          <span>net.http</span>
        </label>
        <label class="toggle">
          <input id="macro-permission-clipboard" type="checkbox" />
          <span>clipboard</span>
        </label>
        <label class="toggle">
          <input id="macro-permission-process-exec" type="checkbox" />
          <span>process.exec</span>
        </label>
        <label class="toggle">
          <input id="macro-permission-udf" type="checkbox" />
          <span>udf</span>
        </label>
        <label class="toggle">
          <input id="macro-permission-events-emit" type="checkbox" />
          <span>events.emit</span>
        </label>
      </div>
      <div class="field">
        <label class="toggle">
          <input id="macro-remember-policy" type="checkbox" />
          <span>Remember this script permission policy</span>
        </label>
        <p id="macro-policy-status" class="summary hint">No macro policy loaded.</p>
      </div>
      <div class="actions">
        <button id="macro-clear-policy" class="btn secondary">Clear Saved Policy</button>
      </div>
      <div class="actions">
        <button id="run-macro" class="btn primary">Run Macro</button>
      </div>
      <pre id="macro-output" class="output">No macro runs yet.</pre>
    </section>

    <section class="panel" id="capture-section-lifecycle">
      <h2>Edit Lifecycle Telemetry</h2>
      <p class="summary hint">
        Command-level trace, latency, and error visibility for open/edit/paste/undo/redo/save/recalc flows.
      </p>
      <pre id="edit-lifecycle-output" class="output compact" role="status" aria-live="polite">No edit lifecycle events yet.</pre>
    </section>

    <section class="panel" id="capture-section-save-recalc">
      <h2>Save + Recalc</h2>
      <label class="field">
        <span>Output Path</span>
        <input id="save-path" placeholder="C:\\path\\to\\workbook.rootcellar.preserve.xlsx" />
      </label>
      <label class="toggle">
        <input id="save-promote-source" type="checkbox" />
        <span>Use saved output as current source for next preserve save</span>
      </label>
      <div class="row">
        <label class="field compact">
          <span>Save Mode</span>
          <select id="save-mode">
            <option value="preserve">Preserve (interop-first)</option>
            <option value="normalize">Normalize</option>
          </select>
        </label>
        <label class="field compact">
          <span>Recalc Sheet (Optional)</span>
          <input id="recalc-sheet" placeholder="Sheet1" />
        </label>
      </div>
      <div class="actions">
        <button id="pick-save-path" class="btn secondary">Choose Output...</button>
        <button id="save-workbook" class="btn primary">Save Workbook</button>
        <button id="recalc-loaded" class="btn secondary">Recalc Loaded Workbook</button>
      </div>
      <div class="freshness-row">
        <span class="freshness-label">Recalc Freshness</span>
        <span id="recalc-freshness-badge" class="freshness-badge pending">No Recalc Yet</span>
      </div>
      <pre id="save-output" class="output">No save/recalc run yet.</pre>
      <pre id="recalc-status" class="output compact" aria-live="polite">No recalc run yet for this workbook.</pre>
    </section>

    <section class="panel">
      <h2>App Status</h2>
      <pre id="status-output" class="output">Loading...</pre>
    </section>

    <section class="panel">
      <h2>Round-Trip Result</h2>
      <pre id="roundtrip-output" class="output">No run yet.</pre>
    </section>
  </main>
`;

const statusOutput = document.querySelector<HTMLPreElement>("#status-output");
const roundTripOutput = document.querySelector<HTMLPreElement>("#roundtrip-output");
const refreshButton = document.querySelector<HTMLButtonElement>("#refresh-status");
const runRoundTripButton = document.querySelector<HTMLButtonElement>("#run-round-trip");
const openPathInput = document.querySelector<HTMLInputElement>("#open-path");
const pickOpenPathButton = document.querySelector<HTMLButtonElement>("#pick-open-path");
const openWorkbookButton = document.querySelector<HTMLButtonElement>("#open-workbook");
const refreshSessionButton = document.querySelector<HTMLButtonElement>("#refresh-session");
const sessionOutput = document.querySelector<HTMLPreElement>("#session-output");
const compatOutput = document.querySelector<HTMLPreElement>("#compat-output");
const previewSheetSelect = document.querySelector<HTMLSelectElement>("#preview-sheet");
const previewLimitInput = document.querySelector<HTMLInputElement>("#preview-limit");
const refreshPreviewButton = document.querySelector<HTMLButtonElement>("#refresh-preview");
const jumpLastEditedButton = document.querySelector<HTMLButtonElement>("#jump-last-edited");
const copyPreviewA1Button = document.querySelector<HTMLButtonElement>("#copy-preview-a1");
const copyPreviewValueButton = document.querySelector<HTMLButtonElement>("#copy-preview-value");
const copyPreviewFormulaButton = document.querySelector<HTMLButtonElement>("#copy-preview-formula");
const pasteClipboardCellButton = document.querySelector<HTMLButtonElement>("#paste-clipboard-cell");
const formulaBarCellInput = document.querySelector<HTMLInputElement>("#formula-bar-cell");
const formulaBarModeSelect = document.querySelector<HTMLSelectElement>("#formula-bar-mode");
const formulaBarInput = document.querySelector<HTMLInputElement>("#formula-bar-input");
const formulaBarApplyButton = document.querySelector<HTMLButtonElement>("#formula-bar-apply");
const srAnnouncer = document.querySelector<HTMLDivElement>("#sr-announcer");
const srAnnouncerAssertive = document.querySelector<HTMLDivElement>("#sr-announcer-assertive");
const previewOutput = document.querySelector<HTMLPreElement>("#preview-output");
const previewSelectedOutput = document.querySelector<HTMLPreElement>("#preview-selected-output");
const previewCopyStatus = document.querySelector<HTMLPreElement>("#preview-copy-status");
const previewWrap = document.querySelector<HTMLDivElement>("#preview-wrap");
const previewTable = document.querySelector<HTMLTableElement>("#preview-table");
const editLifecycleOutput = document.querySelector<HTMLPreElement>("#edit-lifecycle-output");
const editSheetSelect = document.querySelector<HTMLSelectElement>("#edit-sheet");
const editCellInput = document.querySelector<HTMLInputElement>("#edit-cell");
const editModeSelect = document.querySelector<HTMLSelectElement>("#edit-mode");
const editInput = document.querySelector<HTMLInputElement>("#edit-input");
const presetRow3Button = document.querySelector<HTMLButtonElement>("#preset-row-3");
const presetCol3Button = document.querySelector<HTMLButtonElement>("#preset-col-3");
const presetBlock2x2Button = document.querySelector<HTMLButtonElement>("#preset-block-2x2");
const undoLastEditButton = document.querySelector<HTMLButtonElement>("#undo-last-edit");
const redoLastEditButton = document.querySelector<HTMLButtonElement>("#redo-last-edit");
const presetRowsInput = document.querySelector<HTMLInputElement>("#preset-rows");
const presetColsInput = document.querySelector<HTMLInputElement>("#preset-cols");
const presetApplyCustomButton = document.querySelector<HTMLButtonElement>("#preset-apply-custom");
const presetApplyLastButton = document.querySelector<HTMLButtonElement>("#preset-apply-last");
const applyEditButton = document.querySelector<HTMLButtonElement>("#apply-edit");
const editOutput = document.querySelector<HTMLPreElement>("#edit-output");
const macroScriptPathInput = document.querySelector<HTMLInputElement>("#macro-script-path");
const macroNameInput = document.querySelector<HTMLInputElement>("#macro-name");
const macroArgsInput = document.querySelector<HTMLTextAreaElement>("#macro-args");
const macroPermissionFsReadInput = document.querySelector<HTMLInputElement>("#macro-permission-fs-read");
const macroPermissionFsWriteInput = document.querySelector<HTMLInputElement>("#macro-permission-fs-write");
const macroPermissionNetHttpInput = document.querySelector<HTMLInputElement>("#macro-permission-net-http");
const macroPermissionClipboardInput = document.querySelector<HTMLInputElement>("#macro-permission-clipboard");
const macroPermissionProcessExecInput = document.querySelector<HTMLInputElement>("#macro-permission-process-exec");
const macroPermissionUdfInput = document.querySelector<HTMLInputElement>("#macro-permission-udf");
const macroPermissionEventsEmitInput = document.querySelector<HTMLInputElement>("#macro-permission-events-emit");
const macroRememberPolicyInput = document.querySelector<HTMLInputElement>("#macro-remember-policy");
const macroPolicyStatus = document.querySelector<HTMLElement>("#macro-policy-status");
const macroClearPolicyButton = document.querySelector<HTMLButtonElement>("#macro-clear-policy");
const runMacroButton = document.querySelector<HTMLButtonElement>("#run-macro");
const macroOutput = document.querySelector<HTMLPreElement>("#macro-output");
const savePathInput = document.querySelector<HTMLInputElement>("#save-path");
const savePromoteSourceInput = document.querySelector<HTMLInputElement>("#save-promote-source");
const pickSavePathButton = document.querySelector<HTMLButtonElement>("#pick-save-path");
const saveModeSelect = document.querySelector<HTMLSelectElement>("#save-mode");
const saveWorkbookButton = document.querySelector<HTMLButtonElement>("#save-workbook");
const recalcSheetInput = document.querySelector<HTMLInputElement>("#recalc-sheet");
const recalcLoadedButton = document.querySelector<HTMLButtonElement>("#recalc-loaded");
const recalcFreshnessBadge = document.querySelector<HTMLSpanElement>("#recalc-freshness-badge");
const saveOutput = document.querySelector<HTMLPreElement>("#save-output");
const recalcStatusOutput = document.querySelector<HTMLPreElement>("#recalc-status");

if (
  !statusOutput ||
  !roundTripOutput ||
  !refreshButton ||
  !runRoundTripButton ||
  !openPathInput ||
  !pickOpenPathButton ||
  !openWorkbookButton ||
  !refreshSessionButton ||
  !sessionOutput ||
  !compatOutput ||
  !previewSheetSelect ||
  !previewLimitInput ||
  !refreshPreviewButton ||
  !jumpLastEditedButton ||
  !copyPreviewA1Button ||
  !copyPreviewValueButton ||
  !copyPreviewFormulaButton ||
  !pasteClipboardCellButton ||
  !formulaBarCellInput ||
  !formulaBarModeSelect ||
  !formulaBarInput ||
  !formulaBarApplyButton ||
  !srAnnouncer ||
  !srAnnouncerAssertive ||
  !previewOutput ||
  !previewSelectedOutput ||
  !previewCopyStatus ||
  !previewWrap ||
  !previewTable ||
  !editLifecycleOutput ||
  !editSheetSelect ||
  !editCellInput ||
  !editModeSelect ||
  !editInput ||
  !presetRow3Button ||
  !presetCol3Button ||
  !presetBlock2x2Button ||
  !undoLastEditButton ||
  !redoLastEditButton ||
  !presetRowsInput ||
  !presetColsInput ||
  !presetApplyCustomButton ||
  !presetApplyLastButton ||
  !applyEditButton ||
  !editOutput ||
  !macroScriptPathInput ||
  !macroNameInput ||
  !macroArgsInput ||
  !macroPermissionFsReadInput ||
  !macroPermissionFsWriteInput ||
  !macroPermissionNetHttpInput ||
  !macroPermissionClipboardInput ||
  !macroPermissionProcessExecInput ||
  !macroPermissionUdfInput ||
  !macroPermissionEventsEmitInput ||
  !macroRememberPolicyInput ||
  !macroPolicyStatus ||
  !macroClearPolicyButton ||
  !runMacroButton ||
  !macroOutput ||
  !savePathInput ||
  !savePromoteSourceInput ||
  !pickSavePathButton ||
  !saveModeSelect ||
  !saveWorkbookButton ||
  !recalcSheetInput ||
  !recalcLoadedButton ||
  !recalcFreshnessBadge ||
  !saveOutput ||
  !recalcStatusOutput
) {
  throw new Error("missing required UI elements");
}

const statusOutputEl = statusOutput;
const roundTripOutputEl = roundTripOutput;
const runRoundTripButtonEl = runRoundTripButton;
const refreshButtonEl = refreshButton;
const openPathInputEl = openPathInput;
const pickOpenPathButtonEl = pickOpenPathButton;
const openWorkbookButtonEl = openWorkbookButton;
const refreshSessionButtonEl = refreshSessionButton;
const sessionOutputEl = sessionOutput;
const compatOutputEl = compatOutput;
const previewSheetSelectEl = previewSheetSelect;
const previewLimitInputEl = previewLimitInput;
const refreshPreviewButtonEl = refreshPreviewButton;
const jumpLastEditedButtonEl = jumpLastEditedButton;
const copyPreviewA1ButtonEl = copyPreviewA1Button;
const copyPreviewValueButtonEl = copyPreviewValueButton;
const copyPreviewFormulaButtonEl = copyPreviewFormulaButton;
const pasteClipboardCellButtonEl = pasteClipboardCellButton;
const formulaBarCellInputEl = formulaBarCellInput;
const formulaBarModeSelectEl = formulaBarModeSelect;
const formulaBarInputEl = formulaBarInput;
const formulaBarApplyButtonEl = formulaBarApplyButton;
const srAnnouncerEl = srAnnouncer;
const srAnnouncerAssertiveEl = srAnnouncerAssertive;
const previewOutputEl = previewOutput;
const previewSelectedOutputEl = previewSelectedOutput;
const previewCopyStatusEl = previewCopyStatus;
const previewWrapEl = previewWrap;
const previewTableEl = previewTable;
const editLifecycleOutputEl = editLifecycleOutput;
const editSheetSelectEl = editSheetSelect;
const editCellInputEl = editCellInput;
const editModeSelectEl = editModeSelect;
const editInputEl = editInput;
const presetRow3ButtonEl = presetRow3Button;
const presetCol3ButtonEl = presetCol3Button;
const presetBlock2x2ButtonEl = presetBlock2x2Button;
const undoLastEditButtonEl = undoLastEditButton;
const redoLastEditButtonEl = redoLastEditButton;
const presetRowsInputEl = presetRowsInput;
const presetColsInputEl = presetColsInput;
const presetApplyCustomButtonEl = presetApplyCustomButton;
const presetApplyLastButtonEl = presetApplyLastButton;
const applyEditButtonEl = applyEditButton;
const editOutputEl = editOutput;
const macroScriptPathInputEl = macroScriptPathInput;
const macroNameInputEl = macroNameInput;
const macroArgsInputEl = macroArgsInput;
const macroPermissionFsReadInputEl = macroPermissionFsReadInput;
const macroPermissionFsWriteInputEl = macroPermissionFsWriteInput;
const macroPermissionNetHttpInputEl = macroPermissionNetHttpInput;
const macroPermissionClipboardInputEl = macroPermissionClipboardInput;
const macroPermissionProcessExecInputEl = macroPermissionProcessExecInput;
const macroPermissionUdfInputEl = macroPermissionUdfInput;
const macroPermissionEventsEmitInputEl = macroPermissionEventsEmitInput;
const macroRememberPolicyInputEl = macroRememberPolicyInput;
const macroPolicyStatusEl = macroPolicyStatus;
const macroClearPolicyButtonEl = macroClearPolicyButton;
const runMacroButtonEl = runMacroButton;
const macroOutputEl = macroOutput;
const savePathInputEl = savePathInput;
const savePromoteSourceInputEl = savePromoteSourceInput;
const pickSavePathButtonEl = pickSavePathButton;
const saveModeSelectEl = saveModeSelect;
const saveWorkbookButtonEl = saveWorkbookButton;
const recalcSheetInputEl = recalcSheetInput;
const recalcLoadedButtonEl = recalcLoadedButton;
const recalcFreshnessBadgeEl = recalcFreshnessBadge;
const saveOutputEl = saveOutput;
const recalcStatusOutputEl = recalcStatusOutput;

let currentSession: InteropSessionStatusResponse | null = null;
let currentPreview: InteropSheetPreviewResponse | null = null;
let selectedPreviewCell: SelectedPreviewCell | null = null;
let lastEditedCell: { sheet: string; cell: string } | null = null;
let recalcStatusState: RecalcStatusState = resetRecalcStatusState();
let lastCustomPreset: PresetDimensions | null = null;

function updateHistoryControls(): void {
  const canUndo = currentSession?.loaded && currentSession.undoCount > 0;
  const canRedo = currentSession?.loaded && currentSession.redoCount > 0;
  undoLastEditButtonEl.disabled = !canUndo;
  redoLastEditButtonEl.disabled = !canRedo;
  runMacroButtonEl.disabled = !Boolean(currentSession?.loaded);
}

function applyMacroPermissionConfig(config: InteropMacroPermissionConfig): void {
  macroPermissionFsReadInputEl.checked = config.fsRead;
  macroPermissionFsWriteInputEl.checked = config.fsWrite;
  macroPermissionNetHttpInputEl.checked = config.netHttp;
  macroPermissionClipboardInputEl.checked = config.clipboard;
  macroPermissionProcessExecInputEl.checked = config.processExec;
  macroPermissionUdfInputEl.checked = config.udf;
  macroPermissionEventsEmitInputEl.checked = config.eventsEmit;
}

function collectMacroPermissionLabel(config: InteropMacroPermissionConfig): string {
  const enabled = [];
  if (config.fsRead) {
    enabled.push("fs.read");
  }
  if (config.fsWrite) {
    enabled.push("fs.write");
  }
  if (config.netHttp) {
    enabled.push("net.http");
  }
  if (config.clipboard) {
    enabled.push("clipboard");
  }
  if (config.processExec) {
    enabled.push("process.exec");
  }
  if (config.udf) {
    enabled.push("udf");
  }
  if (config.eventsEmit) {
    enabled.push("events.emit");
  }
  return enabled.length > 0 ? enabled.join(", ") : "none";
}

function renderMacroPolicyStatus(
  source: MacroPermissionPolicySource,
  config: InteropMacroPermissionConfig,
): void {
  const enabledSummary = collectMacroPermissionLabel(config);
  const policyText =
    source === "stored"
      ? "Using saved policy for this script path."
      : "No saved policy found for this script path.";
  const riskHint = hasElevatedMacroPermissions(config)
    ? " Elevated permissions enabled. Trust prompt required before first run unless policy is stored."
    : "";
  macroPolicyStatusEl.textContent = `${policyText} Enabled permissions: ${enabledSummary}.${riskHint}`;
}

function resolveMacroPolicySource(
  scriptPath: string,
): MacroPermissionPolicySource {
  const stored = loadMacroPermissionPolicyMetadata(scriptPath);
  if (!stored) {
    return "fresh";
  }

  const currentConfig = collectMacroPermissionConfig();
  return permissionsEqual(stored.permissions, currentConfig) ? "stored" : "fresh";
}

function applySavedMacroPermissionPolicy(scriptPath: string): void {
  const stored = loadMacroPermissionPolicy(scriptPath);
  if (stored) {
    applyMacroPermissionConfig(stored);
    macroRememberPolicyInputEl.checked = true;
    renderMacroPolicyStatus("stored", stored);
    return;
  }

  renderMacroPolicyStatus("fresh", collectMacroPermissionConfig());
  macroRememberPolicyInputEl.checked = false;
}

function normalizeMacroScriptPath(): string {
  return macroScriptPathInputEl.value.trim();
}

function formatMacroPolicyPrompt(
  scriptPath: string,
  config: InteropMacroPermissionConfig,
): string {
  const permissions = collectMacroPermissionLabel(config);
  const shortPath = scriptPath.trim() || "this macro";
  return `Allow ${shortPath} to use permissions: ${permissions}?`;
}

function fillSheetSelect(select: HTMLSelectElement, sheets: string[]): void {
  const currentValue = select.value;
  select.innerHTML = "";

  if (sheets.length === 0) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "No sheets loaded";
    select.appendChild(option);
    select.disabled = true;
    return;
  }

  select.disabled = false;
  for (const sheet of sheets) {
    const option = document.createElement("option");
    option.value = sheet;
    option.textContent = sheet;
    select.appendChild(option);
  }

  if (sheets.includes(currentValue)) {
    select.value = currentValue;
  } else {
    select.value = sheets[0];
  }
}

function setSheetOptions(sheets: string[]): void {
  fillSheetSelect(editSheetSelectEl, sheets);
  fillSheetSelect(previewSheetSelectEl, sheets);
}

function normalizeA1(value: string): string {
  return value.trim().toUpperCase();
}

function updatePreviewCopyButtons(): void {
  const hasSelection = Boolean(selectedPreviewCell);
  copyPreviewA1ButtonEl.disabled = !hasSelection;
  copyPreviewValueButtonEl.disabled = !hasSelection;
  copyPreviewFormulaButtonEl.disabled =
    !hasSelection || selectedPreviewCell?.cell.formula == null;
  pasteClipboardCellButtonEl.disabled = !hasSelection;
  formulaBarApplyButtonEl.disabled = !hasSelection;
}

function cellValueToInputValue(value: CellValue): string {
  if (value === "empty") {
    return "";
  }
  if ("text" in value) {
    return value.text;
  }
  if ("number" in value) {
    return String(value.number);
  }
  if ("bool" in value) {
    return value.bool ? "true" : "false";
  }
  if ("error" in value) {
    return value.error;
  }
  return "";
}

function populateEditFormFromSelection(selection: SelectedPreviewCell): void {
  if (!editSheetSelectEl.disabled) {
    editSheetSelectEl.value = selection.sheet;
  }
  editCellInputEl.value = selection.cell.cell;
  if (selection.cell.formula) {
    editModeSelectEl.value = "formula";
    editInputEl.value = selection.cell.formula;
  } else {
    editModeSelectEl.value = "value";
    editInputEl.value = cellValueToInputValue(selection.cell.value);
  }
}

function populateFormulaBarFromSelection(selection: SelectedPreviewCell): void {
  formulaBarCellInputEl.value = `${selection.sheet}!${selection.cell.cell}`;
  if (selection.cell.formula) {
    formulaBarModeSelectEl.value = "formula";
    formulaBarInputEl.value = selection.cell.formula;
  } else {
    formulaBarModeSelectEl.value = "value";
    formulaBarInputEl.value = cellValueToInputValue(selection.cell.value);
  }
}

function syncEditFormFromFormulaBarSelectionInput(): void {
  if (!selectedPreviewCell) {
    return;
  }
  if (!editSheetSelectEl.disabled) {
    editSheetSelectEl.value = selectedPreviewCell.sheet;
  }
  editCellInputEl.value = selectedPreviewCell.cell.cell;
  editModeSelectEl.value = normalizeEditMode(formulaBarModeSelectEl.value);
  editInputEl.value = formulaBarInputEl.value;
}

function announcePreviewSelection(selection: SelectedPreviewCell | null): void {
  if (!selection) {
    announceToScreenReader("No preview cell selected.");
    return;
  }

  const detail = `${selection.sheet}!${selection.cell.cell}`;
  const formatHint = selection.cell.formula
    ? `formula ${selection.cell.formula}`
    : `value ${cellValueToInputValue(selection.cell.value)}`;
  announceToScreenReader(`Selected ${detail}. ${formatHint}`);
}

function setSelectedPreviewCell(selection: SelectedPreviewCell | null): void {
  selectedPreviewCell = selection;
  updatePreviewCopyButtons();

  if (!selection) {
    previewSelectedOutputEl.textContent = "No cell selected.";
    formulaBarCellInputEl.value = "-";
    formulaBarInputEl.value = "";
    return;
  }

  populateEditFormFromSelection(selection);
  populateFormulaBarFromSelection(selection);

  const detail = {
    sheet: selection.sheet,
    cell: selection.cell.cell,
    value: formatCellValue(selection.cell.value),
    formula: selection.cell.formula,
  };
  previewSelectedOutputEl.textContent = JSON.stringify(detail, null, 2);
  announcePreviewSelection(selection);
}

function updateJumpLastEditedButton(): void {
  jumpLastEditedButtonEl.disabled = !lastEditedCell;
}

function renderRecalcStatus(): void {
  renderRecalcStatusView({
    state: recalcStatusState,
    badgeEl: recalcFreshnessBadgeEl,
    statusEl: recalcStatusOutputEl,
  });
}

function resetRecalcStatus(): void {
  recalcStatusState = resetRecalcStatusState();
  renderRecalcStatus();
}

function markRecalcPending(): void {
  recalcStatusState = markRecalcPendingState(recalcStatusState);
  renderRecalcStatus();
}

function markRecalcCompleted(scope: string | null): void {
  recalcStatusState = markRecalcCompletedState(recalcStatusState, scope, new Date().toISOString());
  renderRecalcStatus();
}

function loadLastCustomPresetFromStorage(): PresetDimensions | null {
  try {
    return parsePresetDimensions(
      localStorage.getItem(PRESET_DIMENSIONS_STORAGE_KEY),
      maxPresetRows,
      maxPresetCols,
    );
  } catch {
    return null;
  }
}

function persistLastCustomPreset(dimensions: PresetDimensions): void {
  try {
    localStorage.setItem(
      PRESET_DIMENSIONS_STORAGE_KEY,
      serializePresetDimensions(dimensions),
    );
  } catch {
    // Persistence is a convenience feature only.
  }
}

function renderLastCustomPresetButton(): void {
  renderLastCustomPresetButtonView({
    buttonEl: presetApplyLastButtonEl,
    dimensions: lastCustomPreset,
  });
}

function seedUiCaptureDemo(): void {
  const demoInputPath = "C:\\Demo\\quarterly-forecast.xlsx";
  const demoSheet = "Summary";
  const baseTrace: TraceEcho = {
    traceId: "ui-capture",
    spanId: "ui-capture",
    parentSpanId: null,
    sessionId: "ui-capture",
  };
  const editTrace: TraceEcho = {
    traceId: "ui-capture-edit",
    spanId: "ui-capture-edit",
    parentSpanId: "ui-capture",
    sessionId: "ui-capture",
  };
  const saveTrace: TraceEcho = {
    traceId: "ui-capture-save",
    spanId: "ui-capture-save",
    parentSpanId: "ui-capture",
    sessionId: "ui-capture",
  };
  const recalcTrace: TraceEcho = {
    traceId: "ui-capture-recalc",
    spanId: "ui-capture-recalc",
    parentSpanId: "ui-capture",
    sessionId: "ui-capture",
  };

  const statusPayload: InteropSessionStatusResponse = {
    loaded: true,
    inputPath: demoInputPath,
    workbookId: "ui-capture-demo-workbook",
    sheetCount: 2,
    cellCount: 10,
    issueCount: 1,
    unknownPartCount: 2,
    dirtySheetCount: 1,
    dirtySheets: [demoSheet],
    undoCount: 0,
    redoCount: 0,
    sheets: [demoSheet, "Detail"],
  };
  renderSessionStatus(statusPayload);

  openPathInputEl.value = demoInputPath;
  savePathInputEl.value = suggestOutputPath(demoInputPath, normalizeMode(saveModeSelectEl.value));
  compatOutputEl.textContent = jsonWithTraceHeader(baseTrace, [
    "Feature score: 92",
    "Issues: 1",
    "Unknown parts: 2",
    "Part graph: nodes=41, edges=62, dangling=0, external=2",
    "",
    "Compatibility Issues:",
    "- [partially_supported] xl.comments: threaded comments preserved in preserve mode",
  ].join("\n"));

  const previewCells: InteropPreviewCell[] = [
    { cell: "A1", row: 1, col: 1, value: { text: "Region" }, formula: null },
    { cell: "B1", row: 1, col: 2, value: { text: "Actual" }, formula: null },
    { cell: "C1", row: 1, col: 3, value: { text: "Forecast" }, formula: null },
    { cell: "A2", row: 2, col: 1, value: { text: "North" }, formula: null },
    { cell: "B2", row: 2, col: 2, value: { number: 1250 }, formula: null },
    { cell: "C2", row: 2, col: 3, value: { number: 1330 }, formula: "=B2*1.064" },
    { cell: "A3", row: 3, col: 1, value: { text: "South" }, formula: null },
    { cell: "B3", row: 3, col: 2, value: { number: 980 }, formula: null },
    { cell: "C3", row: 3, col: 3, value: { number: 1060 }, formula: "=B3*1.082" },
    { cell: "D3", row: 3, col: 4, value: { bool: true }, formula: null },
  ];

  currentPreview = {
    workbookId: "ui-capture-demo-workbook",
    sheet: demoSheet,
    totalCells: previewCells.length,
    shownCells: previewCells.length,
    truncated: false,
    cells: previewCells,
    trace: baseTrace,
  };
  previewSheetSelectEl.value = demoSheet;
  lastEditedCell = { sheet: demoSheet, cell: "C3" };
  updateJumpLastEditedButton();

  const selected = previewCells.find((cell) => cell.cell === "C3") ?? previewCells[0];
  setSelectedPreviewCell({ sheet: demoSheet, cell: selected });
  renderPreviewTable(currentPreview);
  previewOutputEl.textContent = [
    `Latest trace: ${formatTraceHeader(baseTrace)}`,
    `Workbook: ${currentPreview.workbookId}`,
    `Sheet: ${currentPreview.sheet}`,
    `Shown cells: ${currentPreview.shownCells}/${currentPreview.totalCells}`,
    "Truncated: no",
  ].join("\n");
  previewCopyStatusEl.textContent = `UI capture mode (${uiCaptureRecalcState}) seeded at ${demoSheet}!${selected.cell}.`;
  saveOutputEl.textContent = "UI capture mode: save/recalc outputs are simulated for visual review.";
  lastCustomPreset = { rows: 3, cols: 4 };
  presetRowsInputEl.value = String(lastCustomPreset.rows);
  presetColsInputEl.value = String(lastCustomPreset.cols);
  renderLastCustomPresetButton();

  const recalcIso = "2026-03-02T20:00:00.000Z";
  recalcStatusState = resetRecalcStatusState();
  if (uiCaptureRecalcState === "pending") {
    recalcStatusState = markRecalcPendingState(recalcStatusState);
  } else if (uiCaptureRecalcState === "fresh") {
    recalcStatusState = markRecalcCompletedState(recalcStatusState, demoSheet, recalcIso);
  } else {
    recalcStatusState = markRecalcCompletedState(recalcStatusState, demoSheet, recalcIso);
    recalcStatusState = markRecalcPendingState(recalcStatusState);
  }
  renderRecalcStatus();

  if (uiCaptureSection === "edit-cell") {
    editSheetSelectEl.value = demoSheet;
    editCellInputEl.value = "C3";
    editModeSelectEl.value = "formula";
    editInputEl.value = "=B3*1.082";
    const editDemoOutput = {
      workbookId: "ui-capture-demo-workbook",
      sheet: demoSheet,
      cell: "C3",
      anchorCell: "C3",
      appliedCellCount: 1,
      mode: "formula",
      value: "number=1060",
      formula: "=B3*1.082",
      dirtySheetCount: 1,
      dirtySheets: [demoSheet],
      trace: editTrace,
    };
    editOutputEl.textContent = jsonWithTraceHeader(editTrace, JSON.stringify(editDemoOutput, null, 2));
    previewCopyStatusEl.textContent = "UI capture mode: edit-cell output prepared for close-up review.";
  } else if (uiCaptureSection === "save-recalc") {
    recalcSheetInputEl.value = demoSheet;
    savePromoteSourceInputEl.checked = true;
    const saveDemoOutput = {
      inputPath: demoInputPath,
      outputPath: suggestOutputPath(demoInputPath, "preserve"),
      workbookId: "ui-capture-demo-workbook",
      mode: "preserve",
      sheetCount: 2,
      cellCount: 10,
      copiedBytes: 1280,
      partGraph: {
        nodeCount: 41,
        edgeCount: 62,
        danglingEdgeCount: 0,
        externalEdgeCount: 2,
        unknownPartCount: 2,
      },
      partGraphFlags: {
        strategy: "interop_preserve",
        sourceGraphReused: true,
        relationshipsPreserved: true,
        unknownPartsPreserved: true,
      },
      trace: saveTrace,
    };
    const recalcDemoOutput = {
      workbookId: "ui-capture-demo-workbook",
      reports: [
        {
          sheet: demoSheet,
          evaluatedCells: 3,
          cycleCount: 0,
          parseErrorCount: 0,
        },
      ],
      trace: recalcTrace,
    };
    const saveSectionLines = [
      "Latest save command:",
      jsonWithTraceHeader(saveTrace, JSON.stringify(saveDemoOutput, null, 2)),
      "",
      "Latest recalc command:",
      jsonWithTraceHeader(recalcTrace, JSON.stringify(recalcDemoOutput, null, 2)),
    ];
    saveOutputEl.textContent = saveSectionLines.join("\n");
  } else {
    editOutputEl.textContent = "UI capture mode: default seeded scenario for full-shell review.";
    macroScriptPathInputEl.value = "C:\\Macros\\quarterly_adjustments.py";
    macroNameInputEl.value = "adjust_forecast";
    macroArgsInputEl.value = [
      "region=North",
      "quarter=Q1",
      "factor=1.08",
    ].join("\n");
    macroPermissionFsReadInputEl.checked = true;
    macroPermissionFsWriteInputEl.checked = false;
    macroPermissionNetHttpInputEl.checked = false;
    macroPermissionClipboardInputEl.checked = false;
    macroPermissionProcessExecInputEl.checked = false;
    macroPermissionUdfInputEl.checked = false;
    macroPermissionEventsEmitInputEl.checked = false;
    macroRememberPolicyInputEl.checked = false;
    renderMacroPolicyStatus("fresh", collectMacroPermissionConfig());
    macroOutputEl.textContent = jsonWithTraceHeader({
      traceId: "ui-capture-macro",
      spanId: "ui-capture-macro",
      parentSpanId: "ui-capture",
      sessionId: "ui-capture",
    }, JSON.stringify({
      scriptPath: "C:\\Macros\\quarterly_adjustments.py",
      macroName: "adjust_forecast",
      requestedPermissions: ["fs.read", "fs.write"],
      permissionEvents: [],
      permissionGranted: 0,
      permissionDenied: 0,
      mutationCount: 0,
      changedSheets: ["Summary"],
      mutations: [
        {
          sheet: "Summary",
          cell: "C3",
          kind: "formula",
          value: { number: 1330 },
        },
      ],
      recalcReports: [
        {
          sheet: "Summary",
          evaluatedCells: 3,
          cycleCount: 0,
          parseErrorCount: 0,
        },
      ],
      stdout: "macro preview output",
      stderr: "",
    }, null, 2), "ui-capture");
  }

  focusCaptureSection(uiCaptureSection);
}

function applyRangePreset(rowSpan: number, colSpan: number, label: string): boolean {
  if (!currentSession?.loaded) {
    editOutputEl.textContent = "open a workbook first";
    return false;
  }

  const anchorSheet = selectedPreviewCell?.sheet ?? editSheetSelectEl.value.trim();
  const anchorA1 = selectedPreviewCell?.cell.cell ?? editCellInputEl.value.trim();
  if (!anchorSheet) {
    editOutputEl.textContent = "select a sheet before applying a preset";
    return false;
  }

  const rangeResult = buildPresetRange(anchorA1, rowSpan, colSpan);
  if (!rangeResult.ok) {
    if (rangeResult.reason === "invalid_anchor") {
      editOutputEl.textContent = "preset anchor must be a single A1 cell (not a range)";
    } else if (rangeResult.reason === "invalid_span") {
      editOutputEl.textContent = "preset dimensions must be whole numbers greater than zero";
    } else {
      editOutputEl.textContent = "preset range exceeds Excel bounds";
    }
    return false;
  }

  if (!editSheetSelectEl.disabled) {
    editSheetSelectEl.value = anchorSheet;
  }
  editCellInputEl.value = rangeResult.range;
  previewCopyStatusEl.textContent = `Preset ${label} selected ${anchorSheet}!${rangeResult.range}.`;
  editOutputEl.textContent = `Preset ${label} applied to range ${anchorSheet}!${rangeResult.range}.`;
  return true;
}

function parsePresetSpanInput(
  raw: string,
  min: number,
  max: number,
  fieldLabel: string,
): number | null {
  const parsed = Number.parseInt(raw.trim(), 10);
  if (!Number.isInteger(parsed) || parsed < min || parsed > max) {
    editOutputEl.textContent = `${fieldLabel} must be an integer from ${min} to ${max}.`;
    return null;
  }
  return parsed;
}

function scrollPreviewCellIntoView(a1Cell: string): void {
  const target = previewTableEl.querySelector<HTMLTableCellElement>(`td[data-cell="${a1Cell}"]`);
  if (target) {
    target.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
  }
}

function movePreviewSelection(direction: "up" | "down" | "left" | "right"): void {
  if (!currentPreview || currentPreview.cells.length === 0) {
    return;
  }

  const targetCell = computeNextPreviewSelection(
    currentPreview.cells,
    selectedPreviewCell?.sheet === currentPreview.sheet ? selectedPreviewCell.cell : null,
    direction,
  );
  if (!targetCell) {
    return;
  }

  setSelectedPreviewCell({ sheet: currentPreview.sheet, cell: targetCell });
  renderPreviewTable(currentPreview);
  scrollPreviewCellIntoView(targetCell.cell);
  previewCopyStatusEl.textContent = `Selected ${currentPreview.sheet}!${targetCell.cell}`;
}

function clearPreviewTable(message: string): void {
  previewTableEl.innerHTML = `<tbody><tr><td>${escapeHtml(message)}</td></tr></tbody>`;
}

function renderPreviewTable(payload: InteropSheetPreviewResponse): void {
  if (payload.cells.length === 0) {
    clearPreviewTable("No populated cells in this sheet.");
    return;
  }

  const rows = Array.from(new Set(payload.cells.map((cell) => cell.row))).sort((a, b) => a - b);
  const cols = Array.from(new Set(payload.cells.map((cell) => cell.col))).sort((a, b) => a - b);
  const cellMap = new Map(payload.cells.map((cell) => [`${cell.row}:${cell.col}`, cell]));

  const thead = [
    "<thead><tr><th class=\"corner\" role=\"columnheader\">#</th>",
    ...cols.map((col) => `<th role=\"columnheader\">${columnToA1(col)}</th>`),
    "</tr></thead>",
  ].join("");

  const tbodyRows = rows.map((row) => {
    const cells = cols.map((col) => {
      const key = `${row}:${col}`;
      const cell = cellMap.get(key);
      if (!cell) {
        return "<td></td>";
      }
      const isSelected =
        selectedPreviewCell !== null &&
        selectedPreviewCell.sheet === payload.sheet &&
        selectedPreviewCell.cell.cell === cell.cell;
      const classes = ["preview-cell"];
      if (isSelected) {
        classes.push("selected");
      }
      if (
        lastEditedCell &&
        lastEditedCell.sheet === payload.sheet &&
        lastEditedCell.cell === cell.cell
      ) {
        classes.push("last-edited");
      }

      return `<td role="gridcell" aria-selected="${isSelected ? "true" : "false"}" tabindex="${isSelected ? "0" : "-1"}" class="${classes.join(" ")}" data-cell="${escapeHtml(cell.cell)}" data-row="${cell.row}" data-col="${cell.col}" title="${escapeHtml(cell.cell)}">${escapeHtml(formatPreviewCell(cell))}</td>`;
    });
    return `<tr><th role="rowheader">${row}</th>${cells.join("")}</tr>`;
  });

  previewTableEl.innerHTML = `${thead}<tbody>${tbodyRows.join("")}</tbody>`;
}

async function loadAppStatus(): Promise<void> {
  statusOutputEl.textContent = "Loading status...";
  try {
    const payload = await invoke<AppStatusResponse>("app_status");
    statusOutputEl.textContent = JSON.stringify(payload, null, 2);
  } catch (error) {
    statusOutputEl.textContent = `status error: ${String(error)}`;
  }
}

function renderSessionStatus(payload: InteropSessionStatusResponse): void {
  currentSession = payload;
  setSheetOptions(payload.sheets);
  sessionOutputEl.textContent = JSON.stringify(payload, null, 2);
  updateHistoryControls();

  if (!payload.loaded) {
    currentPreview = null;
    lastEditedCell = null;
    resetRecalcStatus();
    updateJumpLastEditedButton();
    setSelectedPreviewCell(null);
    compatOutputEl.textContent = "No workbook loaded.";
    previewOutputEl.textContent = "No preview loaded.";
    previewCopyStatusEl.textContent = "No copy action yet.";
    clearPreviewTable("Open a workbook to render sheet preview.");
    updateHistoryControls();
    return;
  }

  if (!recalcSheetInputEl.value.trim() && payload.sheets.length > 0) {
    recalcSheetInputEl.placeholder = payload.sheets[0];
  }
}

async function loadInteropSessionStatus(): Promise<void> {
  sessionOutputEl.textContent = "Loading interop session...";
  try {
    const payload = await invoke<InteropSessionStatusResponse>("interop_session_status");
    renderSessionStatus(payload);
  } catch (error) {
    sessionOutputEl.textContent = `interop session error: ${String(error)}`;
  }
}

async function loadSheetPreview(
  sheetOverride?: string,
  focusCell?: string,
  scrollToFocus = false,
  traceContext?: TraceInput | TraceEcho,
): Promise<void> {
  if (!currentSession?.loaded) {
    currentPreview = null;
    setSelectedPreviewCell(null);
    previewOutputEl.textContent = "No preview loaded.";
    previewCopyStatusEl.textContent = "No copy action yet.";
    clearPreviewTable("Open a workbook to render sheet preview.");
    return;
  }

  const limitRaw = Number.parseInt(previewLimitInputEl.value.trim(), 10);
  const limit = Number.isFinite(limitRaw) ? Math.min(Math.max(limitRaw, 1), 400) : 120;
  const selectedSheet = (sheetOverride ?? previewSheetSelectEl.value).trim();
  const commandContext = startUiCommand("interop_sheet_preview");
  const trace = toTraceInput(traceContext ?? commandContext);

  refreshPreviewButtonEl.disabled = true;
  previewOutputEl.textContent = "Loading sheet preview...";
  logEditLifecycleEvent(
    "interop_sheet_preview",
    "start",
    `loading preview for ${selectedSheet || "current sheet"}`,
    { trace },
  );
  try {
    const payload = await invoke<InteropSheetPreviewResponse>("interop_sheet_preview", {
      sheet: selectedSheet.length > 0 ? selectedSheet : null,
      limit,
      trace,
    });

    if (payload.sheet !== previewSheetSelectEl.value && currentSession.sheets.includes(payload.sheet)) {
      previewSheetSelectEl.value = payload.sheet;
    }
    currentPreview = payload;

    const normalizedFocusCell = focusCell ? normalizeA1(focusCell) : null;
    if (normalizedFocusCell) {
      const focused = payload.cells.find((cell) => cell.cell === normalizedFocusCell) ?? null;
      if (focused) {
        setSelectedPreviewCell({ sheet: payload.sheet, cell: focused });
      } else {
        previewCopyStatusEl.textContent = payload.truncated
          ? `Cell ${normalizedFocusCell} is outside current preview window. Increase cell limit.`
          : `Cell ${normalizedFocusCell} not found in this sheet preview.`;
      }
    } else {
      const selected = selectedPreviewCell;
      if (
        selected &&
        (selected.sheet !== payload.sheet ||
          !payload.cells.some((cell) => cell.cell === selected.cell.cell))
      ) {
        setSelectedPreviewCell(null);
      }
    }

    const summary = [
      `Latest trace: ${formatTraceHeader(payload.trace)}`,
      `Workbook: ${payload.workbookId}`,
      `Sheet: ${payload.sheet}`,
      `Shown cells: ${payload.shownCells}/${payload.totalCells}`,
      `Truncated: ${payload.truncated ? "yes" : "no"}`,
    ];
    previewOutputEl.textContent = summary.join("\n");
    renderPreviewTable(payload);
    logEditLifecycleEvent("interop_sheet_preview", "success", `loaded sheet ${payload.sheet}`, {
      trace: payload.trace,
    });
    if (normalizedFocusCell && scrollToFocus) {
      const focusEl = previewTableEl.querySelector<HTMLTableCellElement>(
        `td[data-cell="${normalizedFocusCell}"]`,
      );
      if (focusEl) {
        focusEl.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
      }
    }
  } catch (error) {
    currentPreview = null;
    setSelectedPreviewCell(null);
    previewOutputEl.textContent = `preview error: ${String(error)}`;
    previewCopyStatusEl.textContent = "Preview unavailable.";
    logEditLifecycleEvent("interop_sheet_preview", "error", `preview load failed: ${String(error)}`, {
      trace,
      error,
    });
    clearPreviewTable("Preview unavailable.");
  } finally {
    refreshPreviewButtonEl.disabled = false;
  }
}

async function copyTextToClipboard(label: string, text: string): Promise<void> {
  try {
    await navigator.clipboard.writeText(text);
    previewCopyStatusEl.textContent = `Copied ${label}: ${text}`;
  } catch (error) {
    previewCopyStatusEl.textContent = `copy failed: ${String(error)}`;
  }
}

async function copySelectedA1(): Promise<void> {
  if (!selectedPreviewCell) {
    previewCopyStatusEl.textContent = "Select a preview cell first.";
    return;
  }
  await copyTextToClipboard("A1", selectedPreviewCell.cell.cell);
}

async function copySelectedValue(): Promise<void> {
  if (!selectedPreviewCell) {
    previewCopyStatusEl.textContent = "Select a preview cell first.";
    return;
  }
  await copyTextToClipboard("value", formatPreviewCell(selectedPreviewCell.cell));
}

async function copySelectedFormula(): Promise<void> {
  if (!selectedPreviewCell) {
    previewCopyStatusEl.textContent = "Select a preview cell first.";
    return;
  }
  if (!selectedPreviewCell.cell.formula) {
    previewCopyStatusEl.textContent = "Selected cell does not contain a formula.";
    return;
  }
  await copyTextToClipboard("formula", selectedPreviewCell.cell.formula);
}

function detectClipboardMode(raw: string): InteropEditMode {
  return raw.trim().startsWith("=") ? "formula" : "value";
}

async function pasteFromClipboardIntoSelection(): Promise<void> {
  if (!selectedPreviewCell) {
    previewCopyStatusEl.textContent = "Select a preview cell first.";
    logEditLifecycleEvent("interop_apply_cell_edit", "error", "paste failed: no selection");
    return;
  }
  logEditLifecycleEvent(
    "interop_apply_cell_edit",
    "start",
    `pasting into ${selectedPreviewCell.sheet}!${selectedPreviewCell.cell.cell}`,
  );

  try {
    const clipboardText = await navigator.clipboard.readText();
    const trimmed = clipboardText.trim();
    if (!trimmed) {
      previewCopyStatusEl.textContent = "Clipboard is empty.";
      logEditLifecycleEvent(
        "interop_apply_cell_edit",
        "error",
        `paste failed: clipboard empty in ${selectedPreviewCell.sheet}!${selectedPreviewCell.cell.cell}`,
      );
      return;
    }

    await applyCellEdit({
      sheet: selectedPreviewCell.sheet,
      cell: selectedPreviewCell.cell.cell,
      input: clipboardText,
      mode: detectClipboardMode(clipboardText),
      source: "edit-form",
    });
  } catch (error) {
    previewCopyStatusEl.textContent = `paste failed: ${String(error)}`;
    logEditLifecycleEvent("interop_apply_cell_edit", "error", `paste failed: ${String(error)}`, {
      error,
    });
  }
}

async function jumpToLastEdited(): Promise<void> {
  if (!lastEditedCell) {
    previewCopyStatusEl.textContent = "No edited cell recorded yet.";
    return;
  }

  previewSheetSelectEl.value = lastEditedCell.sheet;
  await loadSheetPreview(lastEditedCell.sheet, lastEditedCell.cell, true);
}

async function pickOpenPath(): Promise<void> {
  try {
    const selected = await openDialog({
      title: "Open Excel Workbook",
      multiple: false,
      filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
    });
    if (typeof selected !== "string") {
      return;
    }
    openPathInputEl.value = selected;

    if (!savePathInputEl.value.trim()) {
      savePathInputEl.value = suggestOutputPath(selected, normalizeMode(saveModeSelectEl.value));
    }
  } catch (error) {
    sessionOutputEl.textContent = `open dialog error: ${String(error)}`;
  }
}

async function openWorkbook(): Promise<void> {
  const path = openPathInputEl.value.trim();
  if (!path) {
    sessionOutputEl.textContent = "enter a workbook path first";
    logEditLifecycleEvent("interop_open_workbook", "error", "open failed: workbook path missing");
    return;
  }

  openWorkbookButtonEl.disabled = true;
  sessionOutputEl.textContent = "Opening workbook...";
  compatOutputEl.textContent = "Inspecting compatibility...";
  const commandContext = startUiCommand("interop_open_workbook");
  logEditLifecycleEvent("interop_open_workbook", "start", `open requested for ${path}`, {
    trace: commandContext,
  });

  try {
    const payload = await invoke<InteropOpenResponse>("interop_open_workbook", {
      path,
      trace: commandContext,
    });

    const statusPayload: InteropSessionStatusResponse = {
      loaded: true,
      inputPath: payload.inputPath,
      workbookId: payload.workbookId,
      sheetCount: payload.sheetCount,
      cellCount: payload.cellCount,
      issueCount: payload.issueCount,
      unknownPartCount: payload.unknownPartCount,
      dirtySheetCount: 0,
      dirtySheets: [],
      undoCount: 0,
      redoCount: 0,
      sheets: payload.sheets,
    };
    renderSessionStatus(statusPayload);
    const sessionView = {
      ...statusPayload,
      trace: payload.trace,
    };
    sessionOutputEl.textContent = jsonWithTraceHeader(
      payload.trace,
      JSON.stringify(sessionView, null, 2),
      commandContext.commandId,
    );
    compatOutputEl.textContent = jsonWithTraceHeader(
      payload.trace,
      formatCompatibility(payload),
      commandContext.commandId,
    );
    logEditLifecycleEvent("interop_open_workbook", "success", `opened ${payload.inputPath}`, {
      trace: payload.trace,
    });
    lastEditedCell = null;
    resetRecalcStatus();
    updateJumpLastEditedButton();
    setSelectedPreviewCell(null);
    previewCopyStatusEl.textContent = "No copy action yet.";
    savePathInputEl.value = suggestOutputPath(payload.inputPath, normalizeMode(saveModeSelectEl.value));
    await loadSheetPreview(payload.sheets[0], undefined, false, payload.trace);
  } catch (error) {
    const message = `open workbook error: ${String(error)}`;
    sessionOutputEl.textContent = message;
    compatOutputEl.textContent = message;
    previewCopyStatusEl.textContent = message;
    previewOutputEl.textContent = message;
    logEditLifecycleEvent("interop_open_workbook", "error", `open failed: ${String(error)}`, {
      trace: commandContext,
      error,
    });
    clearPreviewTable("Preview unavailable.");
  } finally {
    openWorkbookButtonEl.disabled = false;
  }
}

async function applyCellEdit(overrides?: {
  sheet?: string;
  cell?: string;
  input?: string;
  mode?: InteropEditMode;
  source?: EditActionSource;
}): Promise<void> {
  if (!currentSession?.loaded) {
    editOutputEl.textContent = "open a workbook first";
    logEditLifecycleEvent("interop_apply_cell_edit", "error", "edit failed: open a workbook first");
    return;
  }

  const source = overrides?.source ?? "edit-form";
  const sheet = (overrides?.sheet ?? editSheetSelectEl.value).trim();
  const cell = (overrides?.cell ?? editCellInputEl.value).trim();
  const input = overrides?.input ?? editInputEl.value;
  const mode = overrides?.mode ?? normalizeEditMode(editModeSelectEl.value);

  if (!sheet) {
    editOutputEl.textContent = "select a sheet";
    logEditLifecycleEvent("interop_apply_cell_edit", "error", "edit failed: select a sheet");
    return;
  }
  if (!cell) {
    editOutputEl.textContent = "enter an A1 cell or range reference";
    logEditLifecycleEvent(
      "interop_apply_cell_edit",
      "error",
      `edit failed: missing target cell in ${sheet}`,
    );
    return;
  }

  applyEditButtonEl.disabled = true;
  formulaBarApplyButtonEl.disabled = true;
  const commandContext = startUiCommand("interop_apply_cell_edit");
  logEditLifecycleEvent(
    "interop_apply_cell_edit",
    "start",
    `editing ${sheet}!${cell} from ${source}`,
    { trace: commandContext },
  );
  editOutputEl.textContent =
    source === "formula-bar" ? "Applying edit from formula bar..." : "Applying edit...";
  try {
    const payload = await invoke<InteropCellEditResponse>("interop_apply_cell_edit", {
      sheet,
      cell,
      input,
      mode,
      trace: commandContext,
    });
    const view = {
      workbookId: payload.workbookId,
      sheet: payload.sheet,
      cell: payload.cell,
      anchorCell: payload.anchorCell,
      appliedCellCount: payload.appliedCellCount,
      mode: payload.mode,
      value: formatCellValue(payload.value),
      formula: payload.formula,
      dirtySheetCount: payload.dirtySheetCount,
      dirtySheets: payload.dirtySheets,
      trace: payload.trace,
    };
    editOutputEl.textContent = jsonWithTraceHeader(
      payload.trace,
      JSON.stringify(view, null, 2),
      commandContext.commandId,
    );
    lastEditedCell = { sheet: payload.sheet, cell: normalizeA1(payload.anchorCell) };
    updateJumpLastEditedButton();
    if (payload.appliedCellCount > 1) {
      previewCopyStatusEl.textContent =
        `Last edit: ${payload.sheet}!${payload.cell} ` +
        `(${payload.appliedCellCount} cells; anchor ${payload.anchorCell}).`;
      announceToScreenReader(
        `Edited ${payload.appliedCellCount} cells starting at ${payload.sheet}!${payload.anchorCell}`,
      );
    } else {
      previewCopyStatusEl.textContent = `Last edited cell set to ${payload.sheet}!${payload.anchorCell}.`;
      announceToScreenReader(`Edited ${payload.sheet}!${payload.anchorCell}.`);
    }
    logEditLifecycleEvent(
      "interop_apply_cell_edit",
      "success",
      `edited ${payload.sheet}!${payload.anchorCell}`,
      { trace: payload.trace },
    );
    markRecalcPending();
    await loadInteropSessionStatus();
    await loadSheetPreview(payload.sheet, payload.anchorCell, true, payload.trace);
  } catch (error) {
    editOutputEl.textContent = `edit error: ${String(error)}`;
    logEditLifecycleEvent("interop_apply_cell_edit", "error", `edit failed: ${String(error)}`, {
      trace: commandContext,
      error,
    });
  } finally {
    applyEditButtonEl.disabled = false;
    updatePreviewCopyButtons();
  }
}

async function applyFormulaBarEdit(): Promise<void> {
  if (!selectedPreviewCell) {
    previewCopyStatusEl.textContent = "Select a preview cell first.";
    return;
  }
  const selection = selectedPreviewCell;
  syncEditFormFromFormulaBarSelectionInput();
  await applyCellEdit({
    sheet: selection.sheet,
    cell: selection.cell.cell,
    input: formulaBarInputEl.value,
    mode: normalizeEditMode(formulaBarModeSelectEl.value),
    source: "formula-bar",
  });
}

async function applyHistoryAction(action: "undo" | "redo"): Promise<void> {
  if (!currentSession?.loaded) {
    previewCopyStatusEl.textContent = "open a workbook first";
    logEditLifecycleEvent(
      `interop_${action}_edit`,
      "error",
      `${action} failed: no active workbook session`,
    );
    return;
  }

  const commandContext = startUiCommand(`interop_${action}_edit`);
  logEditLifecycleEvent(`interop_${action}_edit`, "start", `${action} requested`, {
    trace: commandContext,
  });
  try {
    const payload = await invoke<InteropUndoRedoResponse>(`interop_${action}_edit`, {
      trace: commandContext,
    });

    const summary = `${payload.action} complete`;
    previewCopyStatusEl.textContent = summary;
    announceToScreenReader(summary);
    logEditLifecycleEvent(`interop_${action}_edit`, "success", summary, { trace: payload.trace });
    if (currentSession) {
      currentSession.undoCount = payload.undoCount;
      currentSession.redoCount = payload.redoCount;
      updateHistoryControls();
    }
    await loadInteropSessionStatus();
    if (selectedPreviewCell) {
      await loadSheetPreview(selectedPreviewCell.sheet, selectedPreviewCell.cell.cell, true, payload.trace);
    } else {
      await loadSheetPreview(undefined, undefined, false, payload.trace);
    }
  } catch (error) {
    previewCopyStatusEl.textContent = `${action} failed: ${String(error)}`;
    logEditLifecycleEvent(`interop_${action}_edit`, "error", `${action} failed: ${String(error)}`, {
      trace: commandContext,
      error,
    });
    await loadInteropSessionStatus();
  }
}

function collectMacroPermissionConfig(): InteropMacroPermissionConfig {
  return {
    fsRead: macroPermissionFsReadInputEl.checked,
    fsWrite: macroPermissionFsWriteInputEl.checked,
    netHttp: macroPermissionNetHttpInputEl.checked,
    clipboard: macroPermissionClipboardInputEl.checked,
    processExec: macroPermissionProcessExecInputEl.checked,
    udf: macroPermissionUdfInputEl.checked,
    eventsEmit: macroPermissionEventsEmitInputEl.checked,
  };
}

async function runMacro(): Promise<void> {
  if (!currentSession?.loaded) {
    macroOutputEl.textContent = "open a workbook first";
    logEditLifecycleEvent("interop_run_macro", "error", "macro failed: open a workbook first");
    return;
  }

  const scriptPath = macroScriptPathInputEl.value.trim();
  if (!scriptPath) {
    macroOutputEl.textContent = "provide a Python script path first";
    logEditLifecycleEvent("interop_run_macro", "error", "macro failed: missing script path");
    return;
  }

  if (
    !permissionsEqual(
      collectMacroPermissionConfig(),
      loadMacroPermissionPolicy(scriptPath) ?? {
        fsRead: false,
        fsWrite: false,
        netHttp: false,
        clipboard: false,
        processExec: false,
        udf: false,
        eventsEmit: false,
      },
    )
  ) {
    macroRememberPolicyInputEl.checked = false;
  }

  const permissionConfig = collectMacroPermissionConfig();
  const policySource = resolveMacroPolicySource(scriptPath);
  const needsConsent = policySource === "fresh" && hasElevatedMacroPermissions(permissionConfig);
  if (needsConsent && !window.confirm(formatMacroPolicyPrompt(scriptPath, permissionConfig))) {
    macroOutputEl.textContent = "macro run cancelled: elevated permission policy not trusted.";
    logEditLifecycleEvent(
      "interop_run_macro",
      "error",
      "macro run cancelled: user declined trust prompt",
    );
    return;
  }

  const macroName = macroNameInputEl.value.trim() || "main";
  const commandContext = startUiCommand("interop_run_macro");
  runMacroButtonEl.disabled = true;
  macroOutputEl.textContent = "Running macro...";
  logEditLifecycleEvent("interop_run_macro", "start", `running ${macroName} from ${scriptPath}`, {
    trace: commandContext,
  });

  try {
    const payload = await invoke<InteropRunMacroResponse>("interop_run_macro", {
      scriptPath,
      macroName,
      args: macroArgsInputEl.value,
      permissions: permissionConfig,
      trace: commandContext,
    });
    const resolvedPolicySource = payload.policySource ?? policySource;
    const view = {
      scriptPath: payload.scriptPath,
      macroName: payload.macroName,
      requestedPermissions: payload.requestedPermissions,
      scriptFingerprint: payload.scriptFingerprint,
      trust: payload.trust,
      runtimeEvents: payload.runtimeEvents,
      permissionEvents: payload.permissionEvents,
      policySource: resolvedPolicySource,
      permissionGranted: payload.permissionGranted,
      permissionDenied: payload.permissionDenied,
      mutationCount: payload.mutationCount,
      changedSheets: payload.changedSheets,
      mutations: payload.mutations,
      recalcReports: payload.recalcReports,
      workbookId: payload.workbookId,
      stdout: payload.stdout,
      stderr: payload.stderr,
      trace: payload.trace,
    };
    macroOutputEl.textContent = jsonWithTraceHeader(
      payload.trace,
      JSON.stringify(view, null, 2),
      commandContext.commandId,
    );
    logEditLifecycleEvent("interop_run_macro", "success", "macro run complete", {
      trace: payload.trace,
    });
    if (payload.changedSheets.length > 0) {
      markRecalcPending();
      await loadInteropSessionStatus();
      await loadSheetPreview(payload.changedSheets[0], undefined, false, payload.trace);
    } else {
      await loadInteropSessionStatus();
    }

    if (macroRememberPolicyInputEl.checked) {
      saveMacroPermissionPolicy(scriptPath, permissionConfig);
      renderMacroPolicyStatus("stored", permissionConfig);
      announceToScreenReader(`saved macro permission policy for ${scriptPath}`);
      applySavedMacroPermissionPolicy(scriptPath);
    } else if (loadMacroPermissionPolicy(scriptPath) !== null) {
      deleteMacroPermissionPolicy(scriptPath);
      renderMacroPolicyStatus("fresh", permissionConfig);
      announceToScreenReader(`cleared macro permission policy for ${scriptPath}`);
    }
  } catch (error) {
    macroOutputEl.textContent = `macro run error: ${String(error)}`;
    logEditLifecycleEvent("interop_run_macro", "error", `macro run failed: ${String(error)}`, {
      trace: commandContext,
      error,
    });
  } finally {
    runMacroButtonEl.disabled = false;
  }
}

async function pickSavePath(): Promise<void> {
  const sourcePath = openPathInputEl.value.trim();
  const mode = normalizeMode(saveModeSelectEl.value);
  const suggestion = sourcePath.length > 0 ? suggestOutputPath(sourcePath, mode) : undefined;

  try {
    const selected = await saveDialog({
      title: "Save RootCellar Workbook",
      defaultPath: suggestion,
      filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
    });
    if (typeof selected !== "string") {
      return;
    }
    savePathInputEl.value = selected;
  } catch (error) {
    saveOutputEl.textContent = `save dialog error: ${String(error)}`;
  }
}

async function saveWorkbook(): Promise<void> {
  const outputPath = savePathInputEl.value.trim();
  if (!outputPath) {
    saveOutputEl.textContent = "enter an output path first";
    logEditLifecycleEvent("interop_save_workbook", "error", "save failed: missing output path");
    return;
  }

  const mode = normalizeMode(saveModeSelectEl.value);
  const promoteOutputAsInput = savePromoteSourceInputEl.checked;
  saveWorkbookButtonEl.disabled = true;
  saveOutputEl.textContent = "Saving workbook...";
  const commandContext = startUiCommand("interop_save_workbook");
  logEditLifecycleEvent("interop_save_workbook", "start", `saving workbook to ${outputPath}`, {
    trace: commandContext,
  });

  try {
    const payload = await invoke<InteropSaveResponse>("interop_save_workbook", {
      outputPath,
      mode,
      promoteOutputAsInput,
      trace: commandContext,
    });
    saveOutputEl.textContent = jsonWithTraceHeader(
      payload.trace,
      JSON.stringify(payload, null, 2),
      commandContext.commandId,
    );
    if (promoteOutputAsInput) {
      openPathInputEl.value = payload.outputPath;
    }
    logEditLifecycleEvent("interop_save_workbook", "success", `saved ${payload.outputPath}`, {
      trace: payload.trace,
    });
    await loadInteropSessionStatus();
  } catch (error) {
    saveOutputEl.textContent = `save workbook error: ${String(error)}`;
    logEditLifecycleEvent("interop_save_workbook", "error", `save failed: ${String(error)}`, {
      trace: commandContext,
      error,
    });
  } finally {
    saveWorkbookButtonEl.disabled = false;
  }
}

async function recalcLoadedWorkbook(): Promise<void> {
  const sheet = recalcSheetInputEl.value.trim();
  recalcLoadedButtonEl.disabled = true;
  saveOutputEl.textContent = "Recalculating loaded workbook...";
  const commandContext = startUiCommand("interop_recalc_loaded");
  logEditLifecycleEvent(
    "interop_recalc_loaded",
    "start",
    sheet.length > 0 ? `recalculating ${sheet}` : "recalculating loaded workbook",
    { trace: commandContext },
  );

  try {
    const payload = await invoke<InteropRecalcResponse>("interop_recalc_loaded", {
      sheet: sheet.length > 0 ? sheet : null,
      trace: commandContext,
    });
    saveOutputEl.textContent = jsonWithTraceHeader(
      payload.trace,
      JSON.stringify(payload, null, 2),
      commandContext.commandId,
    );
    const resolvedScope =
      sheet.length > 0
        ? sheet
        : payload.reports.length === 1
          ? payload.reports[0].sheet
          : null;
    markRecalcCompleted(resolvedScope);
    logEditLifecycleEvent(
      "interop_recalc_loaded",
      "success",
      `recalc complete for ${resolvedScope ?? "all loaded sheets"}`,
      { trace: payload.trace },
    );
    await loadInteropSessionStatus();
    await loadSheetPreview(sheet.length > 0 ? sheet : undefined, undefined, false, payload.trace);
  } catch (error) {
    saveOutputEl.textContent = `recalc error: ${String(error)}`;
    logEditLifecycleEvent("interop_recalc_loaded", "error", `recalc failed: ${String(error)}`, {
      trace: commandContext,
      error,
    });
  } finally {
    recalcLoadedButtonEl.disabled = false;
  }
}

async function runEngineRoundTrip(): Promise<void> {
  runRoundTripButtonEl.disabled = true;
  roundTripOutputEl.textContent = "Running engine round-trip...";
  const commandContext = startUiCommand("engine_round_trip");
  logEditLifecycleEvent("engine_round_trip", "start", "running engine round-trip", {
    trace: commandContext,
  });

  try {
    const payload = await invoke<EngineRoundTripResponse>("engine_round_trip", {
      trace: commandContext,
    });

    const view = {
      sheet: payload.sheet,
      formulaCell: payload.formulaCell,
      value: formatCellValue(payload.value),
      evaluatedCells: payload.evaluatedCells,
      cycleCount: payload.cycleCount,
      parseErrorCount: payload.parseErrorCount,
      workbookId: payload.workbookId,
      trace: payload.trace,
    };
    roundTripOutputEl.textContent = jsonWithTraceHeader(
      payload.trace,
      JSON.stringify(view, null, 2),
      commandContext.commandId,
    );
    logEditLifecycleEvent("engine_round_trip", "success", "engine round-trip complete", {
      trace: payload.trace,
    });
  } catch (error) {
    roundTripOutputEl.textContent = `round-trip error: ${String(error)}`;
    logEditLifecycleEvent("engine_round_trip", "error", `engine round-trip failed: ${String(error)}`, {
      trace: commandContext,
      error,
    });
  } finally {
    runRoundTripButtonEl.disabled = false;
  }
}

refreshButtonEl.addEventListener("click", () => {
  void loadAppStatus();
});

runRoundTripButtonEl.addEventListener("click", () => {
  void runEngineRoundTrip();
});

pickOpenPathButtonEl.addEventListener("click", () => {
  void pickOpenPath();
});

openWorkbookButtonEl.addEventListener("click", () => {
  void openWorkbook();
});

refreshSessionButtonEl.addEventListener("click", async () => {
  await loadInteropSessionStatus();
  await loadSheetPreview();
});

refreshPreviewButtonEl.addEventListener("click", () => {
  void loadSheetPreview();
});

jumpLastEditedButtonEl.addEventListener("click", () => {
  void jumpToLastEdited();
});

copyPreviewA1ButtonEl.addEventListener("click", () => {
  void copySelectedA1();
});

copyPreviewValueButtonEl.addEventListener("click", () => {
  void copySelectedValue();
});

copyPreviewFormulaButtonEl.addEventListener("click", () => {
  void copySelectedFormula();
});

pasteClipboardCellButtonEl.addEventListener("click", () => {
  void pasteFromClipboardIntoSelection();
});

formulaBarModeSelectEl.addEventListener("change", () => {
  syncEditFormFromFormulaBarSelectionInput();
});

formulaBarInputEl.addEventListener("input", () => {
  syncEditFormFromFormulaBarSelectionInput();
});

bindFormulaBarApplyHandlers(formulaBarInputEl, formulaBarApplyButtonEl, () => {
  void applyFormulaBarEdit();
});

previewSheetSelectEl.addEventListener("change", () => {
  void loadSheetPreview(previewSheetSelectEl.value);
});

previewTableEl.addEventListener("click", (event) => {
  const target = event.target as HTMLElement;
  const cellEl = target.closest("td[data-cell]") as HTMLTableCellElement | null;
  if (!cellEl || !currentPreview) {
    return;
  }

  const cellRef = normalizeA1(cellEl.dataset.cell ?? "");
  if (cellRef.length === 0) {
    return;
  }

  const matched = currentPreview.cells.find((cell) => cell.cell === cellRef);
  if (!matched) {
    return;
  }

  setSelectedPreviewCell({ sheet: currentPreview.sheet, cell: matched });
  previewCopyStatusEl.textContent = `Selected ${currentPreview.sheet}!${matched.cell}`;
  renderPreviewTable(currentPreview);
  previewWrapEl.focus();
});

previewWrapEl.addEventListener("keydown", (event) => {
  if (!currentPreview) {
    return;
  }

  const key = event.key;
  const isUndo = (key === "z" || key === "Z") && (event.ctrlKey || event.metaKey);
  const isRedo = (key === "y" || key === "Y") && (event.ctrlKey || event.metaKey);
  const isPaste = (key === "v" || key === "V") && (event.ctrlKey || event.metaKey);
  if (event.key === "ArrowUp") {
    event.preventDefault();
    movePreviewSelection("up");
    return;
  }
  if (event.key === "ArrowDown") {
    event.preventDefault();
    movePreviewSelection("down");
    return;
  }
  if (event.key === "ArrowLeft") {
    event.preventDefault();
    movePreviewSelection("left");
    return;
  }
  if (event.key === "ArrowRight") {
    event.preventDefault();
    movePreviewSelection("right");
    return;
  }
  if (event.key === "Enter") {
    event.preventDefault();
    void applyCellEdit();
    return;
  }
  if (isUndo) {
    event.preventDefault();
    void applyHistoryAction("undo");
    return;
  }
  if (isRedo) {
    event.preventDefault();
    void applyHistoryAction("redo");
    return;
  }
  if (isPaste) {
    event.preventDefault();
    void pasteFromClipboardIntoSelection();
  }
});

formulaBarInputEl.addEventListener("keydown", (event) => {
  if ((event.key === "z" || event.key === "Z") && (event.ctrlKey || event.metaKey)) {
    event.preventDefault();
    void applyHistoryAction("undo");
    return;
  }
  if ((event.key === "y" || event.key === "Y") && (event.ctrlKey || event.metaKey)) {
    event.preventDefault();
    void applyHistoryAction("redo");
  }
});

macroScriptPathInputEl.addEventListener("blur", () => {
  applySavedMacroPermissionPolicy(normalizeMacroScriptPath());
});

macroScriptPathInputEl.addEventListener("change", () => {
  applySavedMacroPermissionPolicy(normalizeMacroScriptPath());
});

applyEditButtonEl.addEventListener("click", () => {
  void applyCellEdit();
});

runMacroButtonEl.addEventListener("click", () => {
  void runMacro();
});

macroClearPolicyButtonEl.addEventListener("click", () => {
  const scriptPath = normalizeMacroScriptPath();
  if (!scriptPath) {
    macroPolicyStatusEl.textContent = "No script path to clear policy for.";
    return;
  }
  if (!loadMacroPermissionPolicy(scriptPath)) {
    macroPolicyStatusEl.textContent = "No saved policy found for this script path.";
    return;
  }

  deleteMacroPermissionPolicy(scriptPath);
  macroRememberPolicyInputEl.checked = false;
  renderMacroPolicyStatus("fresh", collectMacroPermissionConfig());
  announceToScreenReader(`cleared macro permission policy for ${scriptPath}`);
});

presetRow3ButtonEl.addEventListener("click", () => {
  applyRangePreset(1, 3, "Row x3");
});

presetCol3ButtonEl.addEventListener("click", () => {
  applyRangePreset(3, 1, "Col x3");
});

presetBlock2x2ButtonEl.addEventListener("click", () => {
  applyRangePreset(2, 2, "Block 2x2");
});

undoLastEditButtonEl.addEventListener("click", () => {
  void applyHistoryAction("undo");
});

redoLastEditButtonEl.addEventListener("click", () => {
  void applyHistoryAction("redo");
});

const applyCustomPreset = (): void => {
  const rows = parsePresetSpanInput(
    presetRowsInputEl.value,
    1,
    maxPresetRows,
    "Preset rows",
  );
  if (rows == null) {
    return;
  }

  const cols = parsePresetSpanInput(
    presetColsInputEl.value,
    1,
    maxPresetCols,
    "Preset cols",
  );
  if (cols == null) {
    return;
  }

  const applied = applyRangePreset(rows, cols, `Custom ${rows}x${cols}`);
  if (!applied) {
    return;
  }

  lastCustomPreset = { rows, cols };
  persistLastCustomPreset(lastCustomPreset);
  renderLastCustomPresetButton();
};

presetApplyCustomButtonEl.addEventListener("click", () => {
  applyCustomPreset();
});

presetApplyLastButtonEl.addEventListener("click", () => {
  if (!lastCustomPreset) {
    editOutputEl.textContent = "No last custom preset available yet.";
    return;
  }

  presetRowsInputEl.value = String(lastCustomPreset.rows);
  presetColsInputEl.value = String(lastCustomPreset.cols);
  applyRangePreset(lastCustomPreset.rows, lastCustomPreset.cols, `Last ${lastCustomPreset.rows}x${lastCustomPreset.cols}`);
});

presetRowsInputEl.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") {
    return;
  }
  event.preventDefault();
  applyCustomPreset();
});

presetColsInputEl.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") {
    return;
  }
  event.preventDefault();
  applyCustomPreset();
});

pickSavePathButtonEl.addEventListener("click", () => {
  void pickSavePath();
});

saveWorkbookButtonEl.addEventListener("click", () => {
  void saveWorkbook();
});

recalcLoadedButtonEl.addEventListener("click", () => {
  void recalcLoadedWorkbook();
});

saveModeSelectEl.addEventListener("change", () => {
  const currentInputPath = openPathInputEl.value.trim();
  if (!currentInputPath || savePathInputEl.value.trim()) {
    return;
  }
  savePathInputEl.value = suggestOutputPath(currentInputPath, normalizeMode(saveModeSelectEl.value));
});

clearPreviewTable("Open a workbook to render sheet preview.");
setSelectedPreviewCell(null);
previewCopyStatusEl.textContent = "No copy action yet.";
renderEditLifecycleOutput();
updateJumpLastEditedButton();
renderRecalcStatus();
lastCustomPreset = loadLastCustomPresetFromStorage();
if (lastCustomPreset) {
  presetRowsInputEl.value = String(lastCustomPreset.rows);
  presetColsInputEl.value = String(lastCustomPreset.cols);
}
renderLastCustomPresetButton();
applySavedMacroPermissionPolicy(normalizeMacroScriptPath());

if (uiCaptureMode) {
  seedUiCaptureDemo();
} else {
  void loadAppStatus();
  void loadInteropSessionStatus();
}

