import { invoke } from "@tauri-apps/api/core";
import { open as openDialog, save as saveDialog } from "@tauri-apps/plugin-dialog";
import "./styles.css";

type TraceInput = {
  traceId: string;
  spanId: string;
  parentSpanId: string | null;
  sessionId: string;
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

type SelectedPreviewCell = {
  sheet: string;
  cell: InteropPreviewCell;
};

type EditActionSource = "edit-form" | "formula-bar";

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

function newTrace(): TraceInput {
  return {
    traceId: crypto.randomUUID(),
    spanId: crypto.randomUUID(),
    parentSpanId: null,
    sessionId: crypto.randomUUID(),
  };
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

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) {
  throw new Error("missing #app container");
}

app.innerHTML = `
  <main class="shell">
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
      <pre id="preview-selected-output" class="output compact">No cell selected.</pre>
      <pre id="preview-copy-status" class="output compact">No copy action yet.</pre>
      <div id="preview-wrap" class="table-wrap" tabindex="0">
        <table id="preview-table" class="preview-table"></table>
      </div>
    </section>

    <section class="panel">
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
      <div class="actions">
        <button id="apply-edit" class="btn primary">Apply Cell Edit</button>
      </div>
      <pre id="edit-output" class="output">No edits yet.</pre>
    </section>

    <section class="panel">
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
      <pre id="save-output" class="output">No save/recalc run yet.</pre>
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
const formulaBarCellInput = document.querySelector<HTMLInputElement>("#formula-bar-cell");
const formulaBarModeSelect = document.querySelector<HTMLSelectElement>("#formula-bar-mode");
const formulaBarInput = document.querySelector<HTMLInputElement>("#formula-bar-input");
const formulaBarApplyButton = document.querySelector<HTMLButtonElement>("#formula-bar-apply");
const previewOutput = document.querySelector<HTMLPreElement>("#preview-output");
const previewSelectedOutput = document.querySelector<HTMLPreElement>("#preview-selected-output");
const previewCopyStatus = document.querySelector<HTMLPreElement>("#preview-copy-status");
const previewWrap = document.querySelector<HTMLDivElement>("#preview-wrap");
const previewTable = document.querySelector<HTMLTableElement>("#preview-table");
const editSheetSelect = document.querySelector<HTMLSelectElement>("#edit-sheet");
const editCellInput = document.querySelector<HTMLInputElement>("#edit-cell");
const editModeSelect = document.querySelector<HTMLSelectElement>("#edit-mode");
const editInput = document.querySelector<HTMLInputElement>("#edit-input");
const applyEditButton = document.querySelector<HTMLButtonElement>("#apply-edit");
const editOutput = document.querySelector<HTMLPreElement>("#edit-output");
const savePathInput = document.querySelector<HTMLInputElement>("#save-path");
const savePromoteSourceInput = document.querySelector<HTMLInputElement>("#save-promote-source");
const pickSavePathButton = document.querySelector<HTMLButtonElement>("#pick-save-path");
const saveModeSelect = document.querySelector<HTMLSelectElement>("#save-mode");
const saveWorkbookButton = document.querySelector<HTMLButtonElement>("#save-workbook");
const recalcSheetInput = document.querySelector<HTMLInputElement>("#recalc-sheet");
const recalcLoadedButton = document.querySelector<HTMLButtonElement>("#recalc-loaded");
const saveOutput = document.querySelector<HTMLPreElement>("#save-output");

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
  !formulaBarCellInput ||
  !formulaBarModeSelect ||
  !formulaBarInput ||
  !formulaBarApplyButton ||
  !previewOutput ||
  !previewSelectedOutput ||
  !previewCopyStatus ||
  !previewWrap ||
  !previewTable ||
  !editSheetSelect ||
  !editCellInput ||
  !editModeSelect ||
  !editInput ||
  !applyEditButton ||
  !editOutput ||
  !savePathInput ||
  !savePromoteSourceInput ||
  !pickSavePathButton ||
  !saveModeSelect ||
  !saveWorkbookButton ||
  !recalcSheetInput ||
  !recalcLoadedButton ||
  !saveOutput
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
const formulaBarCellInputEl = formulaBarCellInput;
const formulaBarModeSelectEl = formulaBarModeSelect;
const formulaBarInputEl = formulaBarInput;
const formulaBarApplyButtonEl = formulaBarApplyButton;
const previewOutputEl = previewOutput;
const previewSelectedOutputEl = previewSelectedOutput;
const previewCopyStatusEl = previewCopyStatus;
const previewWrapEl = previewWrap;
const previewTableEl = previewTable;
const editSheetSelectEl = editSheetSelect;
const editCellInputEl = editCellInput;
const editModeSelectEl = editModeSelect;
const editInputEl = editInput;
const applyEditButtonEl = applyEditButton;
const editOutputEl = editOutput;
const savePathInputEl = savePathInput;
const savePromoteSourceInputEl = savePromoteSourceInput;
const pickSavePathButtonEl = pickSavePathButton;
const saveModeSelectEl = saveModeSelect;
const saveWorkbookButtonEl = saveWorkbookButton;
const recalcSheetInputEl = recalcSheetInput;
const recalcLoadedButtonEl = recalcLoadedButton;
const saveOutputEl = saveOutput;

let currentSession: InteropSessionStatusResponse | null = null;
let currentPreview: InteropSheetPreviewResponse | null = null;
let selectedPreviewCell: SelectedPreviewCell | null = null;
let lastEditedCell: { sheet: string; cell: string } | null = null;

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
}

function updateJumpLastEditedButton(): void {
  jumpLastEditedButtonEl.disabled = !lastEditedCell;
}

type PreviewIndex = {
  rows: number[];
  cols: number[];
  rowToCols: Map<number, number[]>;
  cellMap: Map<string, InteropPreviewCell>;
};

function buildPreviewIndex(preview: InteropSheetPreviewResponse): PreviewIndex {
  const rowSet = new Set<number>();
  const colSet = new Set<number>();
  const rowToCols = new Map<number, number[]>();
  const cellMap = new Map<string, InteropPreviewCell>();

  for (const cell of preview.cells) {
    rowSet.add(cell.row);
    colSet.add(cell.col);
    const cols = rowToCols.get(cell.row) ?? [];
    cols.push(cell.col);
    rowToCols.set(cell.row, cols);
    cellMap.set(`${cell.row}:${cell.col}`, cell);
  }

  const rows = Array.from(rowSet).sort((a, b) => a - b);
  const cols = Array.from(colSet).sort((a, b) => a - b);
  for (const [row, colsForRow] of rowToCols.entries()) {
    rowToCols.set(row, colsForRow.sort((a, b) => a - b));
  }

  return { rows, cols, rowToCols, cellMap };
}

function nearestAvailableCol(targetCol: number, candidates: number[]): number {
  if (candidates.includes(targetCol)) {
    return targetCol;
  }
  return candidates.reduce((best, candidate) =>
    Math.abs(candidate - targetCol) < Math.abs(best - targetCol) ? candidate : best,
  candidates[0]);
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

  const index = buildPreviewIndex(currentPreview);
  if (index.rows.length === 0 || index.cols.length === 0) {
    return;
  }

  const selected = selectedPreviewCell;
  const baseRow = selected?.cell.row ?? index.rows[0];
  const baseCol = selected?.cell.col ?? index.cols[0];
  const rowIndex = Math.max(index.rows.indexOf(baseRow), 0);
  const colIndex = Math.max(index.cols.indexOf(baseCol), 0);

  let targetRowIndex = rowIndex;
  let targetColIndex = colIndex;
  if (direction === "up") {
    targetRowIndex = Math.max(0, rowIndex - 1);
  } else if (direction === "down") {
    targetRowIndex = Math.min(index.rows.length - 1, rowIndex + 1);
  } else if (direction === "left") {
    targetColIndex = Math.max(0, colIndex - 1);
  } else {
    targetColIndex = Math.min(index.cols.length - 1, colIndex + 1);
  }

  const targetRow = index.rows[targetRowIndex];
  const desiredCol = index.cols[targetColIndex];
  const rowCols = index.rowToCols.get(targetRow);
  if (!rowCols || rowCols.length === 0) {
    return;
  }
  const resolvedCol = nearestAvailableCol(desiredCol, rowCols);
  const targetCell = index.cellMap.get(`${targetRow}:${resolvedCol}`);
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
    "<thead><tr><th class=\"corner\">#</th>",
    ...cols.map((col) => `<th>${columnToA1(col)}</th>`),
    "</tr></thead>",
  ].join("");

  const tbodyRows = rows.map((row) => {
    const cells = cols.map((col) => {
      const key = `${row}:${col}`;
      const cell = cellMap.get(key);
      if (!cell) {
        return "<td></td>";
      }
      const classes = ["preview-cell"];
      if (
        selectedPreviewCell &&
        selectedPreviewCell.sheet === payload.sheet &&
        selectedPreviewCell.cell.cell === cell.cell
      ) {
        classes.push("selected");
      }
      if (
        lastEditedCell &&
        lastEditedCell.sheet === payload.sheet &&
        lastEditedCell.cell === cell.cell
      ) {
        classes.push("last-edited");
      }

      return `<td class="${classes.join(" ")}" data-cell="${escapeHtml(cell.cell)}" data-row="${cell.row}" data-col="${cell.col}" title="${escapeHtml(cell.cell)}">${escapeHtml(formatPreviewCell(cell))}</td>`;
    });
    return `<tr><th>${row}</th>${cells.join("")}</tr>`;
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

  if (!payload.loaded) {
    currentPreview = null;
    lastEditedCell = null;
    updateJumpLastEditedButton();
    setSelectedPreviewCell(null);
    compatOutputEl.textContent = "No workbook loaded.";
    previewOutputEl.textContent = "No preview loaded.";
    previewCopyStatusEl.textContent = "No copy action yet.";
    clearPreviewTable("Open a workbook to render sheet preview.");
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

  refreshPreviewButtonEl.disabled = true;
  previewOutputEl.textContent = "Loading sheet preview...";
  try {
    const payload = await invoke<InteropSheetPreviewResponse>("interop_sheet_preview", {
      sheet: selectedSheet.length > 0 ? selectedSheet : null,
      limit,
      trace: newTrace(),
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
      `Workbook: ${payload.workbookId}`,
      `Sheet: ${payload.sheet}`,
      `Shown cells: ${payload.shownCells}/${payload.totalCells}`,
      `Truncated: ${payload.truncated ? "yes" : "no"}`,
    ];
    previewOutputEl.textContent = summary.join("\n");
    renderPreviewTable(payload);
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
    return;
  }

  openWorkbookButtonEl.disabled = true;
  sessionOutputEl.textContent = "Opening workbook...";
  compatOutputEl.textContent = "Inspecting compatibility...";

  try {
    const payload = await invoke<InteropOpenResponse>("interop_open_workbook", {
      path,
      trace: newTrace(),
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
      sheets: payload.sheets,
    };
    renderSessionStatus(statusPayload);
    compatOutputEl.textContent = formatCompatibility(payload);
    lastEditedCell = null;
    updateJumpLastEditedButton();
    setSelectedPreviewCell(null);
    previewCopyStatusEl.textContent = "No copy action yet.";
    savePathInputEl.value = suggestOutputPath(payload.inputPath, normalizeMode(saveModeSelectEl.value));
    await loadSheetPreview(payload.sheets[0]);
  } catch (error) {
    const message = `open workbook error: ${String(error)}`;
    sessionOutputEl.textContent = message;
    compatOutputEl.textContent = message;
    previewCopyStatusEl.textContent = message;
    previewOutputEl.textContent = message;
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
    return;
  }

  const source = overrides?.source ?? "edit-form";
  const sheet = (overrides?.sheet ?? editSheetSelectEl.value).trim();
  const cell = (overrides?.cell ?? editCellInputEl.value).trim();
  const input = overrides?.input ?? editInputEl.value;
  const mode = overrides?.mode ?? normalizeEditMode(editModeSelectEl.value);

  if (!sheet) {
    editOutputEl.textContent = "select a sheet";
    return;
  }
  if (!cell) {
    editOutputEl.textContent = "enter an A1 cell or range reference";
    return;
  }

  applyEditButtonEl.disabled = true;
  formulaBarApplyButtonEl.disabled = true;
  editOutputEl.textContent =
    source === "formula-bar" ? "Applying edit from formula bar..." : "Applying edit...";
  try {
    const payload = await invoke<InteropCellEditResponse>("interop_apply_cell_edit", {
      sheet,
      cell,
      input,
      mode,
      trace: newTrace(),
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
    editOutputEl.textContent = JSON.stringify(view, null, 2);
    lastEditedCell = { sheet: payload.sheet, cell: normalizeA1(payload.anchorCell) };
    updateJumpLastEditedButton();
    if (payload.appliedCellCount > 1) {
      previewCopyStatusEl.textContent =
        `Last edit: ${payload.sheet}!${payload.cell} ` +
        `(${payload.appliedCellCount} cells; anchor ${payload.anchorCell}).`;
    } else {
      previewCopyStatusEl.textContent = `Last edited cell set to ${payload.sheet}!${payload.anchorCell}.`;
    }
    await loadInteropSessionStatus();
    await loadSheetPreview(payload.sheet, payload.anchorCell, true);
  } catch (error) {
    editOutputEl.textContent = `edit error: ${String(error)}`;
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
    return;
  }

  const mode = normalizeMode(saveModeSelectEl.value);
  const promoteOutputAsInput = savePromoteSourceInputEl.checked;
  saveWorkbookButtonEl.disabled = true;
  saveOutputEl.textContent = "Saving workbook...";

  try {
    const payload = await invoke<InteropSaveResponse>("interop_save_workbook", {
      outputPath,
      mode,
      promoteOutputAsInput,
      trace: newTrace(),
    });
    saveOutputEl.textContent = JSON.stringify(payload, null, 2);
    if (promoteOutputAsInput) {
      openPathInputEl.value = payload.outputPath;
    }
    await loadInteropSessionStatus();
  } catch (error) {
    saveOutputEl.textContent = `save workbook error: ${String(error)}`;
  } finally {
    saveWorkbookButtonEl.disabled = false;
  }
}

async function recalcLoadedWorkbook(): Promise<void> {
  const sheet = recalcSheetInputEl.value.trim();
  recalcLoadedButtonEl.disabled = true;
  saveOutputEl.textContent = "Recalculating loaded workbook...";

  try {
    const payload = await invoke<InteropRecalcResponse>("interop_recalc_loaded", {
      sheet: sheet.length > 0 ? sheet : null,
      trace: newTrace(),
    });
    saveOutputEl.textContent = JSON.stringify(payload, null, 2);
    await loadInteropSessionStatus();
    await loadSheetPreview(sheet.length > 0 ? sheet : undefined);
  } catch (error) {
    saveOutputEl.textContent = `recalc error: ${String(error)}`;
  } finally {
    recalcLoadedButtonEl.disabled = false;
  }
}

async function runEngineRoundTrip(): Promise<void> {
  runRoundTripButtonEl.disabled = true;
  roundTripOutputEl.textContent = "Running engine round-trip...";
  try {
    const payload = await invoke<EngineRoundTripResponse>("engine_round_trip", {
      trace: newTrace(),
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
    roundTripOutputEl.textContent = JSON.stringify(view, null, 2);
  } catch (error) {
    roundTripOutputEl.textContent = `round-trip error: ${String(error)}`;
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

formulaBarModeSelectEl.addEventListener("change", () => {
  syncEditFormFromFormulaBarSelectionInput();
});

formulaBarInputEl.addEventListener("input", () => {
  syncEditFormFromFormulaBarSelectionInput();
});

formulaBarInputEl.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") {
    return;
  }
  event.preventDefault();
  void applyFormulaBarEdit();
});

formulaBarApplyButtonEl.addEventListener("click", () => {
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
  }
});

applyEditButtonEl.addEventListener("click", () => {
  void applyCellEdit();
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

void loadAppStatus();
void loadInteropSessionStatus();
clearPreviewTable("Open a workbook to render sheet preview.");
setSelectedPreviewCell(null);
previewCopyStatusEl.textContent = "No copy action yet.";
updateJumpLastEditedButton();
