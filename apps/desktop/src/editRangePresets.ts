export const EXCEL_MAX_ROW = 1_048_576;
export const EXCEL_MAX_COL = 16_384;

export type A1CellRef = {
  row: number;
  col: number;
};

export type BuildPresetRangeResult =
  | {
      ok: true;
      range: string;
      anchor: A1CellRef;
    }
  | {
      ok: false;
      reason: "invalid_anchor" | "invalid_span" | "out_of_bounds";
    };

export function parseSingleA1Cell(value: string): A1CellRef | null {
  const normalized = value.trim().toUpperCase();
  if (normalized.length === 0 || normalized.includes(":")) {
    return null;
  }

  const match = normalized.match(/^([A-Z]+)([1-9][0-9]*)$/);
  if (!match) {
    return null;
  }

  const [, colPart, rowPart] = match;
  let col = 0;
  for (const ch of colPart) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }

  const row = Number.parseInt(rowPart, 10);
  if (!Number.isFinite(row) || row <= 0 || col <= 0) {
    return null;
  }

  return { row, col };
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

function formatA1Cell(row: number, col: number): string {
  return `${columnToA1(col)}${row}`;
}

export function buildPresetRange(
  anchorA1: string,
  rowSpan: number,
  colSpan: number,
): BuildPresetRangeResult {
  if (!Number.isInteger(rowSpan) || !Number.isInteger(colSpan) || rowSpan < 1 || colSpan < 1) {
    return { ok: false, reason: "invalid_span" };
  }

  const anchor = parseSingleA1Cell(anchorA1);
  if (!anchor) {
    return { ok: false, reason: "invalid_anchor" };
  }

  const endRow = anchor.row + rowSpan - 1;
  const endCol = anchor.col + colSpan - 1;
  if (endRow > EXCEL_MAX_ROW || endCol > EXCEL_MAX_COL) {
    return { ok: false, reason: "out_of_bounds" };
  }

  const start = formatA1Cell(anchor.row, anchor.col);
  const end = formatA1Cell(endRow, endCol);
  return {
    ok: true,
    anchor,
    range: rowSpan === 1 && colSpan === 1 ? start : `${start}:${end}`,
  };
}
