export type PreviewNavDirection = "up" | "down" | "left" | "right";

export type PreviewNavCell = {
  cell: string;
  row: number;
  col: number;
};

function nearestAvailableCol(targetCol: number, candidates: number[]): number {
  return candidates.reduce((best, candidate) =>
    Math.abs(candidate - targetCol) < Math.abs(best - targetCol) ? candidate : best,
  candidates[0]);
}

function nextColInDirection(
  direction: PreviewNavDirection,
  baseCol: number,
  rowCols: number[],
): number {
  if (direction === "right") {
    const next = rowCols.find((col) => col > baseCol);
    return next ?? rowCols[rowCols.length - 1];
  }
  if (direction === "left") {
    for (let idx = rowCols.length - 1; idx >= 0; idx -= 1) {
      if (rowCols[idx] < baseCol) {
        return rowCols[idx];
      }
    }
    return rowCols[0];
  }
  return nearestAvailableCol(baseCol, rowCols);
}

export function computeNextPreviewSelection<TCell extends PreviewNavCell>(
  cells: TCell[],
  selected: TCell | null,
  direction: PreviewNavDirection,
): TCell | null {
  if (cells.length === 0) {
    return null;
  }

  const rowToCols = new Map<number, number[]>();
  const cellMap = new Map<string, TCell>();
  for (const cell of cells) {
    const cols = rowToCols.get(cell.row) ?? [];
    cols.push(cell.col);
    rowToCols.set(cell.row, cols);
    cellMap.set(`${cell.row}:${cell.col}`, cell);
  }

  const rows = Array.from(rowToCols.keys()).sort((a, b) => a - b);
  for (const [row, cols] of rowToCols.entries()) {
    rowToCols.set(row, cols.sort((a, b) => a - b));
  }

  const firstRow = rows[0];
  const firstCols = rowToCols.get(firstRow);
  if (!firstCols || firstCols.length === 0) {
    return null;
  }

  const baseRow = selected?.row ?? firstRow;
  const baseRowCols = rowToCols.get(baseRow) ?? firstCols;
  const baseCol = selected?.col ?? baseRowCols[0];

  const baseRowIndex = Math.max(rows.indexOf(baseRow), 0);
  let targetRowIndex = baseRowIndex;
  if (direction === "up") {
    targetRowIndex = Math.max(0, baseRowIndex - 1);
  } else if (direction === "down") {
    targetRowIndex = Math.min(rows.length - 1, baseRowIndex + 1);
  }
  const targetRow = rows[targetRowIndex];

  const targetRowCols = rowToCols.get(targetRow);
  if (!targetRowCols || targetRowCols.length === 0) {
    return null;
  }

  const targetCol = nextColInDirection(direction, baseCol, targetRowCols);
  return cellMap.get(`${targetRow}:${targetCol}`) ?? null;
}

export function bindFormulaBarApplyHandlers(
  inputEl: HTMLInputElement,
  applyButtonEl: HTMLButtonElement,
  onApply: () => void,
): () => void {
  const handleInputKeydown = (event: KeyboardEvent): void => {
    if (event.key !== "Enter") {
      return;
    }
    event.preventDefault();
    onApply();
  };

  const handleApplyClick = (): void => {
    onApply();
  };

  inputEl.addEventListener("keydown", handleInputKeydown);
  applyButtonEl.addEventListener("click", handleApplyClick);

  return () => {
    inputEl.removeEventListener("keydown", handleInputKeydown);
    applyButtonEl.removeEventListener("click", handleApplyClick);
  };
}
