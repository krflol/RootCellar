import { describe, expect, it, vi } from "vitest";
import {
  bindFormulaBarApplyHandlers,
  computeNextPreviewSelection,
  type PreviewNavCell,
} from "./previewInteractions";

const PREVIEW_CELLS: PreviewNavCell[] = [
  { cell: "A1", row: 1, col: 1 },
  { cell: "C1", row: 1, col: 3 },
  { cell: "B2", row: 2, col: 2 },
  { cell: "C3", row: 3, col: 3 },
];

describe("computeNextPreviewSelection", () => {
  it("moves right to the next populated cell in the same row", () => {
    const selected = PREVIEW_CELLS[0];
    const next = computeNextPreviewSelection(PREVIEW_CELLS, selected, "right");
    expect(next?.cell).toBe("C1");
  });

  it("moves down using nearest available column in target row", () => {
    const selected = PREVIEW_CELLS[1];
    const next = computeNextPreviewSelection(PREVIEW_CELLS, selected, "down");
    expect(next?.cell).toBe("B2");
  });

  it("stays on boundary when moving beyond first row", () => {
    const selected = PREVIEW_CELLS[0];
    const next = computeNextPreviewSelection(PREVIEW_CELLS, selected, "up");
    expect(next?.cell).toBe("A1");
  });

  it("falls back to the first populated cell when selection is missing", () => {
    const next = computeNextPreviewSelection(PREVIEW_CELLS, null, "left");
    expect(next?.cell).toBe("A1");
  });
});

describe("bindFormulaBarApplyHandlers", () => {
  it("triggers apply on Enter and button click, and detaches cleanly", () => {
    const input = document.createElement("input");
    const button = document.createElement("button");
    const onApply = vi.fn();

    const dispose = bindFormulaBarApplyHandlers(input, button, onApply);

    const enterEvent = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    input.dispatchEvent(enterEvent);
    expect(enterEvent.defaultPrevented).toBe(true);
    expect(onApply).toHaveBeenCalledTimes(1);

    const tabEvent = new KeyboardEvent("keydown", { key: "Tab", cancelable: true });
    input.dispatchEvent(tabEvent);
    expect(onApply).toHaveBeenCalledTimes(1);

    button.click();
    expect(onApply).toHaveBeenCalledTimes(2);

    dispose();

    const secondEnter = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    input.dispatchEvent(secondEnter);
    button.click();
    expect(onApply).toHaveBeenCalledTimes(2);
  });
});
