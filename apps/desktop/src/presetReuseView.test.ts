import { describe, expect, it } from "vitest";
import { renderLastCustomPresetButtonView } from "./presetReuseView";

describe("renderLastCustomPresetButtonView", () => {
  it("renders default disabled state when no dimensions are available", () => {
    const button = document.createElement("button");

    renderLastCustomPresetButtonView({
      buttonEl: button,
      dimensions: null,
    });

    expect(button.textContent).toBe("Apply Last Custom");
    expect(button.disabled).toBe(true);
  });

  it("renders enabled dimensions label when a preset exists", () => {
    const button = document.createElement("button");

    renderLastCustomPresetButtonView({
      buttonEl: button,
      dimensions: { rows: 3, cols: 4 },
    });

    expect(button.textContent).toBe("Apply Last 3x4");
    expect(button.disabled).toBe(false);
  });
});
