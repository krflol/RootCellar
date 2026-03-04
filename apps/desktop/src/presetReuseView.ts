import type { PresetDimensions } from "./presetReuse";

type RenderLastCustomPresetButtonArgs = {
  buttonEl: HTMLButtonElement;
  dimensions: PresetDimensions | null;
};

export function renderLastCustomPresetButtonView({
  buttonEl,
  dimensions,
}: RenderLastCustomPresetButtonArgs): void {
  if (!dimensions) {
    buttonEl.textContent = "Apply Last Custom";
    buttonEl.disabled = true;
    return;
  }

  buttonEl.textContent = `Apply Last ${dimensions.rows}x${dimensions.cols}`;
  buttonEl.disabled = false;
}
