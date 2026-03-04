export const PRESET_DIMENSIONS_STORAGE_KEY =
  "rootcellar.desktop.last_custom_preset_dimensions.v1";

export type PresetDimensions = {
  rows: number;
  cols: number;
};

export function isValidPresetDimensions(
  dimensions: PresetDimensions,
  maxRows: number,
  maxCols: number,
): boolean {
  return (
    Number.isInteger(dimensions.rows) &&
    Number.isInteger(dimensions.cols) &&
    dimensions.rows >= 1 &&
    dimensions.cols >= 1 &&
    dimensions.rows <= maxRows &&
    dimensions.cols <= maxCols
  );
}

export function serializePresetDimensions(dimensions: PresetDimensions): string {
  return JSON.stringify({
    rows: dimensions.rows,
    cols: dimensions.cols,
  });
}

export function parsePresetDimensions(
  raw: string | null,
  maxRows: number,
  maxCols: number,
): PresetDimensions | null {
  if (!raw) {
    return null;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }

  if (!parsed || typeof parsed !== "object") {
    return null;
  }

  const candidate = parsed as {
    rows?: unknown;
    cols?: unknown;
  };
  if (typeof candidate.rows !== "number" || typeof candidate.cols !== "number") {
    return null;
  }

  const dimensions: PresetDimensions = {
    rows: candidate.rows,
    cols: candidate.cols,
  };
  return isValidPresetDimensions(dimensions, maxRows, maxCols) ? dimensions : null;
}
