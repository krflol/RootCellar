import { describe, expect, it } from "vitest";
import {
  isValidPresetDimensions,
  parsePresetDimensions,
  serializePresetDimensions,
} from "./presetReuse";

describe("presetReuse", () => {
  it("validates dimensions against integer and bounds requirements", () => {
    expect(isValidPresetDimensions({ rows: 2, cols: 3 }, 1_048_576, 16_384)).toBe(true);
    expect(isValidPresetDimensions({ rows: 0, cols: 3 }, 1_048_576, 16_384)).toBe(false);
    expect(isValidPresetDimensions({ rows: 2.5, cols: 3 }, 1_048_576, 16_384)).toBe(false);
    expect(isValidPresetDimensions({ rows: 2, cols: 20_000 }, 1_048_576, 16_384)).toBe(false);
  });

  it("serializes and parses valid preset dimensions", () => {
    const serialized = serializePresetDimensions({ rows: 4, cols: 5 });
    expect(parsePresetDimensions(serialized, 1_048_576, 16_384)).toEqual({
      rows: 4,
      cols: 5,
    });
  });

  it("rejects malformed or out-of-range stored values", () => {
    expect(parsePresetDimensions(null, 1_048_576, 16_384)).toBeNull();
    expect(parsePresetDimensions("{bad-json", 1_048_576, 16_384)).toBeNull();
    expect(parsePresetDimensions(JSON.stringify({ rows: 2, cols: "3" }), 1_048_576, 16_384)).toBeNull();
    expect(parsePresetDimensions(JSON.stringify({ rows: 2, cols: 20_000 }), 1_048_576, 16_384)).toBeNull();
  });
});
