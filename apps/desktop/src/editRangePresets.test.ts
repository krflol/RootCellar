import { describe, expect, it } from "vitest";
import { buildPresetRange, parseSingleA1Cell } from "./editRangePresets";

describe("parseSingleA1Cell", () => {
  it("parses single A1 refs", () => {
    expect(parseSingleA1Cell("B3")).toEqual({ row: 3, col: 2 });
    expect(parseSingleA1Cell(" xfd1048576 ")).toEqual({ row: 1_048_576, col: 16_384 });
  });

  it("rejects ranges and invalid refs", () => {
    expect(parseSingleA1Cell("A1:B2")).toBeNull();
    expect(parseSingleA1Cell("0A")).toBeNull();
    expect(parseSingleA1Cell("A0")).toBeNull();
  });
});

describe("buildPresetRange", () => {
  it("builds row and block presets", () => {
    expect(buildPresetRange("A1", 1, 3)).toEqual({
      ok: true,
      anchor: { row: 1, col: 1 },
      range: "A1:C1",
    });
    expect(buildPresetRange("C3", 2, 2)).toEqual({
      ok: true,
      anchor: { row: 3, col: 3 },
      range: "C3:D4",
    });
  });

  it("rejects invalid anchors and spans", () => {
    expect(buildPresetRange("A1:B2", 2, 2)).toEqual({ ok: false, reason: "invalid_anchor" });
    expect(buildPresetRange("A1", 0, 2)).toEqual({ ok: false, reason: "invalid_span" });
    expect(buildPresetRange("A1", 2, 0)).toEqual({ ok: false, reason: "invalid_span" });
  });

  it("rejects ranges that exceed Excel bounds", () => {
    expect(buildPresetRange("XFD1048576", 1, 2)).toEqual({
      ok: false,
      reason: "out_of_bounds",
    });
    expect(buildPresetRange("XFD1048576", 2, 1)).toEqual({
      ok: false,
      reason: "out_of_bounds",
    });
  });
});
