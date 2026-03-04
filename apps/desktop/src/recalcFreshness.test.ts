import { describe, expect, it } from "vitest";
import {
  deriveRecalcFreshnessBadge,
  formatRecalcStatusText,
  markRecalcCompletedState,
  markRecalcPendingState,
  resetRecalcStatusState,
} from "./recalcFreshness";

describe("recalcFreshness state transitions", () => {
  it("starts with no recalc run and no pending edits", () => {
    const state = resetRecalcStatusState();
    expect(state).toEqual({
      runCount: 0,
      pendingEditsSinceLastRun: false,
      lastScope: null,
      lastAtIso: null,
    });
    expect(formatRecalcStatusText(state)).toBe("No recalc run yet for this workbook.");
    expect(deriveRecalcFreshnessBadge(state)).toEqual({
      label: "No Recalc Yet",
      tone: "pending",
    });
  });

  it("marks pending edits before first recalc", () => {
    const state = markRecalcPendingState(resetRecalcStatusState());
    expect(state.runCount).toBe(0);
    expect(state.pendingEditsSinceLastRun).toBe(true);
    expect(formatRecalcStatusText(state)).toContain("Recalc pending:");
    expect(deriveRecalcFreshnessBadge(state)).toEqual({
      label: "Pending Recalc",
      tone: "pending",
    });
  });

  it("records recalc completion and clears stale status", () => {
    const initial = markRecalcPendingState(resetRecalcStatusState());
    const recalcTime = "2026-03-02T20:00:00.000Z";
    const state = markRecalcCompletedState(initial, "Sheet1", recalcTime);

    expect(state).toEqual({
      runCount: 1,
      pendingEditsSinceLastRun: false,
      lastScope: "Sheet1",
      lastAtIso: recalcTime,
    });

    const text = formatRecalcStatusText(state, (iso) => iso);
    expect(text).toContain("Recalc runs: 1");
    expect(text).toContain("Last scope: Sheet1");
    expect(text).toContain(`Last run: ${recalcTime}`);
    expect(text).toContain("Freshness: fresh (no edits since last recalc)");
    expect(deriveRecalcFreshnessBadge(state)).toEqual({
      label: "Fresh",
      tone: "fresh",
    });
  });

  it("marks stale after additional edits and resets cleanly on workbook reset", () => {
    const recalcTime = "2026-03-02T20:00:00.000Z";
    const recalcDone = markRecalcCompletedState(
      resetRecalcStatusState(),
      "Sheet1",
      recalcTime,
    );
    const stale = markRecalcPendingState(recalcDone);

    expect(stale.runCount).toBe(1);
    expect(stale.pendingEditsSinceLastRun).toBe(true);
    expect(stale.lastScope).toBe("Sheet1");
    expect(stale.lastAtIso).toBe(recalcTime);
    expect(formatRecalcStatusText(stale, (iso) => iso)).toContain(
      "Freshness: stale (edits after last recalc)",
    );
    expect(deriveRecalcFreshnessBadge(stale)).toEqual({
      label: "Stale",
      tone: "stale",
    });

    const reset = resetRecalcStatusState();
    expect(reset).toEqual({
      runCount: 0,
      pendingEditsSinceLastRun: false,
      lastScope: null,
      lastAtIso: null,
    });
  });
});
