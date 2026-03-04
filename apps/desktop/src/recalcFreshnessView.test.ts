import { describe, expect, it } from "vitest";
import {
  markRecalcCompletedState,
  markRecalcPendingState,
  resetRecalcStatusState,
} from "./recalcFreshness";
import { renderRecalcStatusView } from "./recalcFreshnessView";

function createElements(): { badgeEl: HTMLSpanElement; statusEl: HTMLPreElement } {
  return {
    badgeEl: document.createElement("span"),
    statusEl: document.createElement("pre"),
  };
}

describe("renderRecalcStatusView", () => {
  it("renders no-run state with pending-tone badge", () => {
    const { badgeEl, statusEl } = createElements();

    renderRecalcStatusView({
      state: resetRecalcStatusState(),
      badgeEl,
      statusEl,
    });

    expect(badgeEl.textContent).toBe("No Recalc Yet");
    expect(badgeEl.className).toBe("freshness-badge pending");
    expect(statusEl.textContent).toBe("No recalc run yet for this workbook.");
  });

  it("renders fresh state with fresh badge and status details", () => {
    const { badgeEl, statusEl } = createElements();
    const recalcIso = "2026-03-02T20:00:00.000Z";
    const state = markRecalcCompletedState(resetRecalcStatusState(), "Sheet1", recalcIso);

    renderRecalcStatusView({
      state,
      badgeEl,
      statusEl,
      formatTimestamp: (iso) => iso,
    });

    expect(badgeEl.textContent).toBe("Fresh");
    expect(badgeEl.className).toBe("freshness-badge fresh");
    expect(statusEl.textContent).toContain("Last scope: Sheet1");
    expect(statusEl.textContent).toContain(`Last run: ${recalcIso}`);
    expect(statusEl.textContent).toContain("Freshness: fresh (no edits since last recalc)");
  });

  it("renders stale state with stale badge after post-recalc edit", () => {
    const { badgeEl, statusEl } = createElements();
    const completed = markRecalcCompletedState(
      resetRecalcStatusState(),
      "Sheet1",
      "2026-03-02T20:00:00.000Z",
    );
    const stale = markRecalcPendingState(completed);

    renderRecalcStatusView({
      state: stale,
      badgeEl,
      statusEl,
    });

    expect(badgeEl.textContent).toBe("Stale");
    expect(badgeEl.className).toBe("freshness-badge stale");
    expect(statusEl.textContent).toContain("Freshness: stale (edits after last recalc)");
  });
});
