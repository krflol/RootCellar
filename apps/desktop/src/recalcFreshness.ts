export type RecalcStatusState = {
  runCount: number;
  pendingEditsSinceLastRun: boolean;
  lastScope: string | null;
  lastAtIso: string | null;
};

export type RecalcFreshnessBadge = {
  label: string;
  tone: "fresh" | "stale" | "pending";
};

export function resetRecalcStatusState(): RecalcStatusState {
  return {
    runCount: 0,
    pendingEditsSinceLastRun: false,
    lastScope: null,
    lastAtIso: null,
  };
}

export function markRecalcPendingState(state: RecalcStatusState): RecalcStatusState {
  return {
    ...state,
    pendingEditsSinceLastRun: true,
  };
}

export function markRecalcCompletedState(
  state: RecalcStatusState,
  scope: string | null,
  atIso: string,
): RecalcStatusState {
  return {
    runCount: state.runCount + 1,
    pendingEditsSinceLastRun: false,
    lastScope: scope,
    lastAtIso: atIso,
  };
}

export function formatRecalcStatusText(
  state: RecalcStatusState,
  formatTimestamp: (iso: string) => string = (iso) => new Date(iso).toLocaleString(),
): string {
  if (state.runCount === 0) {
    return state.pendingEditsSinceLastRun
      ? "Recalc pending: edits were applied since workbook load. Run \"Recalc Loaded Workbook\" to refresh dependent formulas."
      : "No recalc run yet for this workbook.";
  }

  const lines = [
    `Recalc runs: ${state.runCount}`,
    `Last scope: ${state.lastScope ?? "all loaded sheets"}`,
    `Last run: ${state.lastAtIso ? formatTimestamp(state.lastAtIso) : "unknown"}`,
    `Freshness: ${
      state.pendingEditsSinceLastRun
        ? "stale (edits after last recalc)"
        : "fresh (no edits since last recalc)"
    }`,
  ];
  return lines.join("\n");
}

export function deriveRecalcFreshnessBadge(state: RecalcStatusState): RecalcFreshnessBadge {
  if (state.runCount === 0) {
    if (state.pendingEditsSinceLastRun) {
      return { label: "Pending Recalc", tone: "pending" };
    }
    return { label: "No Recalc Yet", tone: "pending" };
  }

  if (state.pendingEditsSinceLastRun) {
    return { label: "Stale", tone: "stale" };
  }

  return { label: "Fresh", tone: "fresh" };
}
