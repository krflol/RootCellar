import {
  deriveRecalcFreshnessBadge,
  formatRecalcStatusText,
  type RecalcStatusState,
} from "./recalcFreshness";

type RenderRecalcStatusViewArgs = {
  state: RecalcStatusState;
  badgeEl: HTMLElement;
  statusEl: HTMLElement;
  formatTimestamp?: (iso: string) => string;
};

export function renderRecalcStatusView({
  state,
  badgeEl,
  statusEl,
  formatTimestamp,
}: RenderRecalcStatusViewArgs): void {
  const badge = deriveRecalcFreshnessBadge(state);
  badgeEl.textContent = badge.label;
  badgeEl.className = `freshness-badge ${badge.tone}`;
  statusEl.textContent = formatRecalcStatusText(state, formatTimestamp);
}
