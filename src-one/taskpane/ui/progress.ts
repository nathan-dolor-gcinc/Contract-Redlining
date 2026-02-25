// src-one/taskpane/ui/progress.ts
//
// Manages the section progress track bar at the top of the panel.

import { setText } from "./dom";

export type ProgressState = "done" | "active" | "warn";

// ─── Build ────────────────────────────────────────────────────────────────────

/**
 * Render `total` empty segment divs into the progress track.
 * Each segment represents one CONTRACT SECTION (not one redline).
 */
export function buildProgressTrack(total: number): void {
  const track = document.getElementById("progressTrack");
  if (!track) return;

  track.innerHTML = "";
  for (let i = 0; i < total; i++) {
    const seg = document.createElement("div");
    seg.className = "progress-seg";
    seg.id = `seg-${i}`;
    track.appendChild(seg);
  }

  setText("progress-label", `0 of ${total} sections`);
}

// ─── Update ───────────────────────────────────────────────────────────────────

/** Set a segment to done / active / warn. */
export function markProgress(index: number, state: ProgressState): void {
  const seg = document.getElementById(`seg-${index}`);
  if (seg) seg.className = `progress-seg ${state}`;
}

/**
 * Update the "X of Y sections" label and the reviewed-changes count.
 * @param reviewedSections  How many sections have been fully reviewed.
 * @param totalSections     Total number of redlined sections.
 * @param reviewedChanges   Total individual tracked changes reviewed so far.
 */
export function updateProgressLabel(
  reviewedSections: number,
  totalSections: number,
  reviewedChanges?: number
): void {
  setText("stat-reviewed", String(reviewedChanges ?? reviewedSections));
  setText("progress-label", `${reviewedSections} of ${totalSections} sections`);
}