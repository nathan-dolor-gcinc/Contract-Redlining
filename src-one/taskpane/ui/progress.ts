// src-one/taskpane/ui/progress.ts

import { setText } from "./dom";

export type ProgressState = "done" | "active" | "warn";

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
  setText("progress-label", `0 of ${total} clusters`);
}

export function markProgress(index: number, state: ProgressState): void {
  const seg = document.getElementById(`seg-${index}`);
  if (seg) seg.className = `progress-seg ${state}`;
}

export function updateProgressLabel(
  reviewedClusters: number,
  totalClusters: number,
  reviewedChanges?: number
): void {
  setText("stat-reviewed", String(reviewedChanges ?? reviewedClusters));
  setText("progress-label", `${reviewedClusters} of ${totalClusters} clusters`);
}