// src-one/taskpane/state/session.ts

import type { RedlinedSection, RedlineCluster } from "../tools/wordTools";
import type { AnalysisRecommendation } from "../ui/cards";

interface SessionState {
  conversationId: string | null;
  redlinedSections: RedlinedSection[];
  allClusters: RedlineCluster[];
  currentClusterIndex: number;
  lastRecommendation: (AnalysisRecommendation & { allChangeIds?: string[] }) | null;
  sessionEnded: boolean;
}

export const session: SessionState = {
  conversationId: null,
  redlinedSections: [],
  allClusters: [],
  currentClusterIndex: 0,
  lastRecommendation: null,
  sessionEnded: false,
};

// FIX: Resets all session state so the scan can be re-run when the task pane
// is reopened without a full page reload (Office.js does not always reload the
// page between document opens).
export function resetSession(): void {
  session.conversationId = null;
  session.redlinedSections = [];
  session.allClusters = [];
  session.currentClusterIndex = 0;
  session.lastRecommendation = null;
  session.sessionEnded = false;
}