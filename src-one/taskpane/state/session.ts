// src-one/taskpane/state/session.ts
//
// Single source of truth for all mutable runtime state.

import type { RedlinedSection } from "../tools/wordTools";
import type { AnalysisRecommendation } from "../ui/cards";

interface SessionState {
  /** Azure thread ID — null until the first message is sent. */
  conversationId: string | null;

  /**
   * Contract sections that contain tracked changes, in document order.
   * Each section groups all its redlines together.
   */
  redlinedSections: RedlinedSection[];

  /** Index of the section currently being reviewed. */
  currentSectionIndex: number;

  /** The last recommendation card shown (used by action handlers). */
  lastRecommendation: (AnalysisRecommendation & { allChangeIds?: string[] }) | null;

  /** Set to true after "end session" so the UI can block further input. */
  sessionEnded: boolean;
}

export const session: SessionState = {
  conversationId: null,
  redlinedSections: [],
  currentSectionIndex: 0,
  lastRecommendation: null,
  sessionEnded: false,
};