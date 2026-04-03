// src-one/taskpane/review.ts

import { session, resetSession } from "./state/session";
import { registerAdvanceCallback, registerCommentWrittenCallback } from "./tools/dispatchTools";
import { sendPrompt, BACKEND_BASE_URL } from "./api/client";
import {
  getRedlinedSections,
  addWordCommentOnRedline,
  addCommentOnCluster,
  clusterSectionSemantically,
} from "./tools/wordTools";
import type { RedlineCluster } from "./tools/wordTools";
import { appendAssistantBubble, appendSysMsg, setLoading } from "./ui/chat";
import {
  appendScanSummary,
  appendStartReviewButton,
  appendAnalysisCard,
  setCardDisabled,
  tryParseRecommendation,
} from "./ui/cards";
import { buildProgressTrack, markProgress, updateProgressLabel } from "./ui/progress";
import { showEl, setText, getDocumentName } from "./ui/dom";

// ─── Section ordering ─────────────────────────────────────────────────────────

let _sectionOrder: Record<string, number> | null = null;

async function fetchSectionOrder(): Promise<Record<string, number>> {
  if (_sectionOrder) return _sectionOrder;
  try {
    const resp = await fetch(`${BACKEND_BASE_URL}/api/section-order`);
    const data = await resp.json() as { sections: Record<string, number> };
    _sectionOrder = data.sections;
  } catch (err) {
    console.warn("[review] Could not fetch section order — sections will not be sorted:", err);
    _sectionOrder = {};
  }
  return _sectionOrder;
}

function getSectionOrder(sectionTitle: string, order: Record<string, number>): number {
  if (!sectionTitle) return 999;
  const cleaned = sectionTitle.replace(/^\d+(\.\d+)*\.?\s*/, "").trim().toUpperCase();
  for (const [key, val] of Object.entries(order)) {
    if (key === cleaned || key.includes(cleaned) || cleaned.includes(key)) return val;
  }
  return 999;
}

// ─── Initialization & scan ────────────────────────────────────────────────────

export async function initializeScan(): Promise<void> {
  resetSession();

  registerAdvanceCallback(analyzeCurrentCluster);

  registerCommentWrittenCallback(() => {
    const cardIndex = session.currentClusterIndex;
    markProgress(cardIndex, "done");
    session.currentClusterIndex++;
    const totalChangesReviewed = session.allClusters
      .slice(0, session.currentClusterIndex)
      .reduce((sum, cl) => sum + cl.changes.length, 0);
    updateProgressLabel(session.currentClusterIndex, session.allClusters.length, totalChangesReviewed);
    setCardDisabled(cardIndex, true);
  });

  // No longer prime the thread with the full contract body — the agent
  // calls read_word_body on demand if it needs contract context. This keeps
  // the thread small and avoids timeout errors on large documents.
  const redlinedSections = await getRedlinedSections();

  appendSysMsg("Scanning contract for tracked changes…");

  // ── Sort sections into document order ─────────────────────────────────────
  const sectionOrder = await fetchSectionOrder();
  const sortedSections = [...redlinedSections].sort(
    (a, b) => getSectionOrder(a.sectionTitle, sectionOrder) - getSectionOrder(b.sectionTitle, sectionOrder)
  );
  session.redlinedSections = sortedSections;

  // ── Semantic clustering via /api/cluster (Foundry model, no agent thread) ───
  appendSysMsg("Grouping tracked changes into review clusters…");

  const clusterArrays = await Promise.all(
    sortedSections.map((s) =>
      clusterSectionSemantically(s.sectionNumber, s.sectionTitle, s.changes)
    )
  );
  session.allClusters = clusterArrays.flat();

  console.log("[clusters]", session.allClusters.map((c, i) => {
    const types = c.changes.map((ch) => `${ch.type}(${ch.author})`).join(", ");
    return `${i}: [${c.sectionNumber}] ${c.sectionTitle} (${c.changes.length} edits) — ${types}`;
  }));

  const totalChanges = session.allClusters.reduce((sum, cl) => sum + cl.changes.length, 0);

  document.getElementById("initial-loading")?.remove();

  showEl("doc-status-bar");
  showEl("progress-section");
  setText("doc-name-text", getDocumentName());
  setText("stat-total", String(totalChanges));
  setText("stat-reviewed", "0");

  buildProgressTrack(session.allClusters.length);

  if (session.allClusters.length === 0) {
    appendSysMsg("No tracked changes found in this document.");
    appendAssistantBubble(
      "This document has no tracked changes. Feel free to ask me anything about its contents."
    );
    return;
  }

  appendScanSummary(session.redlinedSections, session.allClusters);
  appendStartReviewButton(session.allClusters.length, totalChanges, startReview);
}

// ─── Cluster-by-cluster review ────────────────────────────────────────────────

async function startReview(): Promise<void> {
  if (session.allClusters.length === 0) return;
  session.currentClusterIndex = 0;
  await analyzeCurrentCluster();
}

export async function analyzeCurrentCluster(): Promise<void> {
  const clusters = session.allClusters;

  if (session.currentClusterIndex >= clusters.length) {
    appendSysMsg("All clusters reviewed.");
    appendAssistantBubble("Review complete — all redlined clusters have been assessed.");
    return;
  }

  const cluster = clusters[session.currentClusterIndex];
  markProgress(session.currentClusterIndex, "active");
  setLoading(true);

  const prompt = buildClusterPrompt(cluster, session.currentClusterIndex, clusters.length);
  const { reply, conversationId } = await sendPrompt(prompt, session.conversationId);
  session.conversationId = conversationId;
  setLoading(false);

  if (!reply) return;

  const parsed = tryParseRecommendation(reply);
  if (parsed) {
    session.lastRecommendation = {
      ...parsed,
      changeId: cluster.changes[0]?.id,
      allChangeIds: cluster.changes.map((c) => c.id),
    };
    appendAnalysisCard(
      session.lastRecommendation,
      session.currentClusterIndex,
      handleCardAction,
      cluster
    );
  } else {
    appendAssistantBubble(reply);
  }
}

// ─── Card action handler ──────────────────────────────────────────────────────

export async function handleCardAction(
  action: "accept" | "reject" | "insertAlt",
  cardIndex: number,
  altIndex?: number
): Promise<void> {
  if (!session.lastRecommendation) return;

  setCardDisabled(cardIndex, true);
  const rec = session.lastRecommendation;

  try {
    let commentText: string;
    let sysMessage: string;

    if (action === "accept") {
      commentText = `AI: ACCEPT CHANGES`;
      sysMessage  = `✓ Accept comment added to ${rec.sectionTitle}.`;
    } else if (action === "reject") {
      commentText = `AI: REJECT CHANGES`;
      sysMessage  = `✗ Reject comment added to ${rec.sectionTitle}.`;
    } else {
      const alt = rec.alternativeLanguageOptions?.[altIndex ?? 0];
      if (!alt) { setCardDisabled(cardIndex, false); return; }
      commentText = `AI Review: ALTERNATIVE — ${alt.label}: "${alt.text}"`;
      sysMessage  = `✏ Alternative ${(altIndex ?? 0) + 1} comment added to ${rec.sectionTitle}.`;
    }

    setLoading(true);

    const cluster = session.allClusters[cardIndex];
    const firstChange = cluster?.changes[0];
    const isDeletion = firstChange?.type?.toLowerCase().includes("delete");

    const result = await addCommentOnCluster({
      changeIds:        cluster?.changes.map((c) => c.id) ?? [],
      commentText,
      firstChangeText:  isDeletion ? "" : (firstChange?.text ?? ""),
      paragraphContext: cluster?.paragraphText ?? firstChange?.paragraphContext ?? "",
      sectionTitle:     rec.sectionTitle,
      sectionNumber:    rec.sectionNumber ?? "",
    });

    setLoading(false);

    if (!result.ok) {
      throw new Error(`Could not anchor comment in "${rec.sectionTitle}". Error: ${result.error ?? "unknown"}`);
    }

    appendSysMsg(sysMessage);
    markProgress(cardIndex, "done");

    const nextIndex = cardIndex + 1;
    session.currentClusterIndex = nextIndex;

    const totalChangesReviewed = session.allClusters
      .slice(0, nextIndex)
      .reduce((sum, cl) => sum + cl.changes.length, 0);
    updateProgressLabel(nextIndex, session.allClusters.length, totalChangesReviewed);

    const remaining = session.allClusters.length - nextIndex;
    const actionLabel =
      action === "accept" ? "accepted" :
      action === "reject" ? "rejected" :
      `inserted Alternative ${String.fromCharCode(65 + (altIndex ?? 0))}`;

    const notifyPrompt =
      `[SYSTEM] The user ${actionLabel} the changes in section "${rec.sectionTitle}". ` +
      `The Word comment has already been written — do NOT call add_word_comment again. ` +
      `Respond with a SHORT plain-English confirmation only — do NOT output JSON. ` +
      (remaining > 0
        ? `There are ${remaining} cluster(s) remaining. Confirm what was done in one sentence and ask if they want to continue.`
        : `That was the last cluster. Confirm and tell them the review is complete.`);

    setLoading(true);
    const { reply: confirmReply, conversationId } = await sendPrompt(notifyPrompt, session.conversationId);
    session.conversationId = conversationId;
    setLoading(false);

    if (confirmReply) {
      const looksLikeJson = confirmReply.trimStart().startsWith("{") || confirmReply.trimStart().startsWith("```");
      if (looksLikeJson) {
        console.warn("[review] notifyPrompt response was JSON — suppressing and using fallback");
        const fallback = remaining > 0
          ? `Changes in "${rec.sectionTitle}" ${actionLabel}. Ready for the next cluster when you are.`
          : `Changes in "${rec.sectionTitle}" ${actionLabel}. Review complete — all clusters have been assessed.`;
        appendAssistantBubble(fallback);
      } else {
        appendAssistantBubble(confirmReply);
      }
    }

  } catch (err) {
    console.error("[review] Action error:", err);
    appendAssistantBubble(`⚠ Action failed: ${(err as Error).message}`);
    setCardDisabled(cardIndex, false);
    setLoading(false);
  }
}

// ─── Prompt builders ──────────────────────────────────────────────────────────

function buildClusterPrompt(cluster: RedlineCluster, index: number, total: number): string {
  const changeList = cluster.changes
    .map((c, i) => `  Edit ${i + 1}: [${c.type} by ${c.author}] "${c.text?.slice(0, 120) ?? "unknown"}"`)
    .join("\n");

  return (
    `[SYSTEM] Analyse the following tracked-change cluster and respond with ONLY a raw JSON object ` +
    `(no markdown fences, no preamble). Use the analysis response format from your instructions.\n\n` +
    `Cluster ${index + 1} of ${total}:\n` +
    `Section: ${cluster.sectionTitle}\n` +
    `Paragraph: ${cluster.paragraphText.slice(0, 400)}\n\n` +
    `This cluster contains ${cluster.changes.length} edit${cluster.changes.length !== 1 ? "s" : ""} ` +
    `to the same paragraph:\n` +
    changeList
  );
}