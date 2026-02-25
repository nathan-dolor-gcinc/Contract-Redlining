// src-one/taskpane/review.ts

import { session } from "./state/session";
import { sendPrompt, primeThread } from "./api/client";
import { getRedlinedSections, readWordBodyText, addWordCommentOnRedline } from "./tools/wordTools";
import type { RedlinedSection } from "./tools/wordTools";
import { appendAssistantBubble, appendSysMsg, setLoading } from "./ui/chat";
import {
  appendScanSummary,
  appendStartReviewButton,
  appendNextSectionButton,
  appendAnalysisCard,
  setCardDisabled,
  tryParseRecommendation,
} from "./ui/cards";
import { buildProgressTrack, markProgress, updateProgressLabel } from "./ui/progress";
import { showEl, setText, getDocumentName } from "./ui/dom";

// ─── Initialization & scan ────────────────────────────────────────────────────

// Guard: Office.onReady can fire multiple times during hot-reload / iframe reuse.
// We only want to run the scan once per add-in lifetime.
let scanStarted = false;

export async function initializeScan(): Promise<void> {
  if (scanStarted) {
    console.warn("[initializeScan] Already started — skipping duplicate call.");
    return;
  }
  scanStarted = true;

  const [redlinedSections, bodyText] = await Promise.all([
    getRedlinedSections(),
    readWordBodyText(40_000),
  ]);
  session.redlinedSections = redlinedSections;

  appendSysMsg("Reading contract and initialising agent…");

  // primeThread adds the document text to a fresh thread WITHOUT starting a run.
  // This means the thread is immediately idle and the first /api/chat call
  // (the section analysis) will never hit a "run active" race condition.
  try {
    const conversationId = await primeThread(
      `The following is the full text of the contract document the user will be reviewing. ` +
        `Use it to answer questions about the contract.\n\n` +
        `--- DOCUMENT START ---\n${bodyText}\n--- DOCUMENT END ---`
    );
    if (conversationId) {
      session.conversationId = conversationId;
      console.log("[initializeScan] Thread primed:", conversationId);
    }
  } catch (err) {
    console.warn("[initializeScan] Could not prime thread:", err);
  }

  const totalChanges = session.redlinedSections.reduce((sum, s) => sum + s.changes.length, 0);

  document.getElementById("initial-loading")?.remove();

  showEl("doc-status-bar");
  showEl("progress-section");
  setText("doc-name-text", getDocumentName());
  setText("stat-total", String(totalChanges));
  setText("stat-reviewed", "0");

  buildProgressTrack(session.redlinedSections.length);

  if (session.redlinedSections.length === 0) {
    appendSysMsg("No tracked changes found in this document.");
    appendAssistantBubble(
      "This document has no tracked changes. Feel free to ask me anything about its contents."
    );
    return;
  }

  appendScanSummary(session.redlinedSections);
  appendStartReviewButton(session.redlinedSections.length, totalChanges, startReview);
}

// ─── Section-by-section review ────────────────────────────────────────────────

async function startReview(): Promise<void> {
  if (session.redlinedSections.length === 0) return;
  session.currentSectionIndex = 0;
  await analyzeCurrentSection();
}

export async function analyzeCurrentSection(): Promise<void> {
  const sections = session.redlinedSections;

  if (session.currentSectionIndex >= sections.length) {
    appendSysMsg("All sections reviewed.");
    appendAssistantBubble("Review complete — all redlined sections have been assessed.");
    return;
  }

  const section = sections[session.currentSectionIndex];
  markProgress(session.currentSectionIndex, "active");
  setLoading(true);

  const prompt = buildSectionPrompt(section, session.currentSectionIndex, sections.length);
  const { reply, conversationId } = await sendPrompt(prompt, session.conversationId);
  session.conversationId = conversationId;
  setLoading(false);

  if (!reply) return;

  const parsed = tryParseRecommendation(reply);
  if (parsed) {
    session.lastRecommendation = {
      ...parsed,
      changeId: section.changes[0]?.id,
      allChangeIds: section.changes.map((c) => c.id),
    };
    appendAnalysisCard(session.lastRecommendation, session.currentSectionIndex, handleCardAction);
  } else {
    appendAssistantBubble(reply);
    appendNextSectionButton(session.currentSectionIndex, sections.length, analyzeCurrentSection);
  }
}

// ─── Card action handler ──────────────────────────────────────────────────────

async function handleCardAction(
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
      commentText = `AI Review: ACCEPT — ${rec.commentDraft ?? rec.reasoning}`;
      sysMessage  = `✓ Accept comment added to ${rec.sectionTitle}.`;

    } else if (action === "reject") {
      commentText = `AI Review: REJECT — ${rec.reasoning}`;
      sysMessage  = `✗ Reject comment added to ${rec.sectionTitle}.`;

    } else {
      const alt = rec.alternativeLanguageOptions?.[altIndex ?? 0];
      if (!alt) {
        setCardDisabled(cardIndex, false);
        return;
      }
      commentText = `AI Review: ALTERNATIVE — ${alt.label}: "${alt.text}"`;
      sysMessage  = `✏ Alternative ${(altIndex ?? 0) + 1} comment added to ${rec.sectionTitle}.`;
    }

    setLoading(true);

    const section = session.redlinedSections[cardIndex];
    const firstChange = section?.changes[0];

    const result = await addWordCommentOnRedline({
      changeText:       firstChange?.text             ?? "",
      paragraphContext: firstChange?.paragraphContext  ?? rec.originalText ?? "",
      sectionTitle:     rec.sectionTitle,
      sectionNumber:    rec.sectionNumber ?? "",
      commentText,
    });

    setLoading(false);

    if (!result.ok) {
      throw new Error(
        `Could not anchor comment in section "${rec.sectionTitle}". ` +
          `Error: ${result.error ?? "unknown"}`
      );
    }

    console.log(`[handleCardAction] Comment inserted — matches: ${result.matches}, used index: ${result.usedIndex}`);

    appendSysMsg(sysMessage);

    markProgress(cardIndex, "done");
    session.currentSectionIndex++;

    const totalChangesReviewed = session.redlinedSections
      .slice(0, session.currentSectionIndex)
      .reduce((sum, s) => sum + s.changes.length, 0);

    updateProgressLabel(
      session.currentSectionIndex,
      session.redlinedSections.length,
      totalChangesReviewed
    );

    setTimeout(() => {
      if (session.currentSectionIndex < session.redlinedSections.length) {
        appendNextSectionButton(
          session.currentSectionIndex,
          session.redlinedSections.length,
          analyzeCurrentSection
        );
      } else {
        appendSysMsg("All sections reviewed.");
        appendAssistantBubble(
          "Review complete — all redlined sections have been assessed. Feel free to ask any follow-up questions."
        );
      }
    }, 600);

  } catch (err) {
    console.error("[review] Action error:", err);
    appendAssistantBubble(`⚠ Action failed: ${(err as Error).message}`);
    setCardDisabled(cardIndex, false);
    setLoading(false);
  }
}

// ─── Prompt builder ───────────────────────────────────────────────────────────

function buildSectionPrompt(section: RedlinedSection, index: number, total: number): string {
  const changeList = section.changes
    .map(
      (c, i) =>
        `  Redline ${i + 1}: [${c.type} by ${c.author}] "${c.text?.slice(0, 120) ?? "unknown"}"`
    )
    .join("\n");

  return (
    `Please analyse Section ${index + 1} of ${total}:\n` +
    `Section: ${section.sectionTitle}\n` +
    `Section context: ${section.sectionContext?.slice(0, 400) ?? ""}\n\n` +
    `This section contains ${section.changes.length} tracked change${section.changes.length !== 1 ? "s" : ""}:\n` +
    changeList
  );
}