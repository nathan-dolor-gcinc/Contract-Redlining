// src-one/taskpane/ui/cards.ts
//
// Builds the richer card-style UI elements.
// Updated to work with RedlinedSection instead of flat TrackedChangeInfo.

import { esc, scrollChat } from "./dom";
import type { RedlinedSection } from "../tools/wordTools";

// ─── Types ────────────────────────────────────────────────────────────────────

export interface AnalysisRecommendation {
  sectionTitle: string;
  sectionNumber?: string;
  originalText: string;
  proposedText: string;
  recommendation: "accept" | "reject" | "review";
  riskLevel?: "low" | "medium" | "high";
  reasoning: string;
  marketContext?: string;
  alternativeLanguageOptions?: Array<{ id: string; label: string; text: string }>;
  commentDraft?: string;
  /** Primary changeId (first change in the section). */
  changeId?: string;
}

export type CardAction = "accept" | "reject" | "insertAlt" | "followup";
export type CardActionHandler = (action: CardAction, cardIndex: number, altIndex?: number) => void;

// ─── Scan summary ─────────────────────────────────────────────────────────────

/**
 * Render the initial scan summary showing each SECTION that has redlines,
 * and how many tracked changes are within each section.
 */
export function appendScanSummary(sections: RedlinedSection[]): void {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const totalChanges = sections.reduce((sum, s) => sum + s.changes.length, 0);

  const msg = document.createElement("div");
  msg.className = "chat-message";
  msg.innerHTML = `
<div class="chat-label">LexAI · Document Scan</div>
<div class="scan-card">
  <div class="scan-card__intro">
    Scan complete — ${totalChanges} tracked change${totalChanges !== 1 ? "s" : ""} across
    ${sections.length} section${sections.length !== 1 ? "s" : ""}.
  </div>
  ${sections
    .map(
      (s) => `
  <div class="scan-row">
    <div class="scan-row__left">
      ${s.sectionNumber ? `<span class="section-num">§ ${esc(s.sectionNumber)}</span>` : ""}
      <span class="clause">${esc(s.sectionTitle.slice(0, 60))}</span>
    </div>
    <span class="status-badge ${s.changes.length > 1 ? "status-badge--flag" : "status-badge--warn"}">
      ${s.changes.length} redline${s.changes.length !== 1 ? "s" : ""}
    </span>
  </div>`
    )
    .join("")}
</div>`;
  chat.appendChild(msg);
  scrollChat();
}

// ─── Start review button ──────────────────────────────────────────────────────

export function appendStartReviewButton(
  sectionCount: number,
  totalRedlines: number,
  onStart: () => void
): void {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const wrapper = document.createElement("div");
  wrapper.className = "chat-message";
  wrapper.innerHTML = `<div class="chat-label">LexAI</div>`;

  const btn = document.createElement("button");
  btn.className = "start-review-btn";
  btn.textContent = `Start Section-by-Section Review — ${sectionCount} sections (${totalRedlines} redlines) →`;
  btn.onclick = () => {
    btn.disabled = true;
    btn.style.opacity = "0.5";
    onStart();
  };

  wrapper.appendChild(btn);
  chat.appendChild(wrapper);
  scrollChat();
}

// ─── Next section button ──────────────────────────────────────────────────────

export function appendNextSectionButton(
  currentIndex: number,
  total: number,
  onNext: () => void
): void {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  if (currentIndex >= total) return;

  const wrapper = document.createElement("div");
  wrapper.className = "chat-message";

  const label = document.createElement("div");
  label.className = "chat-label";
  label.textContent = "LexAI";

  const btn = document.createElement("button");
  btn.className = "next-section-btn";
  btn.innerHTML = `
    Analyse next section
    <span>Section ${currentIndex + 1} of ${total} →</span>`;
  btn.onclick = () => {
    btn.disabled = true;
    onNext();
  };

  wrapper.appendChild(label);
  wrapper.appendChild(btn);
  chat.appendChild(wrapper);
  scrollChat();
}

// ─── Analysis card ────────────────────────────────────────────────────────────

export function appendAnalysisCard(
  rec: AnalysisRecommendation & { allChangeIds?: string[] },
  cardIndex: number,
  onAction: CardActionHandler,
  section?: RedlinedSection
): void {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const recClass = `rec-badge--${rec.recommendation}`;
  const recLabel = rec.recommendation.toUpperCase();

  // Show each individual redline within the section
  const redlinesHtml = section
    ? section.changes
        .map(
          (c, i) => `
<div class="redline-item">
  <span class="redline-item__num">Redline ${i + 1}</span>
  <span class="redline-item__meta">${esc(c.type)} · ${esc(c.author)}</span>
  <div class="redline-diff">
    <span class="redline-del">${esc(c.text?.slice(0, 200) ?? "")}</span>
  </div>
</div>`
        )
        .join("")
    : `<div class="redline-diff">
        <span class="redline-del">${esc(rec.originalText)}</span>
        <span class="redline-ins">${esc(rec.proposedText)}</span>
       </div>`;

  const altsHtml = (rec.alternativeLanguageOptions ?? [])
    .map(
      (a) => `
<div class="alt-item">
  <span class="alt-num">${esc(a.id)}</span>
  <span class="alt-text"><strong>${esc(a.label)}:</strong> ${esc(a.text)}</span>
</div>`
    )
    .join("");

  const altButtons = (rec.alternativeLanguageOptions ?? [])
    .map(
      (_, i) => `
<button class="btn btn--secondary"
        data-action="insertAlt"
        data-alt-index="${i}"
        data-card-index="${cardIndex}">
  Insert Alt ${i + 1}
</button>`
    )
    .join("");

  const wrapper = document.createElement("div");
  wrapper.className = "chat-message";
  wrapper.innerHTML = `
<div class="chat-label">LexAI · Section Analysis</div>
<div class="analysis-card">
  <div class="analysis-card__header">
    <div class="analysis-card__title">
      ${rec.sectionNumber ? `<span class="section-num">§ ${esc(rec.sectionNumber)}</span>` : ""}
      ${esc(rec.sectionTitle)}
    </div>
    <div class="rec-badge ${recClass}">${recLabel}</div>
  </div>

  ${section ? `
  <div class="section-change-count">
    ${section.changes.length} tracked change${section.changes.length !== 1 ? "s" : ""} in this section
  </div>` : ""}

  <div class="analysis-card__body">
    ${redlinesHtml}

    <div class="reason-block">
      <div class="reason-label">Analysis</div>
      ${esc(rec.reasoning)}
    </div>

    ${rec.marketContext ? `
    <div class="reason-block">
      <div class="reason-label">Market Context</div>
      ${esc(rec.marketContext)}
    </div>` : ""}

    ${altsHtml ? `
    <div class="alternatives">
      <div class="alt-label">Suggested Alternatives</div>
      ${altsHtml}
    </div>` : ""}
  </div>

  <div class="action-row" id="card-actions-${cardIndex}">
    <button class="btn btn--reject" data-action="reject" data-card-index="${cardIndex}">✗ Reject All</button>
    <button class="btn btn--accept" data-action="accept" data-card-index="${cardIndex}">✓ Accept All</button>
    ${altButtons}
    <button class="btn btn--secondary btn--full" data-action="followup" data-card-index="${cardIndex}">
      Ask Follow-up
    </button>
  </div>
</div>`;

  // Wire up action buttons
  wrapper.querySelectorAll<HTMLButtonElement>("[data-action]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const action = btn.dataset.action as CardAction;
      const idx = Number(btn.dataset.cardIndex);
      const altIdx = btn.dataset.altIndex !== undefined ? Number(btn.dataset.altIndex) : undefined;

      if (action === "followup") {
        document.getElementById("agent-input")?.focus();
        return;
      }

      onAction(action, idx, altIdx);
    });
  });

  chat.appendChild(wrapper);
  scrollChat();
}

// ─── Card helpers ─────────────────────────────────────────────────────────────

export function setCardDisabled(cardIndex: number, disabled: boolean): void {
  document.querySelectorAll<HTMLButtonElement>(`#card-actions-${cardIndex} .btn`).forEach((b) => {
    b.disabled = disabled;
  });
}

// ─── Parse helpers ────────────────────────────────────────────────────────────

export function tryParseRecommendation(text: string): AnalysisRecommendation | null {
  const cleaned = text
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/```\s*$/i, "")
    .trim();

  if (!cleaned.startsWith("{")) return null;

  try {
    const obj = JSON.parse(cleaned) as AnalysisRecommendation;
    if (obj.recommendation && obj.reasoning) return obj;
    return null;
  } catch {
    return null;
  }
}