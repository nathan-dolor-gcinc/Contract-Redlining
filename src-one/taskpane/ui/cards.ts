// src-one/taskpane/ui/cards.ts

import { esc, scrollChat } from "./dom";
import { buildInlineDiff, extractSnippetClientSide, extractCenteredWindow } from "../tools/wordTools";
import type { RedlinedSection, RedlineCluster, TrackedChangeInfo } from "../tools/wordTools";

// ─── Types ────────────────────────────────────────────────────────────────────

export interface AnalysisRecommendation {
  sectionTitle: string;
  sectionNumber?: string;
  originalText: string;
  proposedText: string;
  recommendation: "accept" | "reject" | "review";
  riskLevel?: "low" | "medium" | "high";
  reasoning: string;
  alternativeLanguageOptions?: Array<{ id: string; label: string; text: string }>;
  commentDraft?: string;
  changeId?: string;
}

export type CardAction = "accept" | "reject" | "insertAlt" | "followup";
export type CardActionHandler = (action: CardAction, cardIndex: number, altIndex?: number) => void;

// ─── Scan summary ─────────────────────────────────────────────────────────────

export function appendScanSummary(
  sections: RedlinedSection[],
  allClusters: RedlineCluster[]
): void {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const totalChanges = sections.reduce((sum, s) => sum + s.changes.length, 0);

  const clustersBySection = new Map<string, RedlineCluster[]>();
  for (const cl of allClusters) {
    const key = cl.sectionNumber || "PREAMBLE";
    if (!clustersBySection.has(key)) clustersBySection.set(key, []);
    clustersBySection.get(key)!.push(cl);
  }

  const visibleSections = sections.filter((s) => {
    if (!s.sectionNumber || s.sectionTitle === "PREAMBLE") return false;
    const clusters = clustersBySection.get(s.sectionNumber) ?? [];
    return clusters.length > 0;
  });

  const msg = document.createElement("div");
  msg.className = "chat-message";
  msg.innerHTML = `
<div class="chat-label">LexAI · Document Scan</div>
<div class="scan-card">
  <div class="scan-card__intro">
    Scan complete — ${totalChanges} tracked change${totalChanges !== 1 ? "s" : ""}
    grouped into ${allClusters.length} cluster${allClusters.length !== 1 ? "s" : ""}
    across ${visibleSections.length} section${visibleSections.length !== 1 ? "s" : ""}.
  </div>
  ${visibleSections.map((s) => {
    const clusters = clustersBySection.get(s.sectionNumber) ?? [];
    return `
  <div class="scan-row">
    <div class="scan-row__left">
      <span class="section-num">§ ${esc(s.sectionNumber)}</span>
      <span class="clause">${esc(s.sectionTitle.slice(0, 60))}</span>
    </div>
    <span class="status-badge ${clusters.length > 1 ? "status-badge--flag" : "status-badge--warn"}">
      ${clusters.length} cluster${clusters.length !== 1 ? "s" : ""}
      · ${s.changes.length} edit${s.changes.length !== 1 ? "s" : ""}
    </span>
  </div>
  ${clusters.map((cl, i) => `
  <div class="scan-cluster-row">
    <span class="cluster-num">Cluster ${i + 1}</span>
    <span class="cluster-preview">${esc(cl.paragraphText.slice(0, 80))}…</span>
    <span class="cluster-badge">${cl.changes.length} edit${cl.changes.length !== 1 ? "s" : ""}</span>
  </div>`).join("")}`;
  }).join("")}
</div>`;
  chat.appendChild(msg);
  scrollChat();
}

// ─── Start review button ──────────────────────────────────────────────────────

export function appendStartReviewButton(
  clusterCount: number,
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
  btn.textContent = `Start Review — ${clusterCount} cluster${clusterCount !== 1 ? "s" : ""} (${totalRedlines} edits) →`;
  btn.onclick = () => {
    btn.disabled = true;
    btn.style.opacity = "0.5";
    onStart();
  };

  wrapper.appendChild(btn);
  chat.appendChild(wrapper);
  scrollChat();
}

// ─── Next cluster button ──────────────────────────────────────────────────────

export function appendNextClusterButton(
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
  btn.innerHTML = `Analyse next cluster <span>Cluster ${currentIndex + 1} of ${total} →</span>`;
  btn.onclick = () => { btn.disabled = true; onNext(); };

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
  cluster?: RedlineCluster
): void {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const displayTitle  = cluster?.sectionTitle  ?? rec.sectionTitle;
  const displayNumber = cluster?.sectionNumber ?? rec.sectionNumber;

  const recClass = `rec-badge--${rec.recommendation}`;
  const recLabel = rec.recommendation.toUpperCase();

  const diffHtml = cluster
    ? buildClusterDiffHtml(cluster)
    : `<div class="redline-diff">
        <span class="redline-del">${esc(rec.originalText)}</span>
        <span class="redline-ins">${esc(rec.proposedText)}</span>
       </div>`;

  const altsHtml = (rec.alternativeLanguageOptions ?? []).map((a) => `
<div class="alt-item">
  <span class="alt-num">${esc(a.id)}</span>
  <span class="alt-text"><strong>${esc(a.label)}:</strong> ${esc(a.text)}</span>
</div>`).join("");

  const altButtons = (rec.alternativeLanguageOptions ?? []).map((_, i) => `
<button class="btn btn--secondary" data-action="insertAlt" data-alt-index="${i}" data-card-index="${cardIndex}">
  Insert Alt ${i + 1}
</button>`).join("");

  const wrapper = document.createElement("div");
  wrapper.className = "chat-message";
  wrapper.innerHTML = `
<div class="chat-label">LexAI · Section Analysis</div>
<div class="analysis-card">
  <div class="analysis-card__header">
    <div class="analysis-card__title">
      ${displayNumber ? `<span class="section-num">§ ${esc(displayNumber)}</span>` : ""}
      ${esc(displayTitle)}
    </div>
    <div class="rec-badge ${recClass}">${recLabel}</div>
  </div>

  ${cluster ? `<div class="section-change-count">${cluster.changes.length} edit${cluster.changes.length !== 1 ? "s" : ""} in this cluster</div>` : ""}

  <div class="analysis-card__body">
    <div class="inline-diff-block">
      <div class="inline-diff-label">
        Tracked changes in context
        <span class="inline-diff-legend">
          <del class="redline-del legend-sample">deletion</del>
          <ins class="redline-ins legend-sample">insertion</ins>
        </span>
      </div>
      ${diffHtml}
    </div>

    <div class="reason-block">
      <div class="reason-label">Analysis</div>
      ${esc(rec.reasoning)}
    </div>

    ${altsHtml ? `<div class="alternatives"><div class="alt-label">Suggested Alternatives</div>${altsHtml}</div>` : ""}
  </div>

  <div class="action-row" id="card-actions-${cardIndex}">
    <button class="btn btn--reject" data-action="reject" data-card-index="${cardIndex}">✗ Reject</button>
    <button class="btn btn--accept" data-action="accept" data-card-index="${cardIndex}">✓ Accept</button>
    ${altButtons}
    <button class="btn btn--secondary btn--full" data-action="followup" data-card-index="${cardIndex}">Ask Follow-up</button>
  </div>
</div>`;

  wrapper.querySelectorAll<HTMLButtonElement>("[data-action]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const action = btn.dataset.action as CardAction;
      const idx = Number(btn.dataset.cardIndex);
      const altIdx = btn.dataset.altIndex !== undefined ? Number(btn.dataset.altIndex) : undefined;
      if (action === "followup") { document.getElementById("agent-input")?.focus(); return; }
      onAction(action, idx, altIdx);
    });
  });

  chat.appendChild(wrapper);
  scrollChat();
}

// ─── Inline diff helpers ──────────────────────────────────────────────────────

// Renders a single unified inline diff for all changes in a cluster.
//
// Changes are grouped by their underlying paragraph (using rawParagraphText
// so deletions — whose text is absent from paragraph.text — are still
// correctly co-located with insertions from the same sentence).
//
// For each paragraph group:
//   1. A union window is derived that covers every change's text, giving
//      buildInlineDiff a single string where ALL marks can be found.
//   2. buildInlineDiff is called ONCE with ALL changes for that paragraph,
//      producing one continuous block with interleaved <del>/<ins> marks.
//
// This replaces the old per-change grouping that produced a separate block
// for every individual change — which was confusing for users.
function buildClusterDiffHtml(cluster: RedlineCluster): string {
  // ── Group changes by their raw (unmodified) paragraph ────────────────────
  // rawParagraphText is the full paragraph without deletion augmentation,
  // so insertions and deletions from the same sentence share the same key.
  const byParagraph = new Map<string, { rawPara: string; changes: typeof cluster.changes }>();

  for (const change of cluster.changes) {
    const key = change.rawParagraphText ?? change.paragraphContext ?? "";
    if (!byParagraph.has(key)) {
      byParagraph.set(key, { rawPara: key, changes: [] });
    }
    byParagraph.get(key)!.changes.push(change);
  }

  return Array.from(byParagraph.values()).map((block) => {
    const { rawPara, changes } = block;

    // ── Build a unified snippet covering ALL changes in this paragraph ──────
    //
    // For each change find its position in rawPara (insertions are present;
    // deletions are not — we append them temporarily for position finding).
    // Then take the union of all windows to get one continuous display region.
    const paraWithDeletions = (() => {
      let augmented = rawPara;
      for (const c of changes) {
        const isDel = (c.type ?? "").toLowerCase().includes("delet");
        if (isDel && c.text.trim()) {
          const norm = (s: string) => s.replace(/[‘’]/g, "'")
            .replace(/[“”]/g, '"').replace(/[–—]/g, "-")
            .replace(/ /g, " ").replace(/\s+/g, " ").toLowerCase().trim();
          if (!norm(augmented).includes(norm(c.text.slice(0, 40)))) {
            augmented = augmented + " " + c.text;
          }
        }
      }
      return augmented;
    })();

    // Expand individual windows and take their union
    const windows = changes.map((c) => extractCenteredWindow(paraWithDeletions, c.text, 300));
    const unifiedSnippet = (() => {
      if (windows.length === 0) return paraWithDeletions.slice(0, 800);
      // Find earliest start and latest end among all windows (by searching their
      // text in the augmented paragraph, falling back to the longest window)
      let earliest = paraWithDeletions.length;
      let latest = 0;
      for (const w of windows) {
        if (!w) continue;
        const pos = paraWithDeletions.indexOf(w.slice(0, 30));
        if (pos !== -1) {
          earliest = Math.min(earliest, pos);
          latest   = Math.max(latest, pos + w.length);
        }
      }
      if (earliest >= latest) {
        // Fallback: just use the longest window
        return windows.reduce((a, b) => (b.length > a.length ? b : a), "");
      }
      // Snap to sentence boundaries
      const beforeCut = paraWithDeletions.lastIndexOf(". ", earliest);
      const afterCut  = paraWithDeletions.indexOf(". ", latest);
      const start = beforeCut !== -1 && beforeCut >= earliest - 200 ? beforeCut + 2 : earliest;
      const end   = afterCut  !== -1 && afterCut  <= latest  + 200 ? afterCut  + 1 : latest;
      return paraWithDeletions.slice(start, end).trim();
    })();

    const authors = [...new Set(changes.map((c) => c.author).filter(Boolean))];
    const authorLine = authors.length ? `by ${authors.join(", ")}` : "";

    // Single buildInlineDiff call with ALL changes → one unified diff block
    const diffHtml = buildInlineDiff(unifiedSnippet, changes);

    return `
<div class="diff-paragraph">
  ${authorLine ? `<div class="diff-paragraph__meta">${esc(authorLine)}</div>` : ""}
  <div class="diff-paragraph__content">${diffHtml}</div>
</div>`;
  }).join("");
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
    .replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/```\s*$/i, "").trim();
  if (!cleaned.startsWith("{")) return null;
  try {
    const obj = JSON.parse(cleaned) as AnalysisRecommendation;
    if (obj.recommendation && obj.reasoning) return obj;
    return null;
  } catch { return null; }
}