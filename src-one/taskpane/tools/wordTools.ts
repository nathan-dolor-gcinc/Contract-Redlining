// src-one/taskpane/tools/wordTools.ts

import { BACKEND_BASE_URL } from "../api/client";

export interface TrackedChangeInfo {
  id: string;
  type: string;
  author: string;
  date: string;
  text: string;
  sectionTitle: string;
  sectionNumber: string;
  paragraphContext: string;
  /** Full unmodified paragraph (no deletion augmentation, no slice limit).
   *  Used by buildClusterDiffHtml to group co-paragraph changes into one
   *  unified inline diff instead of a separate block per change. */
  rawParagraphText: string;
}

export interface RedlinedSection {
  sectionNumber: string;
  sectionTitle: string;
  changes: TrackedChangeInfo[];
}

export interface RedlineCluster {
  clusterId: string;
  sectionNumber: string;
  sectionTitle: string;
  paragraphText: string;
  changes: TrackedChangeInfo[];
}

export interface CommentResult {
  ok: boolean;
  matches?: number;
  usedIndex?: number;
  error?: string;
}

const SECTION_HEADING_RE = /^([A-Z][A-Z\s&,\/()'".-]{2,}[A-Z])\./;

// ─── Read ─────────────────────────────────────────────────────────────────────

export async function readWordBodyText(maxChars = 60_000): Promise<string> {
  const text = await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text ?? "";
  });
  return text.length > maxChars ? text.slice(0, maxChars) + "\n...[truncated]..." : text;
}

export async function getTrackedChanges(): Promise<TrackedChangeInfo[]> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items/text");

    const changes = body.getTrackedChanges();
    changes.load("items/id,items/type,items/date,items/author,items/text");
    await context.sync();

    const paraTexts = paragraphs.items.map((p) => p.text ?? "");
    const sectionEntries = buildSectionMap(paraTexts);

    console.group("[getTrackedChanges] Section map:", sectionEntries.length, "sections");
    for (const entry of sectionEntries) {
      console.log(`  § "${entry.sectionNumber || "PREAMBLE"}" — startParaIndex: ${entry.startParaIndex}, ${entry.paragraphTexts.length} paragraphs`);
    }
    console.groupEnd();

    // Index-based boundary mapping — no fragile text equality
    const paraIndexToSection: SectionEntry[] = new Array(paraTexts.length);
    for (let s = 0; s < sectionEntries.length; s++) {
      const start = sectionEntries[s].startParaIndex;
      const end =
        s + 1 < sectionEntries.length
          ? sectionEntries[s + 1].startParaIndex
          : paraTexts.length;
      for (let i = start; i < end; i++) {
        paraIndexToSection[i] = sectionEntries[s];
      }
    }

    console.group("[getTrackedChanges] Raw changes from Word:", changes.items.length);
    for (const c of changes.items) {
      console.log(`  [${c.type}] text: "${c.text}" | author: "${c.author}"`);
    }
    console.groupEnd();

    // ── Pass 1: attribute each change to a section using paragraph text ─────
    //
    // NOTE on deletions: Word's paragraph.text API omits tracked-deleted text
    // (it returns the post-acceptance state). This means:
    //   • findSectionByPosition / findSectionForText may FAIL for deletions
    //     because the deleted text isn't in any paraTexts[i].
    //   • Even when attribution succeeds, paragraphContext won't contain the
    //     deleted text, so extractSnippetClientSide returns the full paragraph
    //     and buildInlineDiff can't highlight the deletion.
    //
    // Pass 1 handles normal attribution. Pass 2 (below) handles the two
    // failure modes specific to deletions.

    type RawChangeEntry = {
      c: Word.TrackedChange;
      id: string;
      date: string;
      changeText: string;
      section: { sectionNumber: string; sectionTitle: string; paragraphContext: string } | null;
      isDeletion: boolean;
    };

    const rawEntries: RawChangeEntry[] = changes.items.map((c) => {
      const raw = c as unknown as Record<string, unknown>;
      const id = typeof raw["id"] === "string" ? raw["id"] : String(raw["id"] ?? "");

      const rawDate = c.date as string | Date | undefined;
      const date =
        rawDate instanceof Date ? rawDate.toISOString() :
        typeof rawDate === "string" ? rawDate : "";

      const changeText = c.text ?? "";
      const isDeletion = String(c.type ?? "").toLowerCase().includes("delet");

      console.group(`[attribution pass-1] "${changeText.slice(0, 60)}" (${c.type})`);

      let section = findSectionByPosition(changeText, paraTexts, paraIndexToSection);
      if (!section) {
        const found = findSectionForText(changeText, sectionEntries);
        if (found?.sectionNumber) section = found;
      }

      if (section) {
        console.log(`  ✅ § ${section.sectionNumber} | context: "${section.paragraphContext?.slice(0, 80)}"`);
      } else {
        console.warn(isDeletion
          ? "  ⚠ Not found in pass-1 (deletion text absent from paragraph.text) — will retry in pass-2"
          : "  ⚠ Not found — will DROP");
      }
      console.groupEnd();

      return { c, id, date, changeText, section, isDeletion };
    });

    // ── Pass 2: range-based fallback for deletions that weren't attributed ───
    //
    // For deletion changes where pass-1 failed (deleted text not in any
    // paragraph.text), load the change's Range and use the surrounding
    // paragraph's text — which IS in paragraph.text — to identify the section.

    const unattributed = rawEntries.filter((e) => !e.section && e.isDeletion);

    if (unattributed.length > 0) {
      const rangeParaCollections: Array<{ entry: RawChangeEntry; paras: Word.ParagraphCollection }> = [];

      for (const entry of unattributed) {
        try {
          const range = (entry.c as any).getRange() as Word.Range;
          const paras = range.paragraphs;
          paras.load("items/text");
          rangeParaCollections.push({ entry, paras });
        } catch {
          // getRange() not available on this Word version — will stay unattributed
        }
      }

      if (rangeParaCollections.length > 0) {
        await context.sync();

        for (const { entry, paras } of rangeParaCollections) {
          console.group(`[attribution pass-2] "${entry.changeText.slice(0, 60)}"`);

          for (const para of paras.items) {
            const paraText = (para as any).text as string ?? "";
            if (paraText.trim().length < 6) continue;

            // Find this neighbouring paragraph in our index to get its section
            const firstWords = paraText.trim().split(/\s+/).slice(0, 6).join(" ").toLowerCase();
            for (let i = 0; i < paraTexts.length; i++) {
              if (paraTexts[i].toLowerCase().includes(firstWords)) {
                const sec = paraIndexToSection[i];
                if (sec?.sectionNumber) {
                  entry.section = {
                    sectionNumber: sec.sectionNumber,
                    sectionTitle: sec.sectionTitle,
                    // full paragraph — caller will extract a centered window
                    paragraphContext: paraTexts[i],
                  };
                  console.log(`  ✅ § ${sec.sectionNumber} via range neighbour`);
                  break;
                }
              }
            }
            if (entry.section) break;
          }

          if (!entry.section) console.warn("  ⚠ DROPPING — range fallback also failed");
          console.groupEnd();
        }
      }
    }

    // ── Finalise: extract per-change centered paragraphContext ───────────────
    //
    // Each change gets its OWN context window centered on its change text.
    // This fixes two issues that the old slice(0,2000) approach caused:
    //
    //   1. All changes in the same paragraph shared the same leading-edge
    //      context — changes deep in a long paragraph (e.g. "ten percent (10%)"
    //      or the "Notwithstanding" clause in §6 PAYMENT) were past the 2000-
    //      char cutoff and so never appeared in the diff or snippet.
    //
    //   2. extractSnippetClientSide received the whole section (often starting
    //      at the heading) and picked the first sentence instead of the one
    //      containing the change, showing irrelevant text in the review card.
    //
    // For deletion changes Word's paragraph.text never includes the deleted
    // text, so we append it to the full paragraph before searching, giving
    // buildInlineDiff a string it can locate the change text inside.

    return rawEntries.flatMap((entry) => {
      const { id, date, changeText, isDeletion } = entry;
      const section = entry.section;

      if (!section || !section.sectionNumber) {
        console.warn(`[getTrackedChanges] DROPPING "${changeText.slice(0, 60)}" — no section found`);
        return [];
      }

      // Full paragraph text returned by the helpers (no slice limit).
      let fullParagraph = section.paragraphContext;

      // For deletions: deleted text is absent from paragraph.text.
      // Append it so extractCenteredWindow can anchor on it.
      if (isDeletion && changeText.trim().length > 0) {
        const normCtx  = normalizeForMatch(fullParagraph);
        const normHead = normalizeForMatch(changeText.slice(0, 60));
        if (normHead.length >= 4 && !normCtx.includes(normHead)) {
          fullParagraph = fullParagraph ? fullParagraph + " " + changeText : changeText;
        }
      }

      // Extract a ~800-char window centered on the change text.
      const paragraphContext = extractCenteredWindow(fullParagraph, changeText);

      console.log(`[getTrackedChanges] § ${section.sectionNumber} | "${changeText.slice(0, 40)}" → ctx: "${paragraphContext.slice(0, 80)}"`);

      return [{
        id,
        type: String(entry.c.type ?? ""),
        author: entry.c.author ?? "",
        date,
        text: changeText,
        sectionTitle: section.sectionTitle,
        sectionNumber: section.sectionNumber,
        paragraphContext,
        rawParagraphText: section.paragraphContext, // unmodified full paragraph for unified diff grouping
      } satisfies TrackedChangeInfo];
    });
  });
}

export async function getRedlinedSections(): Promise<RedlinedSection[]> {
  const changes = await getTrackedChanges();

  const sectionMap = new Map<string, RedlinedSection>();
  const sectionOrder: string[] = [];

  for (const change of changes) {
    const key = change.sectionNumber;
    if (!key) continue;
    if (!sectionMap.has(key)) {
      sectionMap.set(key, {
        sectionNumber: change.sectionNumber,
        sectionTitle: change.sectionTitle,
        changes: [],
      });
      sectionOrder.push(key);
    }
    sectionMap.get(key)!.changes.push(change);
  }

  const result = sectionOrder.map((key) => sectionMap.get(key)!);

  console.group("[getRedlinedSections] Sections found:", result.length);
  for (const section of result) {
    console.group(`§ ${section.sectionNumber} — ${section.changes.length} change(s)`);
    for (const change of section.changes) {
      console.log(`  [${change.type}] "${change.text?.slice(0, 80)}" | context: "${change.paragraphContext?.slice(0, 80)}"`);
    }
    console.groupEnd();
  }
  console.groupEnd();

  return result;
}

// ─── Semantic clustering ──────────────────────────────────────────────────────
//
// POSTs to /api/cluster which calls the AzureOpenAI model directly.
//
// The backend now returns:
//   { clusters: [{ indices: [0,2], snippet: "..." }, { indices: [1], snippet: "..." }] }
//
// snippet — AI-chosen 1-3 sentence verbatim substring of paragraphContext,
//           used as paragraphText on the cluster for display in the review card.

export async function clusterSectionSemantically(
  sectionNumber: string,
  sectionTitle: string,
  changes: TrackedChangeInfo[]
): Promise<RedlineCluster[]> {
  if (changes.length === 1) {
    return [{
      clusterId: `${sectionNumber}-0`,
      sectionNumber,
      sectionTitle,
      paragraphText: extractSnippetClientSide(changes[0].paragraphContext, changes),
      changes,
    }];
  }

  let clusters: Array<{ indices: number[]; snippet: string }>;

  try {
    const resp = await fetch(`${BACKEND_BASE_URL}/api/cluster`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sectionNumber, sectionTitle, changes }),
    });

    if (!resp.ok) throw new Error(`/api/cluster responded ${resp.status}`);

    const data = await resp.json() as { clusters: Array<{ indices: number[]; snippet: string }> };
    clusters = data.clusters;

    // Validate
    const seen = new Set<number>();
    for (const cl of clusters) {
      if (!Array.isArray(cl.indices) || typeof cl.snippet !== "string") {
        throw new Error(`Malformed cluster: ${JSON.stringify(cl)}`);
      }
      for (const idx of cl.indices) {
        if (idx < 0 || idx >= changes.length || seen.has(idx)) {
          throw new Error(`Invalid index ${idx}`);
        }
        seen.add(idx);
      }
    }
    if (seen.size !== changes.length) {
      throw new Error(`Clusters covered ${seen.size} of ${changes.length} changes`);
    }

    console.log(`[clusterSemantically] § ${sectionNumber} — ${changes.length} changes → ${clusters.length} clusters`);

  } catch (err) {
    console.warn("[clusterSemantically] Falling back to one-per-cluster:", err);
    clusters = changes.map((c, i) => ({
      indices: [i],
      snippet: extractSnippetClientSide(c.paragraphContext, [c]),
    }));
  }

  return clusters.map((cl, i) => {
    const clusterChanges = cl.indices.map((idx) => changes[idx]);
    const fullParagraph = clusterChanges[0].paragraphContext;

    // Verify snippet using normalized comparison so curly-quote/dash differences
    // between the AI response and paragraphContext don\'t cause false fallbacks.
    const snippetIsValid =
      cl.snippet.length > 0 &&
      normalizeForMatch(fullParagraph).includes(normalizeForMatch(cl.snippet));

    const snippet = snippetIsValid
      ? cl.snippet
      : extractSnippetClientSide(fullParagraph, clusterChanges);

    return {
      clusterId: `${sectionNumber}-${i}`,
      sectionNumber,
      sectionTitle,
      paragraphText: snippet,
      changes: clusterChanges,
    };
  });
}

// Shared normalization for snippet validity checks and sentence-hit detection.
// Mirrors the normalization in buildInlineDiff so quote/dash/whitespace
// differences don\'t cause false mismatches.
function normalizeForMatch(s: string): string {
  return s
    .replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'")
    .replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"')
    .replace(/[\u2013\u2014]/g, "-")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .toLowerCase()
    .trim();
}

// ─── Per-change centered window extraction ────────────────────────────────────
//
// paragraphContext stored on each TrackedChangeInfo is a ~800-char window
// centered on the change text within the full paragraph.  This ensures:
//   • Changes deep in long paragraphs (e.g. PAYMENT §6) are not cut off.
//   • Each change in the same paragraph gets its OWN context snippet, not
//     the same leading-edge slice shared by all changes.
//   • buildInlineDiff can always find the change text within the window.
//
// For deletion changes the deleted text is not in paragraph.text, so it is
// appended before the window search so the anchor search can still locate it
// (the appended text will be the highlight target for <del> rendering).

export function extractCenteredWindow(
  fullParagraph: string,
  changeText: string,
  windowRadius = 400
): string {
  if (!fullParagraph) return "";

  const normPara   = normalizeForMatch(fullParagraph);
  const normNeedle = normalizeForMatch(changeText.trim());

  // ── Find the anchor position of the change text ───────────────────────────
  let anchorPos = -1;

  if (normNeedle.length >= 4) {
    anchorPos = normPara.indexOf(normNeedle);

    // Fuzzy: try progressively shorter leading prefixes (handles encoding
    // differences that survive normalization, e.g. soft-hyphens inside tokens)
    if (anchorPos === -1 && normNeedle.length > 20) {
      for (const prefixLen of [60, 40, 20, 10]) {
        if (normNeedle.length <= prefixLen) continue;
        const anchor = normPara.indexOf(normNeedle.slice(0, prefixLen));
        if (anchor !== -1) { anchorPos = anchor; break; }
      }
    }
  }

  // No anchor found (deleted text absent from paragraph.text):
  // return the full paragraph so the caller can decide what to show.
  if (anchorPos === -1) return fullParagraph;

  // ── Map normalized anchor back to original string position ───────────────
  const origAnchor = normToOriginalPos(fullParagraph, anchorPos);

  // ── Expand outward to sentence boundaries within the radius ──────────────
  const rawStart = Math.max(0, origAnchor - windowRadius);
  const rawEnd   = Math.min(fullParagraph.length, origAnchor + normNeedle.length + windowRadius);

  // Trim to a sentence boundary so we don't start/end mid-word
  const beforeCut = fullParagraph.lastIndexOf(". ", origAnchor);
  const afterCut  = fullParagraph.indexOf(". ", origAnchor + normNeedle.length);

  const start = beforeCut !== -1 && beforeCut >= rawStart ? beforeCut + 2 : rawStart;
  const end   = afterCut  !== -1 && afterCut  <= rawEnd   ? afterCut  + 1 : rawEnd;

  return fullParagraph.slice(start, end).trim();
}

// Client-side snippet extractor — fallback when /api/cluster is unavailable
// or returns an invalid snippet.
//
// Uses multiple needle lengths so changes that span sentence boundaries are
// found in BOTH sentences. No hard character cap — always returns complete
// sentences so buildInlineDiff can locate the full change text.
//
// Exported so buildClusterDiffHtml in cards.ts can derive a per-paragraph
// snippet rather than re-using the cluster-level paragraphText (which is
// only valid for the first paragraph in multi-paragraph clusters).
export function extractSnippetClientSide(paragraph: string, changes: TrackedChangeInfo[]): string {
  if (!paragraph) return "";

  const sentences = paragraph.split(/(?<=[.!?])\s+/);
  if (sentences.length <= 2) return paragraph;

  const normSentences = sentences.map(normalizeForMatch);
  const hitIndices = new Set<number>();

  for (const c of changes) {
    const fullText = (c.text ?? "").trim();
    if (!fullText) continue;
    const normFull = normalizeForMatch(fullText);

    // Try progressively shorter needles to catch changes spanning sentence boundaries
    const needles = [
      normFull,
      normFull.slice(0, 80),
      normFull.slice(0, 40),
      normFull.slice(-40),
      normFull.split(" ").slice(0, 6).join(" "),
      normFull.split(" ").slice(-6).join(" "),
    ].filter(n => n.length >= 8);

    for (const needle of needles) {
      for (let si = 0; si < normSentences.length; si++) {
        if (normSentences[si].includes(needle)) {
          hitIndices.add(si);
        }
      }
    }
  }

  if (hitIndices.size === 0) return paragraph;

  const minI = Math.max(0, Math.min(...hitIndices) - 1);
  const maxI = Math.min(sentences.length - 1, Math.max(...hitIndices) + 1);
  return sentences.slice(minI, maxI + 1).join(" ");
}

// ─── Legacy fallback clustering ───────────────────────────────────────────────

export function clusterSection(
  sectionNumber: string,
  sectionTitle: string,
  changes: TrackedChangeInfo[]
): RedlineCluster[] {
  const byPara = new Map<string, TrackedChangeInfo[]>();
  const paraOrder: string[] = [];

  for (const change of changes) {
    const key = (change.paragraphContext ?? "").slice(0, 120).trim();
    if (!byPara.has(key)) {
      byPara.set(key, []);
      paraOrder.push(key);
    }
    byPara.get(key)!.push(change);
  }

  return paraOrder.map((key, i) => {
    const clusterChanges = byPara.get(key)!;
    return {
      clusterId: `${sectionNumber}-${i}`,
      sectionNumber,
      sectionTitle,
      paragraphText: extractSnippetClientSide(clusterChanges[0].paragraphContext, clusterChanges),
      changes: clusterChanges,
    };
  });
}

// ─── Inline diff rendering ────────────────────────────────────────────────────
//
// paragraphContext is captured from Word BEFORE changes are accepted, so both
// inserted and deleted text are present in the paragraph string.
//
// We search the snippet (paragraphText) for ALL change text — both insertions
// and deletions — and mark them inline:
//   Insertions → <ins class="redline-ins">
//   Deletions  → <del class="redline-del">
//
// If a change text is not found in the snippet (it may fall outside the
// visible sentence window) it is silently skipped — no orphan blocks.

export function buildInlineDiff(paragraphText: string, changes: TrackedChangeInfo[]): string {
  if (!paragraphText) return "";

  interface Span {
    start: number;
    end: number;
    change: TrackedChangeInfo;
  }

  // Normalize a string for comparison:
  // Word uses curly quotes/apostrophes and various whitespace that may differ
  // between paragraphContext (captured from the DOM) and change.text (from the
  // tracked-change object). Flattening these lets substring search succeed even
  // when encoding differs.
  function normalize(s: string): string {
    return s
      .replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'")   // curly single quotes → '
      .replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"')   // curly double quotes → "
      .replace(/[\u2013\u2014]/g, "-")                            // en/em dash → -
      .replace(/\u00A0/g, " ")                                    // non-breaking space → space
      .replace(/\s+/g, " ")                                       // collapse whitespace
      .toLowerCase()
      .trim();
  }

  const normPara = normalize(paragraphText);
  const spans: Span[] = [];

  for (const c of changes) {
    const rawNeedle = (c.text ?? "").trim();
    if (!rawNeedle) continue;

    const normNeedle = normalize(rawNeedle);
    if (!normNeedle) continue;

    // ── Strategy 1: exact match on normalized text ─────────────────────────
    let pos = normPara.indexOf(normNeedle);

    // ── Strategy 2: fuzzy anchor — match first 40 chars, extend by length ──
    // Handles cases where the middle of a long change has encoding differences
    // that survived normalization (e.g. mixed ligatures, soft hyphens).
    if (pos === -1 && normNeedle.length > 40) {
      const anchor = normNeedle.slice(0, 40);
      const anchorPos = normPara.indexOf(anchor);
      if (anchorPos !== -1) {
        // Verify the region roughly matches by checking the tail too
        const tail = normNeedle.slice(-20);
        const expectedEnd = anchorPos + normNeedle.length;
        const searchWindow = normPara.slice(anchorPos, expectedEnd + 20);
        if (searchWindow.includes(tail)) {
          pos = anchorPos;
          console.warn(`[buildInlineDiff] Used fuzzy anchor for: "${rawNeedle.slice(0, 60)}"`);
        }
      }
    }

    // ── Strategy 3: first-words match — for very long deletions/insertions ──
    // Match on the first 6 words, use that as a positional anchor.
    if (pos === -1) {
      const firstWords = normNeedle.split(" ").slice(0, 6).join(" ");
      if (firstWords.length >= 10) {
        const anchorPos = normPara.indexOf(firstWords);
        if (anchorPos !== -1) {
          pos = anchorPos;
          console.warn(`[buildInlineDiff] Used first-words anchor for: "${rawNeedle.slice(0, 60)}"`);
        }
      }
    }

    if (pos === -1) {
      console.warn(`[buildInlineDiff] Change text not found (all strategies failed): "${rawNeedle.slice(0, 60)}"`);
      continue;
    }

    // pos is an index into normPara. Because normalize() may have changed
    // character counts (e.g. collapsed multiple spaces into one), we need to
    // find the corresponding position in the ORIGINAL paragraphText.
    // We do this by walking both strings in parallel until we've consumed
    // `pos` normalized characters.
    const originalStart = normToOriginalPos(paragraphText, pos);
    const originalEnd   = normToOriginalPos(paragraphText, pos + normNeedle.length);

    spans.push({ start: originalStart, end: originalEnd, change: c });
  }

  // Sort by position, drop overlapping spans (keep first)
  spans.sort((a, b) => a.start - b.start);
  const merged: Span[] = [];
  for (const span of spans) {
    const prev = merged[merged.length - 1];
    if (prev && span.start < prev.end) {
      console.warn(`[buildInlineDiff] Overlapping span at ${span.start}, skipping`);
      continue;
    }
    merged.push(span);
  }

  let result = "";
  let cursor = 0;

  for (const span of merged) {
    if (span.start > cursor) {
      result += escapeHtml(paragraphText.slice(cursor, span.start));
    }

    const changedText = paragraphText.slice(span.start, span.end);
    const type = (span.change.type ?? "").toLowerCase();
    const author = escapeHtml(span.change.author ?? "");

    if (type.includes("delet")) {
      result += `<del class="redline-del" title="Deleted by ${author}">${escapeHtml(changedText)}</del>`;
    } else {
      result += `<ins class="redline-ins" title="Inserted by ${author}">${escapeHtml(changedText)}</ins>`;
    }

    cursor = span.end;
  }

  if (cursor < paragraphText.length) {
    result += escapeHtml(paragraphText.slice(cursor));
  }

  return result;
}

// Maps a character position in the normalized string back to the corresponding
// position in the original string. normalize() may collapse whitespace runs
// (e.g. "  " → " ") so the two strings can have different lengths. We walk
// both in parallel, advancing the original pointer past any characters that
// normalize() would have removed or merged.
function normToOriginalPos(original: string, normPos: number): number {
  let origIdx = 0;
  let normIdx = 0;

  // Replicate the normalization transformations character by character
  while (origIdx < original.length && normIdx < normPos) {
    const ch = original[origIdx];

    // Skip characters that normalize() removes entirely (none currently, but
    // soft-hyphen \u00AD and zero-width spaces are common in Word docs)
    if (ch === "\u00AD" || ch === "\u200B" || ch === "\uFEFF") {
      origIdx++;
      continue;
    }

    // Whitespace: normalize() collapses runs to a single space.
    // Count one normalized char for the whole run.
    if (/\s/.test(ch)) {
      // advance past the entire whitespace run in original
      while (origIdx < original.length && /\s/.test(original[origIdx])) {
        origIdx++;
      }
      normIdx++; // one space in normalized string
      continue;
    }

    origIdx++;
    normIdx++;
  }

  // If normPos overshoots (e.g. end of string), return end of original
  return Math.min(origIdx, original.length);
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ─── Write ────────────────────────────────────────────────────────────────────

export async function addWordCommentOnRedline(args: {
  changeText: string;
  paragraphContext: string;
  sectionTitle: string;
  sectionNumber: string;
  commentText: string;
}): Promise<CommentResult> {
  const commentText   = (args.commentText ?? "").trim();
  const changeText    = (args.changeText ?? "").trim();
  const paraContext   = (args.paragraphContext ?? "").trim();
  const sectionNumber = (args.sectionNumber ?? "").trim();

  if (!commentText) return { ok: false, error: "Missing commentText" };

  return Word.run(async (context) => {
    const body = context.document.body;

    if (changeText.length >= 6) {
      const r1 = body.search(changeText.slice(0, 80), { matchCase: false });
      r1.load("items");
      await context.sync();
      if (r1.items.length) {
        (r1.items[0] as any).insertComment(commentText);
        await context.sync();
        return { ok: true, matches: r1.items.length, usedIndex: 0 };
      }
    }

    if (paraContext.length >= 20) {
      const words = paraContext.replace(/\s+/g, " ").split(" ");
      const snippet = words.slice(2, 10).join(" ").slice(0, 80);
      if (snippet.length >= 10) {
        const r2 = body.search(snippet, { matchCase: false });
        r2.load("items");
        await context.sync();
        if (r2.items.length) {
          (r2.items[0] as any).insertComment(commentText);
          await context.sync();
          return { ok: true, matches: r2.items.length, usedIndex: 0 };
        }
      }
    }

    if (sectionNumber) {
      const r3 = body.search(sectionNumber, { matchCase: false });
      r3.load("items");
      await context.sync();
      if (r3.items.length) {
        (r3.items[0] as any).insertComment(commentText);
        await context.sync();
        return { ok: true, matches: r3.items.length, usedIndex: 0 };
      }
    }

    return { ok: false, error: "Could not find any anchor text in document", matches: 0 };
  });
}

// ─── Cluster-spanning comment insertion ───────────────────────────────────────
//
// Inserts a single Word comment that spans the full range of ALL tracked
// changes in the cluster (from the start of the first change to the end of
// the last), so the comment appears anchored to the entire redlined region
// rather than just the first change's text.
//
// Strategy:
//   1. Load all tracked changes from the document.
//   2. Match them to the cluster's change IDs.
//   3. Get each matching change's Range and union them via expandTo.
//   4. Insert the comment on the unioned range.
//   5. Fall back to addWordCommentOnRedline if getRange() is unavailable or
//      no IDs match (older Word versions).

export async function addCommentOnCluster(args: {
  changeIds: string[];
  commentText: string;
  // fallback fields used if the Revision.range approach fails:
  firstChangeText: string;
  paragraphContext: string;
  sectionTitle: string;
  sectionNumber: string;
}): Promise<CommentResult> {
  const commentText = (args.commentText ?? "").trim();
  if (!commentText) return { ok: false, error: "Missing commentText" };

  if (!args.changeIds.length) {
    return addWordCommentOnRedline({
      changeText: args.firstChangeText,
      paragraphContext: args.paragraphContext,
      sectionTitle: args.sectionTitle,
      sectionNumber: args.sectionNumber,
      commentText,
    });
  }

  return Word.run(async (context) => {
    const body = context.document.body;

    // ── Step 1: load scalars + navigate to range in one load call ────────────
    // Use the nested path "items/range/text" to tell Office JS to also hydrate
    // the range navigation property on each revision in the same sync.
    // Some revision types (formatting, paragraph property) have no text range —
    // we guard against that with null checks before using .range.
    const revisions = body.getTrackedChanges();
    revisions.load("items/id,items/index,items/range/text");
    await context.sync();

    // ── Step 2: find matching revisions that also have a valid range ──────────
    const idSet = new Set(args.changeIds);
    const matching = revisions.items.filter((rev) => {
      const raw = rev as unknown as Record<string, unknown>;
      const id  = typeof raw["id"] === "string" ? raw["id"] : String(raw["id"] ?? "");
      // Only keep revisions that matched an ID AND have a usable range object.
      // Formatting/property revisions may have range === undefined or null.
      return idSet.has(id) && rev.range != null;
    });

    if (matching.length === 0) {
      console.warn("[addCommentOnCluster] No revisions with ranges matched — falling back to text search");
      return addWordCommentOnRedline({
        changeText: args.firstChangeText,
        paragraphContext: args.paragraphContext,
        sectionTitle: args.sectionTitle,
        sectionNumber: args.sectionNumber,
        commentText,
      });
    }

    // ── Step 3: sort by document index, union ranges first → last ────────────
    const sorted = [...matching].sort((a, b) => (a.index ?? 0) - (b.index ?? 0));

    let unionRange: Word.Range = sorted[0].range;
    for (let i = 1; i < sorted.length; i++) {
      // revision.range is only the changed text — not the whole paragraph —
      // so expandTo spans exactly the cluster edits, nothing more.
      unionRange = unionRange.expandTo(sorted[i].range);
    }

    // ── Step 4: insert the comment on the exact revision-spanning range ───────
    unionRange.insertComment(commentText);
    await context.sync();

    console.log(
      `[addCommentOnCluster] Comment inserted on ${sorted.length} revision(s), ` +
      `index ${sorted[0].index} → ${sorted[sorted.length - 1].index}`
    );
    return { ok: true, matches: sorted.length, usedIndex: 0 };
  });
}

export async function addWordCommentByAnchor(args: {
  anchorText: string;
  commentText: string;
  occurrence?: number;
  matchCase?: boolean;
  matchWholeWord?: boolean;
}): Promise<CommentResult> {
  const anchorText  = (args.anchorText ?? "").trim();
  const commentText = (args.commentText ?? "").trim();

  if (!anchorText || !commentText) {
    return { ok: false, error: "Missing anchorText or commentText" };
  }

  const occurrence     = Number.isInteger(args.occurrence) ? (args.occurrence as number) : 0;
  const matchCase      = !!args.matchCase;
  const matchWholeWord = !!args.matchWholeWord;

  return Word.run(async (context) => {
    const results = context.document.body.search(anchorText, { matchCase, matchWholeWord });
    results.load("items");
    await context.sync();

    if (!results.items?.length) {
      return { ok: false, error: "Anchor text not found", matches: 0 };
    }

    const idx = Math.min(Math.max(occurrence, 0), results.items.length - 1);
    (results.items[idx] as any).insertComment(commentText);
    await context.sync();
    return { ok: true, matches: results.items.length, usedIndex: idx };
  });
}

// ─── Private helpers ──────────────────────────────────────────────────────────

interface SectionEntry {
  sectionNumber: string;
  sectionTitle: string;
  paragraphTexts: string[];
  startParaIndex: number;
}

function buildSectionMap(paragraphTexts: string[]): SectionEntry[] {
  const sections: SectionEntry[] = [];
  let current: SectionEntry = {
    sectionNumber: "",
    sectionTitle: "PREAMBLE",
    paragraphTexts: [],
    startParaIndex: 0,
  };

  for (let i = 0; i < paragraphTexts.length; i++) {
    const trimmed = paragraphTexts[i].trim();
    if (!trimmed) continue;

    const headingMatch = trimmed.match(SECTION_HEADING_RE);
    if (headingMatch) {
      sections.push(current);
      const title = headingMatch[1].trim();
      current = {
        sectionNumber: title,
        sectionTitle: title,
        paragraphTexts: [trimmed],
        startParaIndex: i,
      };
      continue;
    }
    current.paragraphTexts.push(trimmed);
  }
  sections.push(current);
  return sections;
}

function findSectionForText(
  changeText: string,
  sections: SectionEntry[]
): { sectionNumber: string; sectionTitle: string; paragraphContext: string } | null {
  const needle = changeText.trim().slice(0, 120).toLowerCase();
  if (needle.length < 20) return null;

  for (const section of sections) {
    if (!section.sectionNumber) continue;
    for (const para of section.paragraphTexts) {
      if (para.toLowerCase().includes(needle)) {
        return {
          sectionNumber: section.sectionNumber,
          sectionTitle: section.sectionTitle,
          paragraphContext: para, // full paragraph — caller will extract a centered window
        };
      }
    }
  }
  return null;
}

function findSectionByPosition(
  changeText: string,
  paraTexts: string[],
  paraIndexToSection: SectionEntry[]
): { sectionNumber: string; sectionTitle: string; paragraphContext: string } | null {
  const words = changeText.trim().split(/\s+/).slice(0, 4).join(" ").toLowerCase();
  if (words.length < 4) return null;

  for (let i = 0; i < paraTexts.length; i++) {
    if (paraTexts[i].toLowerCase().includes(words)) {
      const section = paraIndexToSection[i];
      if (section?.sectionNumber) {
        return {
          sectionNumber: section.sectionNumber,
          sectionTitle: section.sectionTitle,
          paragraphContext: paraTexts[i], // full paragraph — caller will extract a centered window
        };
      }
    }
  }
  return null;
}