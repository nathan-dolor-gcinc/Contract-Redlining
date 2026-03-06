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

    return changes.items.flatMap((c) => {
      const raw = c as unknown as Record<string, unknown>;
      const id = typeof raw["id"] === "string" ? raw["id"] : String(raw["id"] ?? "");

      const rawDate = c.date as string | Date | undefined;
      const date =
        rawDate instanceof Date ? rawDate.toISOString() :
        typeof rawDate === "string" ? rawDate : "";

      const changeText = c.text ?? "";

      console.group(`[attribution] "${changeText.slice(0, 60)}" (${c.type})`);

      let section = findSectionByPosition(changeText, paraTexts, paraIndexToSection);
      if (!section) {
        const found = findSectionForText(changeText, sectionEntries);
        if (found?.sectionNumber) section = found;
      }

      if (!section || !section.sectionNumber) {
        console.warn("  ⚠ DROPPING — could not attribute to any real section");
        console.groupEnd();
        return [];
      }

      console.log(`  ✅ Final: § ${section.sectionNumber} | context: "${section.paragraphContext?.slice(0, 80)}"`);
      console.groupEnd();

      return [{
        id,
        type: String(c.type ?? ""),
        author: c.author ?? "",
        date,
        text: changeText,
        sectionTitle: section.sectionTitle,
        sectionNumber: section.sectionNumber,
        paragraphContext: section.paragraphContext,
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

// Client-side snippet extractor — fallback when /api/cluster is unavailable
// or returns an invalid snippet.
//
// Uses multiple needle lengths so changes that span sentence boundaries are
// found in BOTH sentences. No hard character cap — always returns complete
// sentences so buildInlineDiff can locate the full change text.
function extractSnippetClientSide(paragraph: string, changes: TrackedChangeInfo[]): string {
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
          paragraphContext: para.slice(0, 2000),
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
          paragraphContext: paraTexts[i].slice(0, 2000),
        };
      }
    }
  }
  return null;
}