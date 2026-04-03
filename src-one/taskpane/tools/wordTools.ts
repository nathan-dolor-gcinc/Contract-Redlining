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

// Matches section headings that may be prefixed with a number like "5.0 PAYMENT."
// The numeric prefix is optional — headings without numbers (e.g. "ATTACHMENT A.")
// are also matched. Capture group 1 is always just the title text, e.g. "PAYMENT".
const SECTION_HEADING_RE = /^(?:\d+(?:\.\d+)*\.?\s+)?([A-Z][A-Z\s&,\/()'".-]{2,}[A-Z])\./;

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

    // ── DIAGNOSTIC: getReviewedText on body ───────────────────────────────────
    // Queue both reviewed text calls alongside the other loads so they resolve
    // in the same sync round trip.
    const reviewedCurrent  = (body as any).getReviewedText("current")  as OfficeExtension.ClientResult<string>;
    const reviewedOriginal = (body as any).getReviewedText("original") as OfficeExtension.ClientResult<string>;

    await context.sync();

    try {
      const currText = reviewedCurrent.value ?? "";
      const origText = reviewedOriginal.value ?? "";
      console.group("[DIAGNOSTIC] body.getReviewedText");
      console.log("--- current (accept all) ---");
      console.log(currText.slice(0, 3000));
      console.log("--- original (reject all) ---");
      console.log(origText.slice(0, 3000));
      console.log(`[diff] current length=${currText.length} | original length=${origText.length} | same=${currText === origText}`);
      console.groupEnd();
    } catch (err) {
      console.warn("[DIAGNOSTIC] getReviewedText not available on this Word version:", err);
    }

    const activeItems = changes.items;

    console.group("[getTrackedChanges] Raw changes from Word:", activeItems.length);
    for (const c of activeItems) {
      console.log(`  [${c.type}] text: "${(c.text ?? "").slice(0, 60)}" | author: "${c.author}"`);
    }
    console.groupEnd();

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

    console.group("[getTrackedChanges] Raw changes from Word:", activeItems.length);
    for (const c of activeItems) {
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

    const rawEntries: RawChangeEntry[] = activeItems.map((c) => {
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

    return rawEntries.flatMap((entry) => {
      const { id, date, changeText, isDeletion } = entry;
      const section = entry.section;

      if (!section || !section.sectionNumber) {
        console.warn(`[getTrackedChanges] DROPPING "${changeText.slice(0, 60)}" — no section found`);
        return [];
      }

      let fullParagraph = section.paragraphContext;

      if (isDeletion && changeText.trim().length > 0) {
        const normCtx  = normalizeForMatch(fullParagraph);
        const normHead = normalizeForMatch(changeText.slice(0, 60));
        if (normHead.length >= 4 && !normCtx.includes(normHead)) {
          fullParagraph = fullParagraph ? fullParagraph + " " + changeText : changeText;
        }
      }

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
        rawParagraphText: section.paragraphContext,
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

export function extractCenteredWindow(
  fullParagraph: string,
  changeText: string,
  windowRadius = 400
): string {
  if (!fullParagraph) return "";

  const normPara   = normalizeForMatch(fullParagraph);
  const normNeedle = normalizeForMatch(changeText.trim());

  let anchorPos = -1;

  if (normNeedle.length >= 4) {
    anchorPos = normPara.indexOf(normNeedle);

    if (anchorPos === -1 && normNeedle.length > 20) {
      for (const prefixLen of [60, 40, 20, 10]) {
        if (normNeedle.length <= prefixLen) continue;
        const anchor = normPara.indexOf(normNeedle.slice(0, prefixLen));
        if (anchor !== -1) { anchorPos = anchor; break; }
      }
    }
  }

  if (anchorPos === -1) return fullParagraph;

  const origAnchor = normToOriginalPos(fullParagraph, anchorPos);

  const rawStart = Math.max(0, origAnchor - windowRadius);
  const rawEnd   = Math.min(fullParagraph.length, origAnchor + normNeedle.length + windowRadius);

  const beforeCut = fullParagraph.lastIndexOf(". ", origAnchor);
  const afterCut  = fullParagraph.indexOf(". ", origAnchor + normNeedle.length);

  const start = beforeCut !== -1 && beforeCut >= rawStart ? beforeCut + 2 : rawStart;
  const end   = afterCut  !== -1 && afterCut  <= rawEnd   ? afterCut  + 1 : rawEnd;

  return fullParagraph.slice(start, end).trim();
}

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

export function buildInlineDiff(paragraphText: string, changes: TrackedChangeInfo[]): string {
  if (!paragraphText) return "";

  interface Span {
    start: number;
    end: number;
    change: TrackedChangeInfo;
  }

  function normalize(s: string): string {
    return s
      .replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'")
      .replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"')
      .replace(/[\u2013\u2014]/g, "-")
      .replace(/\u00A0/g, " ")
      .replace(/\s+/g, " ")
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

    let pos = normPara.indexOf(normNeedle);

    if (pos === -1 && normNeedle.length > 40) {
      const anchor = normNeedle.slice(0, 40);
      const anchorPos = normPara.indexOf(anchor);
      if (anchorPos !== -1) {
        const tail = normNeedle.slice(-20);
        const expectedEnd = anchorPos + normNeedle.length;
        const searchWindow = normPara.slice(anchorPos, expectedEnd + 20);
        if (searchWindow.includes(tail)) {
          pos = anchorPos;
          console.warn(`[buildInlineDiff] Used fuzzy anchor for: "${rawNeedle.slice(0, 60)}"`);
        }
      }
    }

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

    const originalStart = normToOriginalPos(paragraphText, pos);
    const originalEnd   = normToOriginalPos(paragraphText, pos + normNeedle.length);

    spans.push({ start: originalStart, end: originalEnd, change: c });
  }

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

function normToOriginalPos(original: string, normPos: number): number {
  let origIdx = 0;
  let normIdx = 0;

  while (origIdx < original.length && normIdx < normPos) {
    const ch = original[origIdx];

    if (ch === "\u00AD" || ch === "\u200B" || ch === "\uFEFF") {
      origIdx++;
      continue;
    }

    if (/\s/.test(ch)) {
      while (origIdx < original.length && /\s/.test(original[origIdx])) {
        origIdx++;
      }
      normIdx++;
      continue;
    }

    origIdx++;
    normIdx++;
  }

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

export async function addCommentOnCluster(args: {
  changeIds: string[];
  commentText: string;
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

    const revisions = body.getTrackedChanges();
    revisions.load("items/id,items/index,items/range/text");
    await context.sync();

    const idSet = new Set(args.changeIds);
    const matching = revisions.items.filter((rev) => {
      const raw = rev as unknown as Record<string, unknown>;
      const id  = typeof raw["id"] === "string" ? raw["id"] : String(raw["id"] ?? "");
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

    const sorted = [...matching].sort((a, b) => (a.index ?? 0) - (b.index ?? 0));

    let unionRange: Word.Range = sorted[0].range;
    for (let i = 1; i < sorted.length; i++) {
      unionRange = unionRange.expandTo(sorted[i].range);
    }

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

      // Extract the numeric prefix separately (e.g. "5.0") from the title (e.g. "PAYMENT")
      // so sectionNumber can be used for ordering and sectionTitle for config lookup.
      const numericMatch = trimmed.match(/^(\d+(?:\.\d+)*)/);
      const numericPrefix = numericMatch?.[1] ?? "";
      const title = headingMatch[1].trim();

      current = {
        sectionNumber: numericPrefix || title, // "5.0" if present, else "PAYMENT"
        sectionTitle: title,                   // always just "PAYMENT"
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
          paragraphContext: para,
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
          paragraphContext: paraTexts[i],
        };
      }
    }
  }
  return null;
}