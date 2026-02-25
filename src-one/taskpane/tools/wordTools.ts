// src-one/taskpane/tools/wordTools.ts
//
// All Office.js document interactions.
// NOTE: This file intentionally has NO accept/reject functions.
// Redlines are never modified by the add-in — only comments are inserted.

// ─── Types ────────────────────────────────────────────────────────────────────

export interface TrackedChangeInfo {
  id: string;
  type: string;
  author: string;
  date: string;
  text: string;
  /** The section heading this change belongs to, e.g. "2.0 SCOPE OF WORK" */
  sectionTitle: string;
  /** The section number, e.g. "2.0" */
  sectionNumber: string;
  /** Full paragraph text that contains this change (up to 400 chars) */
  paragraphContext: string;
}

/**
 * A contract section that contains one or more tracked changes.
 */
export interface RedlinedSection {
  sectionNumber: string;
  sectionTitle: string;
  changes: TrackedChangeInfo[];
  sectionContext: string;
}

export interface CommentResult {
  ok: boolean;
  matches?: number;
  usedIndex?: number;
  error?: string;
}

// ─── Section heading regex ────────────────────────────────────────────────────
// Matches: "1.0 CONTRACT", "7.1 INDEMNITY", "ATTACHMENT A.1 SCOPE OF WORK"
const SECTION_HEADING_RE = /^(\d+\.\d+|\d+\.0|ATTACHMENT\s+[A-Z]\.\d+)\s+(.+)/i;

// ─── Read ─────────────────────────────────────────────────────────────────────

/**
 * Read the full body text of the active Word document.
 */
export async function readWordBodyText(maxChars = 60_000): Promise<string> {
  const text = await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text ?? "";
  });
  return text.length > maxChars ? text.slice(0, maxChars) + "\n...[truncated]..." : text;
}

/**
 * Return all tracked changes, each annotated with the contract section it belongs to.
 *
 * FIX: Previously used text-search to map changes to sections, which failed for
 * deletions (the deleted text no longer appears in the rendered paragraph body).
 * Now loads each change's paragraph range and uses its text to find the section,
 * falling back to a position-based paragraph-index scan.
 */
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

    // Build a flat list of (paragraphIndex → sectionEntry) for O(1) lookup
    const paraIndexToSection: SectionEntry[] = [];
    let sectionIdx = 0;
    for (let i = 0; i < paraTexts.length; i++) {
      // Advance sectionIdx when we hit the next section heading paragraph
      if (
        sectionIdx + 1 < sectionEntries.length &&
        paraTexts[i].trim() === sectionEntries[sectionIdx + 1].paragraphTexts[0]?.trim()
      ) {
        sectionIdx++;
      }
      paraIndexToSection[i] = sectionEntries[sectionIdx];
    }

    // Load each change's range so we can find which paragraph it lives in
    const rangeLoads = changes.items.map((c) => {
      const range = (c as any).getRange?.() as Word.Range | undefined;
      if (range) range.load("paragraphs/items/text");
      return range;
    });
    try {
      await context.sync();
    } catch {
      // getRange may not be available in all Word versions — fall back gracefully
    }

    return changes.items.map((c, idx) => {
      const raw = c as unknown as Record<string, unknown>;
      const id = typeof raw["id"] === "string" ? raw["id"] : String(raw["id"] ?? "");

      const rawDate = c.date as string | Date | undefined;
      const date =
        rawDate instanceof Date
          ? rawDate.toISOString()
          : typeof rawDate === "string"
          ? rawDate
          : "";

      const changeText = c.text ?? "";

      // ── Strategy 1: use the change's own paragraph range (works for deletions) ──
      let section: { sectionNumber: string; sectionTitle: string; paragraphContext: string } | null = null;
      const range = rangeLoads[idx];
      if (range) {
        try {
          const firstParaText = range.paragraphs?.items?.[0]?.text ?? "";
          if (firstParaText) {
            section = findSectionForParaText(firstParaText, sectionEntries);
          }
        } catch { /* range not loaded — fall through */ }
      }

      // ── Strategy 2: text search in paragraph bodies (works for insertions) ──────
      if (!section) {
        section = findSectionForText(changeText, sectionEntries);
      }

      // ── Strategy 3: last-resort position scan using paragraph index ─────────────
      if (!section || section.sectionTitle === "PREAMBLE" && sectionEntries.length > 1) {
        const posSection = findSectionByPosition(changeText, paraTexts, paraIndexToSection);
        if (posSection) section = posSection;
      }

      return {
        id,
        type: String(c.type ?? ""),
        author: c.author ?? "",
        date,
        text: changeText,
        sectionTitle: section.sectionTitle,
        sectionNumber: section.sectionNumber,
        paragraphContext: section.paragraphContext,
      } satisfies TrackedChangeInfo;
    });
  });
}

/**
 * Group tracked changes by their contract section.
 * Returns only sections that contain redlines, in document order.
 */
export async function getRedlinedSections(): Promise<RedlinedSection[]> {
  const changes = await getTrackedChanges();

  const sectionMap = new Map<string, RedlinedSection>();
  const sectionOrder: string[] = [];

  for (const change of changes) {
    const key = change.sectionNumber || "PREAMBLE";
    if (!sectionMap.has(key)) {
      sectionMap.set(key, {
        sectionNumber: change.sectionNumber,
        sectionTitle: change.sectionTitle,
        changes: [],
        sectionContext: change.paragraphContext,
      });
      sectionOrder.push(key);
    }
    sectionMap.get(key)!.changes.push(change);
  }

  return sectionOrder.map((key) => sectionMap.get(key)!);
}

// ─── Write ────────────────────────────────────────────────────────────────────

/**
 * Insert a comment directly on the first tracked change in a section.
 *
 * FIX: Previously anchored to the section heading paragraph which placed the
 * comment at the wrong location. This version searches for the actual redlined
 * text — for insertions that's the inserted text; for deletions we search the
 * surrounding paragraph context instead, since deleted text is gone from the body.
 *
 * Falls back to the section heading only if no better anchor is found.
 */
export async function addWordCommentOnRedline(args: {
  /** The text of the tracked change (c.text). */
  changeText: string;
  /** The paragraphContext stored on the TrackedChangeInfo. */
  paragraphContext: string;
  /** Section heading fallback e.g. "2.0 SCOPE OF WORK". */
  sectionTitle: string;
  sectionNumber: string;
  commentText: string;
}): Promise<CommentResult> {
  const commentText = (args.commentText ?? "").trim();
  const changeText = (args.changeText ?? "").trim();
  const paraContext = (args.paragraphContext ?? "").trim();
  const sectionTitle = (args.sectionTitle ?? "").trim();
  const sectionNumber = (args.sectionNumber ?? "").trim();

  if (!commentText) return { ok: false, error: "Missing commentText" };

  return Word.run(async (context) => {
    const body = context.document.body;

    // ── Anchor candidate 1: the change text itself (works for insertions) ───────
    if (changeText.length >= 6) {
      const anchor1 = changeText.slice(0, 80);
      const r1 = body.search(anchor1, { matchCase: false, matchWholeWord: false });
      r1.load("items");
      await context.sync();
      if (r1.items.length) {
        (r1.items[0] as any).insertComment(commentText);
        await context.sync();
        return { ok: true, matches: r1.items.length, usedIndex: 0 };
      }
    }

    // ── Anchor candidate 2: a distinctive slice of the paragraph context ────────
    // Use words 3-12 of the paragraph to avoid matching the section heading itself
    if (paraContext.length >= 20) {
      const words = paraContext.replace(/\s+/g, " ").split(" ");
      // skip the first 2 words (often the section number/heading), take up to 8
      const snippet = words.slice(2, 10).join(" ").slice(0, 80);
      if (snippet.length >= 10) {
        const r2 = body.search(snippet, { matchCase: false, matchWholeWord: false });
        r2.load("items");
        await context.sync();
        if (r2.items.length) {
          (r2.items[0] as any).insertComment(commentText);
          await context.sync();
          return { ok: true, matches: r2.items.length, usedIndex: 0 };
        }
      }
    }

    // ── Anchor candidate 3: section heading (last resort) ───────────────────────
    const headingAnchor = sectionNumber
      ? `${sectionNumber}`.trim()
      : sectionTitle.slice(0, 40);
    const r3 = body.search(headingAnchor, { matchCase: false, matchWholeWord: false });
    r3.load("items");
    await context.sync();
    if (r3.items.length) {
      (r3.items[0] as any).insertComment(commentText);
      await context.sync();
      return { ok: true, matches: r3.items.length, usedIndex: 0 };
    }

    return { ok: false, error: "Could not find any anchor text in document", matches: 0 };
  });
}

/**
 * Insert a comment at the first (or nth) occurrence of anchorText in the document.
 * Used as a generic fallback. Prefer addWordCommentOnRedline for card actions.
 */
export async function addWordCommentByAnchor(args: {
  anchorText: string;
  commentText: string;
  occurrence?: number;
  matchCase?: boolean;
  matchWholeWord?: boolean;
}): Promise<CommentResult> {
  const anchorText = (args.anchorText ?? "").trim();
  const commentText = (args.commentText ?? "").trim();

  if (!anchorText || !commentText) {
    return { ok: false, error: "Missing anchorText or commentText" };
  }

  const occurrence = Number.isInteger(args.occurrence) ? (args.occurrence as number) : 0;
  const matchCase = !!args.matchCase;
  const matchWholeWord = !!args.matchWholeWord;

  return Word.run(async (context) => {
    const results = context.document.body.search(anchorText, { matchCase, matchWholeWord });
    results.load("items");
    await context.sync();

    if (!results.items?.length) {
      return { ok: false, error: "Anchor text not found", matches: 0 };
    }

    const idx = Math.min(Math.max(occurrence, 0), results.items.length - 1);
    const range = results.items[idx];

    (range as any).insertComment(commentText);
    await context.sync();

    return { ok: true, matches: results.items.length, usedIndex: idx };
  });
}

// ─── Private helpers ──────────────────────────────────────────────────────────

interface SectionEntry {
  sectionNumber: string;
  sectionTitle: string;
  paragraphTexts: string[];
}

function buildSectionMap(paragraphTexts: string[]): SectionEntry[] {
  const sections: SectionEntry[] = [];
  let current: SectionEntry = {
    sectionNumber: "",
    sectionTitle: "PREAMBLE",
    paragraphTexts: [],
  };

  for (const text of paragraphTexts) {
    const trimmed = text.trim();
    const match = trimmed.match(SECTION_HEADING_RE);
    if (match) {
      sections.push(current);
      current = {
        sectionNumber: match[1],
        sectionTitle: trimmed.slice(0, 80),
        paragraphTexts: [trimmed],
      };
    } else {
      current.paragraphTexts.push(trimmed);
    }
  }
  sections.push(current);
  return sections;
}

function findSectionForText(
  changeText: string,
  sections: SectionEntry[]
): { sectionNumber: string; sectionTitle: string; paragraphContext: string } {
  const needle = changeText.trim().slice(0, 60).toLowerCase();

  for (const section of sections) {
    for (const para of section.paragraphTexts) {
      if (needle && para.toLowerCase().includes(needle)) {
        return {
          sectionNumber: section.sectionNumber,
          sectionTitle: section.sectionTitle,
          paragraphContext: para.slice(0, 600),
        };
      }
    }
  }

  // Fallback: return the last NON-PREAMBLE section rather than the first
  // (deletions won't match any paragraph, so we pick the most likely section)
  const nonPreamble = sections.filter(s => s.sectionNumber !== "");
  const fallback = nonPreamble[nonPreamble.length - 1] ?? sections[sections.length - 1];
  return {
    sectionNumber: fallback?.sectionNumber ?? "",
    sectionTitle: fallback?.sectionTitle ?? "Unknown Section",
    paragraphContext: changeText.slice(0, 600),
  };
}

/** Match a change to a section using the paragraph text that contains the change. */
function findSectionForParaText(
  paraText: string,
  sections: SectionEntry[]
): { sectionNumber: string; sectionTitle: string; paragraphContext: string } | null {
  const needle = paraText.trim().slice(0, 80).toLowerCase();
  for (const section of sections) {
    for (const para of section.paragraphTexts) {
      if (needle && para.toLowerCase().includes(needle)) {
        return {
          sectionNumber: section.sectionNumber,
          sectionTitle: section.sectionTitle,
          paragraphContext: para.slice(0, 600),
        };
      }
    }
  }
  return null;
}

/**
 * Fallback: scan paragraph texts by position to find which section a change
 * belongs to, using a partial word match against surrounding text.
 */
function findSectionByPosition(
  changeText: string,
  paraTexts: string[],
  paraIndexToSection: SectionEntry[]
): { sectionNumber: string; sectionTitle: string; paragraphContext: string } | null {
  // Try matching the first 4 words of changeText anywhere in paragraphs
  const words = changeText.trim().split(/\s+/).slice(0, 4).join(" ").toLowerCase();
  if (words.length < 4) return null;

  for (let i = 0; i < paraTexts.length; i++) {
    if (paraTexts[i].toLowerCase().includes(words)) {
      const section = paraIndexToSection[i];
      if (section && section.sectionNumber) {
        return {
          sectionNumber: section.sectionNumber,
          sectionTitle: section.sectionTitle,
          paragraphContext: paraTexts[i].slice(0, 600),
        };
      }
    }
  }
  return null;
}