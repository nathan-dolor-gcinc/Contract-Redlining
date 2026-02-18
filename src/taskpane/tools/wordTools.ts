// src/taskpane/tools/wordTools.ts

export async function readWordBodyText(maxChars = 60000): Promise<string> {
  const text = await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text || "";
  }); // Word body text is accessible via body.text [3](https://github.com/Azure/azure-sdk-for-python/blob/main/sdk/ai/azure-ai-projects/README.md)

  return text.length > maxChars ? text.slice(0, maxChars) + "\n...[truncated]..." : text;
}

export async function addWordCommentByAnchor(args: {
  anchorText: string;
  commentText: string;
  occurrence?: number;
  matchCase?: boolean;
  matchWholeWord?: boolean;
}): Promise<{ ok: boolean; matches?: number; usedIndex?: number; error?: string }> {
  const anchorText = (args.anchorText || "").trim();
  const commentText = (args.commentText || "").trim();

  const occurrence = Number.isInteger(args.occurrence) ? (args.occurrence as number) : 0;
  const matchCase = !!args.matchCase;
  const matchWholeWord = !!args.matchWholeWord;

  if (!anchorText || !commentText) {
    return { ok: false, error: "Missing anchorText/commentText" };
  }

  return Word.run(async (context) => {
    const results = context.document.body.search(anchorText, { matchCase, matchWholeWord });
    results.load("items");
    await context.sync();

    if (!results.items || results.items.length === 0) {
      return { ok: false, error: "Anchor text not found", matches: 0 };
    }

    const idx = Math.min(Math.max(occurrence, 0), results.items.length - 1);
    const range = results.items[idx];

    // insertComment exists on ranges/selection in Word API (WordApi 1.4+) [4](https://learn.microsoft.com/en-us/python/api/azure-ai-projects/azure.ai.projects.aiprojectclient?view=azure-python)
    (range as any).insertComment(commentText);

    await context.sync();
    return { ok: true, matches: results.items.length, usedIndex: idx };
  });
}