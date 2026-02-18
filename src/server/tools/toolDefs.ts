// src/server/tools/toolDefs.ts
import type { ChatCompletionTool } from "openai/resources/chat/completions";

export const TOOL_DEFS = [
  {
    type: "function",
    function: {
      name: "read_word_body",
      description: "Read the full body text of the currently open Word document.",
      parameters: {
        type: "object",
        properties: {},
        required: [],
        additionalProperties: false
      }
      // If you want stricter schema adherence and your types allow it:
      // strict: true,
    }
  },
  {
    type: "function",
    function: {
      name: "add_word_comment",
      description:
        "Insert a Word comment anchored to a span of text in the document. Provide anchorText that appears in the document and commentText to insert.",
      parameters: {
        type: "object",
        properties: {
          anchorText: {
            type: "string",
            description:
              "Exact snippet of text (10–30 words recommended) copied from the document body to locate where the comment should be anchored."
          },
          commentText: {
            type: "string",
            description:
              "The content of the Word comment. Should include suggested change + brief rationale."
          },
          occurrence: {
            type: "integer",
            description:
              "If anchorText appears multiple times, which match to use (0 = first match).",
            default: 0,
            minimum: 0
          },
          matchCase: {
            type: "boolean",
            description: "Whether the search for anchorText is case-sensitive.",
            default: false
          },
          matchWholeWord: {
            type: "boolean",
            description:
              "Whether the search should match whole words only (prevents partial matches).",
            default: false
          }
          // Optional future-proofing:
          // applyToAllMatches: {
          //   type: "boolean",
          //   description: "If true, insert the comment at all matches of anchorText.",
          //   default: false
          // }
        },
        required: ["anchorText", "commentText"],
        additionalProperties: false
      }
      // If you want stricter schema adherence and your types allow it:
      // strict: true,
    }
  }

] satisfies ChatCompletionTool[];