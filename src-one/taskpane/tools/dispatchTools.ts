// src-one/taskpane/tools/dispatchTools.ts
//
// Executes tool calls received from the backend.
// Only three tools are registered — there are no resolve/accept/reject tools
// because the add-in never modifies tracked changes directly.

import { readWordBodyText, addWordCommentByAnchor, getTrackedChanges } from "./wordTools";

export interface ToolCall {
  id: string;
  function: {
    name: string;
    arguments: string; // JSON string
  };
}

export interface ToolResult {
  tool_call_id: string;
  output: string; // JSON string
}

export async function executeToolCall(toolCall: ToolCall): Promise<ToolResult> {
  const { id: tool_call_id, function: fn } = toolCall;
  const name = fn?.name ?? "";
  const args = fn?.arguments ? JSON.parse(fn.arguments) : {};

  try {
    switch (name) {

      case "read_word_body": {
        const text = await readWordBodyText(args.maxChars);
        console.log(`✅ read_word_body — ${text.length} chars`);
        return { tool_call_id, output: text };
      }

      case "get_tracked_changes": {
        const changes = await getTrackedChanges();
        console.log(`✅ get_tracked_changes — ${changes.length} changes`);
        return { tool_call_id, output: JSON.stringify(changes) };
      }

      case "add_word_comment": {
        const result = await addWordCommentByAnchor(args);
        console.log("✅ add_word_comment —", result);
        return { tool_call_id, output: JSON.stringify(result) };
      }

      default:
        console.warn(`⚠ Unknown tool requested: ${name}`);
        return {
          tool_call_id,
          output: JSON.stringify({ ok: false, error: `Unknown tool: ${name}` }),
        };
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(`❌ Tool ${name} threw:`, message);
    return {
      tool_call_id,
      output: JSON.stringify({ ok: false, error: message }),
    };
  }
}