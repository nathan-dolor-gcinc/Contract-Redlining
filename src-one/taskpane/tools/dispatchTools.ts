// src-one/taskpane/tools/dispatchTools.ts

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

let _advanceCallback: (() => Promise<void>) | null = null;

// Called by review.ts after all functions are defined.
export function registerAdvanceCallback(cb: () => Promise<void>): void {
  _advanceCallback = cb;
}

// Called by client.ts AFTER the active run is closed, so the thread is free.
export async function triggerAdvance(): Promise<void> {
  if (_advanceCallback) {
    await _advanceCallback();
  }
}

// Called by client.ts when add_word_comment fires via chat (not via button).
// review.ts registers this to keep session state in sync.
let _commentWrittenCallback: (() => void) | null = null;

export function registerCommentWrittenCallback(cb: () => void): void {
  _commentWrittenCallback = cb;
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
        // Notify review.ts so it can increment the index and update progress.
        _commentWrittenCallback?.();
        return { tool_call_id, output: JSON.stringify(result) };
      }

      case "advance_to_next_cluster": {
        // Handled by client.ts directly — should not reach here.
        console.log("✅ advance_to_next_cluster — handled by client.ts");
        return { tool_call_id, output: JSON.stringify({ ok: true }) };
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