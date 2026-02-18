// src/taskpane/tools/dispatchTool.ts
import { readWordBodyText, addWordCommentByAnchor } from "./wordTools";

export async function executeToolCall(toolCall: any): Promise<{ tool_call_id: string; output: string }> {
  const toolCallId = toolCall.id;
  const name = toolCall.function?.name;
  const args = toolCall.function?.arguments ? JSON.parse(toolCall.function.arguments) : {};

  if (name === "read_word_body") {
    const text = await readWordBodyText();
    console.log("✅ read_word_body length:", text.length);
    console.log("✅ read_word_body preview:", text.slice(0, 300));

    return { tool_call_id: toolCallId, output: text };
  }

  if (name === "add_word_comment") {
    const result = await addWordCommentByAnchor(args);
    return { tool_call_id: toolCallId, output: JSON.stringify(result) };
  }

  return { tool_call_id: toolCallId, output: JSON.stringify({ ok: false, error: `Unknown tool: ${name}` }) };
}