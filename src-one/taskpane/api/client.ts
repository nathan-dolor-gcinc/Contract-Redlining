// src/taskpane/api/client.ts

import { executeToolCall, triggerAdvance } from "../tools/dispatchTools";
import type { ToolCall, ToolResult } from "../tools/dispatchTools";

export const BACKEND_BASE_URL = "https://localhost:3001";

export interface BackendResponse {
  conversationId?: string;
  toolRequests?: ToolCall[];
  reply?: string;
}

// ─── Chat ─────────────────────────────────────────────────────────────────────

export async function sendPrompt(
  prompt: string,
  conversationId: string | null
): Promise<{ reply: string | null; conversationId: string | null }> {
  let currentConversationId = conversationId;

  const resp = await fetch(`${BACKEND_BASE_URL}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, conversationId: currentConversationId }),
  });

  let data = (await resp.json()) as BackendResponse;

  if (!resp.ok) {
    console.error("[api/client] /api/chat error:", data);
    return { reply: null, conversationId: currentConversationId };
  }

  currentConversationId = data.conversationId ?? currentConversationId;

  // ── Tool-call loop ──────────────────────────────────────────────────────────
  while (data.toolRequests?.length) {
    console.log("🔧 Tool requests:", data.toolRequests);

    const toolResults: ToolResult[] = [];
    let advanceToolCallId: string | null = null;

    for (const toolCall of data.toolRequests) {
      console.log("▶ Executing tool:", toolCall.function?.name);

      if (toolCall.function?.name === "advance_to_next_cluster") {
        // Don't execute the callback yet — just record the tool call id.
        // We must submit the tool result first to close the active run,
        // then trigger the UI advance after.
        advanceToolCallId = toolCall.id;
        toolResults.push({
          tool_call_id: toolCall.id,
          output: JSON.stringify({ ok: true }),
        });
      } else {
        toolResults.push(await executeToolCall(toolCall));
      }
    }

    // Submit tool outputs — this closes the active run on the backend
    const resp2 = await fetch(`${BACKEND_BASE_URL}/api/tool-result`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ conversationId: currentConversationId, toolResults }),
    });

    if (!resp2.ok) {
      const errData = await resp2.json().catch(() => ({}));
      console.error("[api/client] /api/tool-result error:", errData);
      return { reply: null, conversationId: currentConversationId };
    }

    data = (await resp2.json()) as BackendResponse;
    console.log("⬅ Backend after tool-result:", data);
    currentConversationId = data.conversationId ?? currentConversationId;

    // Now that the run is closed, trigger the UI advance.
    // This starts a fresh run for the next cluster — no conflict.
    if (advanceToolCallId !== null) {
      await triggerAdvance();
      // Return null reply — the next cluster card is already rendered.
      // Returning the stale reply here would append a duplicate message.
      return { reply: null, conversationId: currentConversationId };
    }
  }

  return { reply: data.reply ?? null, conversationId: currentConversationId };
}

// ─── Session ──────────────────────────────────────────────────────────────────

export interface EndSessionResult {
  ok: boolean;
  message?: string;
  error?: string;
}

export async function endSession(conversationId: string | null): Promise<EndSessionResult> {
  const resp = await fetch(`${BACKEND_BASE_URL}/api/session`, {
    method: "DELETE",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ conversationId }),
  });
  const data = await resp.json();
  return resp.ok ? { ok: true, message: data.message } : { ok: false, error: data.error ?? "unknown error" };
}

// ─── Prime ────────────────────────────────────────────────────────────────────

/*
export async function primeThread(content: string): Promise<string | null> {
  try {
    const resp = await fetch(`${BACKEND_BASE_URL}/api/prime`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ content }),
    });
    const data = await resp.json();
    if (!resp.ok) {
      console.error("[api/client] /api/prime error:", data);
      return null;
    }
    return data.conversationId ?? null;
  } catch (err) {
    console.error("[api/client] /api/prime threw:", err);
    return null;
  }
}
  */