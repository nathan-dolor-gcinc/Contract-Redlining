// src/taskpane/api/client.ts

//

// All communication with the LexAI backend lives here.

// Nothing else should call fetch() directly.

import { executeToolCall } from "../tools/dispatchTools";

import type { ToolCall, ToolResult } from "../tools/dispatchTools";

export const BACKEND_BASE_URL = "https://localhost:3001";

export interface BackendResponse {
  conversationId?: string;

  toolRequests?: ToolCall[];

  reply?: string;
}

// ─── Chat ─────────────────────────────────────────────────────────────────────

/**

* Send a user prompt to the backend, automatically resolving any tool-call

* round-trips, and return the final assistant reply string.

*

* Returns null if the request fails or if the session has ended.

*/

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

    for (const toolCall of data.toolRequests) {
      console.log("▶ Executing tool:", toolCall.function?.name);

      toolResults.push(await executeToolCall(toolCall));
    }

    const resp2 = await fetch(`${BACKEND_BASE_URL}/api/tool-result`, {
      method: "POST",

      headers: { "Content-Type": "application/json" },

      body: JSON.stringify({ conversationId: currentConversationId, toolResults }),
    });

    data = (await resp2.json()) as BackendResponse;

    console.log("⬅ Backend after tool-result:", data);

    if (!resp2.ok) {
      console.error("[api/client] /api/tool-result error:", data);

      return { reply: null, conversationId: currentConversationId };
    }

    currentConversationId = data.conversationId ?? currentConversationId;
  }

  return { reply: data.reply ?? null, conversationId: currentConversationId };
}

// ─── Session ──────────────────────────────────────────────────────────────────

export interface EndSessionResult {
  ok: boolean;

  message?: string;

  error?: string;
}

/**

* End the current session — deletes the thread and agent on the backend,

* then re-creates a fresh agent ready for the next session.

*/

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

/**
 * Create a fresh thread and add the document text as context WITHOUT starting
 * a run. Because no run is created, the thread is immediately ready for the
 * first real /api/chat call — no race condition possible.
 */
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