/* global document, Office */
// src/taskpane/taskpane.ts
//
// Entry point — kept intentionally thin.
// Responsibilities:
//   1. Wait for Office.js to be ready
//   2. Show the app shell
//   3. Kick off the document scan / review flow
//   4. Wire up the chat input and send button
//
// Business logic lives in:
//   review.ts      → section-by-section redline review
//   api/client.ts  → backend communication
//   state/session  → runtime state
//   ui/*           → all DOM manipulation

import { initializeScan } from "./review";
import { session } from "./state/session";
import { sendPrompt, endSession } from "./api/client";
import { appendUserBubble, appendAssistantBubble, appendSysMsg, setLoading } from "./ui/chat";
import { autoResize, lockInput } from "./ui/dom";

// ─── Office.onReady ───────────────────────────────────────────────────────────

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) return;

  const appBody = document.getElementById("app-body");
  if (appBody) appBody.style.display = "flex";

  wireInputEvents();

  // Kick off the initial document scan — errors surfaced in the chat window
  initializeScan().catch((err) => {
    console.error("[taskpane] initializeScan failed:", err);
    appendAssistantBubble(`⚠ Could not read document: ${(err as Error).message}`);
  });
});

// ─── Input wiring ─────────────────────────────────────────────────────────────

function wireInputEvents(): void {
  const input = document.getElementById("agent-input") as HTMLTextAreaElement | null;
  const sendBtn = document.getElementById("run-agent") as HTMLButtonElement | null;

  if (!input || !sendBtn) return;

  // Auto-resize textarea as the user types
  input.addEventListener("input", () => autoResize(input));

  // Send on Enter (without Shift)
  input.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage(input).catch(console.error);
    }
  });

  // Send on button click
  sendBtn.addEventListener("click", () => {
    handleSendMessage(input).catch(console.error);
  });
}

// ─── Message handler ──────────────────────────────────────────────────────────

async function handleSendMessage(input: HTMLTextAreaElement): Promise<void> {
  const text = input.value.trim();
  if (!text) return;

  // ── "end session" intercept ───────────────────────────────────────────────
  if (text.toLowerCase() === "end session") {
    input.value = "";
    autoResize(input);
    await handleEndSession();
    return;
  }

  if (session.sessionEnded) return;

  appendUserBubble(text);
  input.value = "";
  autoResize(input);
  setLoading(true);

  const { reply, conversationId } = await sendPrompt(text, session.conversationId);
  session.conversationId = conversationId;

  setLoading(false);
  if (reply) appendAssistantBubble(reply);
}

// ─── End session ──────────────────────────────────────────────────────────────

async function handleEndSession(): Promise<void> {
  if (session.sessionEnded) return;

  appendUserBubble("end session");
  appendSysMsg("Ending session — deleting agent and clearing conversation…");
  setLoading(true);

  const result = await endSession(session.conversationId);

  setLoading(false);

  if (result.ok) {
    session.sessionEnded = true;
    session.conversationId = null;

    lockInput();
    appendSysMsg("✓ Session ended. Agent and conversation history deleted from Azure.");
    appendAssistantBubble(
      "This session has been closed and all data removed from the server. " +
        "Reload the add-in to start a new review session."
    );
  } else {
    appendAssistantBubble(`⚠ Could not end session cleanly: ${result.error}`);
  }
}
