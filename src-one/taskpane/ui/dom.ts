// src/taskpane/ui/dom.ts
//
// Low-level DOM helpers with zero business logic.
// These are the only functions that should touch document.getElementById
// for generic operations — higher-level builders live in chat.ts / cards.ts.

// ─── Element helpers ──────────────────────────────────────────────────────────

export function showEl(id: string): void {
  const el = document.getElementById(id);
  if (el) el.style.display = "";
}

export function setText(id: string, val: string): void {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

export function scrollChat(): void {
  const chat = document.getElementById("chat-window");
  if (chat) chat.scrollTop = chat.scrollHeight;
}

// ─── Input helpers ────────────────────────────────────────────────────────────

export function autoResize(el: HTMLTextAreaElement): void {
  el.style.height = "auto";
  el.style.height = `${Math.min(el.scrollHeight, 120)}px`;
}

export function lockInput(): void {
  const input = document.getElementById("agent-input") as HTMLTextAreaElement | null;
  const btn = document.getElementById("run-agent") as HTMLButtonElement | null;

  if (input) {
    input.disabled = true;
    input.placeholder = "Session ended — reload to start a new review.";
  }
  if (btn) {
    btn.disabled = true;
  }
}

// ─── Escaping ─────────────────────────────────────────────────────────────────

export function esc(s?: string): string {
  return (s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ─── Document helpers ─────────────────────────────────────────────────────────

export function getDocumentName(): string {
  try {
    return Office.context.document.url?.split("/").pop() ?? "Active Document";
  } catch {
    return "Active Document";
  }
}
