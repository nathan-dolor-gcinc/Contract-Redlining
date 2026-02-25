// src/taskpane/ui/chat.ts
//
// Builds and inserts chat-window elements:
//   - assistant / user message bubbles
//   - system / separator messages
//   - typing indicator

import { esc, scrollChat } from "./dom";

// ─── Bubbles ──────────────────────────────────────────────────────────────────

export function appendAssistantBubble(text: string): void {
  appendMessage("assistant", text);
}

export function appendUserBubble(text: string): void {
  appendMessage("user", text);
}

function appendMessage(role: "user" | "assistant", text: string): void {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const wrapper = document.createElement("div");
  wrapper.className = `chat-message chat-message--${role}`;

  const label = document.createElement("div");
  label.className = "chat-label";
  label.textContent = role === "user" ? "You" : "LexAI";

  const bubble = document.createElement("div");
  bubble.className = "chat-bubble";
  bubble.innerHTML = esc(text).replace(/\n/g, "<br>");

  wrapper.appendChild(label);
  wrapper.appendChild(bubble);
  chat.appendChild(wrapper);
  scrollChat();
}

// ─── System message (horizontal-rule style) ───────────────────────────────────

export function appendSysMsg(text: string): void {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const div = document.createElement("div");
  div.className = "sys-msg";
  div.textContent = text;
  chat.appendChild(div);
  scrollChat();
}

// ─── Typing / loading indicator ───────────────────────────────────────────────

export function setLoading(show: boolean): void {
  document.getElementById("typing-indicator")?.remove();

  if (!show) return;

  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const el = document.createElement("div");
  el.id = "typing-indicator";
  el.className = "chat-message";
  el.innerHTML = `
<div class="chat-label">LexAI</div>
<div class="chat-bubble typing-dots">
<span></span><span></span><span></span>
</div>`;
  chat.appendChild(el);
  scrollChat();
}
