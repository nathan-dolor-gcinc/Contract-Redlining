/* global document, Office, Word */
/*
import { AIProjectClient } from "@azure/ai-projects";
import { AzureKeyCredential } from "@azure/core-auth";
import { DefaultAzureCredential } from "@azure/identity";


let client: InstanceType<typeof AIProjectClient>;
let threadId: string | null = null;

const foundry_api_key: string = process.env.FOUNDRY_API_KEY!;

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {

    // Show UI
    const appBody = document.getElementById("app-body");
    if (appBody) appBody.style.display = "flex";

    // Wire up Send button
    const sendButton = document.getElementById("run-agent");
    if (sendButton) sendButton.onclick = () => tryCatch(handleSendMessage);

    // Initialize client (correct for beta.4)
    const project = new AIProjectClient(
      "https://asasi-mcar2rkf-eastus2.services.ai.azure.com/api/projects/asasi-mcar2rkf-eastus2_project",
      new DefaultAzureCredential());

    // Create thread
    const thread = await client.agents.threads.create();
    threadId = thread.id;

    console.log("Thread created:", threadId);
  }
});

// Handle user sending a message
async function handleSendMessage() {
  const input = document.getElementById("agent-input") as HTMLTextAreaElement;
  const text = input.value.trim();
  if (!text) return;

  appendToChat("user", text);
  input.value = "";

  const reply = await sendToAgent(text);
  if (reply) appendToChat("assistant", reply);
}

// Send message to Azure Agent
async function sendToAgent(prompt: string): Promise<string | null> {
  if (!threadId) return null;

  // Add user message
  await client.agents.messages.create(threadId, "user", prompt);

  // Create run
  let run = await client.agents.runs.create(threadId, "asst_gGwnLmkX0JbtKhY327Ipagi6");

  // Poll until done
  while (run.status === "queued" || run.status === "in_progress") {
    await new Promise((resolve) => setTimeout(resolve, 1000));
    run = await client.agents.runs.get(threadId, run.id);
  }

  if (run.status === "failed") {
    console.error("Agent run failed:", run.lastError);
    return null;
  }

  // Retrieve messages
  const messages = await client.agents.messages.list(threadId, { order: "asc" });

  let lastAssistantMessage = "";

  for await (const m of messages) {
    if (m.role === "assistant") {
      const content = m.content.find((c) => c.type === "text");
      if (content) lastAssistantMessage = content.text.value;
    }
  }

  return lastAssistantMessage;
}

// Add messages to chat window
function appendToChat(role: string, text: string) {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const bubble = document.createElement("div");

  bubble.style.margin = "8px 0";
  bubble.style.padding = "8px";
  bubble.style.borderRadius = "6px";
  bubble.style.maxWidth = "80%";

  if (role === "user") {
    bubble.style.background = "#d0e7ff";
    bubble.style.alignSelf = "flex-end";
  } else {
    bubble.style.background = "#f1f1f1";
  }

  bubble.textContent = text;
  chat.appendChild(bubble);
  chat.scrollTop = chat.scrollHeight;
}

// Error wrapper
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
*/

/* global document, Office, Word */
/*
import { executeToolCall } from "./tools/dispatchTool";

let conversationId: string | null = null;
let backendBaseUrl = "http://localhost:3001"; // change to your deployed API URL

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {
    const appBody = document.getElementById("app-body");
    if (appBody) appBody.style.display = "flex";

    const sendButton = document.getElementById("run-agent");
    if (sendButton) sendButton.onclick = () => tryCatch(handleSendMessage);
  }
});

async function handleSendMessage() {
  const input = document.getElementById("agent-input") as HTMLTextAreaElement;
  const text = input.value.trim();
  if (!text) return;

  appendToChat("user", text);
  input.value = "";

  const reply = await sendToModel(text);
  if (reply) appendToChat("assistant", reply);
}


export async function sendToModel(prompt: string): Promise<string | null> {
  // 1) First request: send user prompt to backend
  let resp = await fetch(`${backendBaseUrl}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, conversationId }),
  });

  let data = await resp.json();
  if (!resp.ok) {
    console.error("Backend /api/chat error:", data);
    return null;
  }

  conversationId = data.conversationId ?? conversationId;

  // 2) Tool loop: as long as backend asks for tools, execute them and return results
  while (data.toolRequests?.length) {
    console.log("🔧 Tool requests from backend:", data.toolRequests);

    const toolResults = [];
    for (const toolCall of data.toolRequests) {
      console.log("▶ executing tool:", toolCall.function?.name);
      toolResults.push(await executeToolCall(toolCall));
    }

    const resp2 = await fetch(`${backendBaseUrl}/api/tool-result`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ conversationId, toolResults })
    });

    data = await resp2.json();
    console.log("⬅ backend after tool-result:", data);

    if (!resp.ok) {
      console.error("Backend /api/tool-result error:", data);
      return null;
    }

    conversationId = data.conversationId ?? conversationId;
  }

  // 3) Final reply (no more tool requests)
  return data.reply ?? null;
}


function appendToChat(role: string, text: string) {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const bubble = document.createElement("div");
  bubble.style.margin = "8px 0";
  bubble.style.padding = "8px";
  bubble.style.borderRadius = "6px";
  bubble.style.maxWidth = "80%";

  if (role === "user") {
    bubble.style.background = "#d0e7ff";
    bubble.style.alignSelf = "flex-end";
  } else {
    bubble.style.background = "#f1f1f1";
  }

  bubble.textContent = text;
  chat.appendChild(bubble);
  chat.scrollTop = chat.scrollHeight;
}

async function tryCatch(callback: () => Promise<void>) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
  */
 /* global document, Office, Word */
import { executeToolCall } from "./tools/dispatchTool";

let conversationId: string | null = null;
let backendBaseUrl = "http://localhost:3001"; // change to your deployed API URL

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {
    const appBody = document.getElementById("app-body");
    if (appBody) appBody.style.display = "flex";

    const sendButton = document.getElementById("run-agent");
    if (sendButton) sendButton.onclick = () => tryCatch(handleSendMessage);
  }
});

async function handleSendMessage() {
  const input = document.getElementById("agent-input") as HTMLTextAreaElement;
  const text = input.value.trim();
  if (!text) return;

  appendToChat("user", text);
  input.value = "";

  const reply = await sendToModel(text);
  if (reply) appendToChat("assistant", reply);
}

export async function sendToModel(prompt: string): Promise<string | null> {
  // 1) First request: send user prompt to backend
  let resp = await fetch(`${backendBaseUrl}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, conversationId }),
  });

  let data = await resp.json();
  if (!resp.ok) {
    console.error("Backend /api/chat error:", data);
    return null;
  }

  conversationId = data.conversationId ?? conversationId;

  // 2) Tool loop: as long as backend asks for tools, execute them and return results
  while (data.toolRequests?.length) {
    console.log("🔧 Tool requests from backend:", data.toolRequests);

    const toolResults: Array<{ tool_call_id: string; output: string }> = [];
    for (const toolCall of data.toolRequests) {
      console.log("▶ executing tool:", toolCall.function?.name);
      toolResults.push(await executeToolCall(toolCall));
    }

    const resp2 = await fetch(`${backendBaseUrl}/api/tool-result`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ conversationId, toolResults }),
    });

    data = await resp2.json();
    console.log("⬅ backend after tool-result:", data);

    // ✅ FIX: check resp2.ok (not resp.ok)
    if (!resp2.ok) {
      console.error("Backend /api/tool-result error:", data);
      return null;
    }

    conversationId = data.conversationId ?? conversationId;
  }

  // 3) Final reply (no more tool requests)
  return data.reply ?? null;
}

function appendToChat(role: string, text: string) {
  const chat = document.getElementById("chat-window");
  if (!chat) return;

  const bubble = document.createElement("div");
  bubble.style.margin = "8px 0";
  bubble.style.padding = "8px";
  bubble.style.borderRadius = "6px";
  bubble.style.maxWidth = "80%";

  if (role === "user") {
    bubble.style.background = "#d0e7ff";
    bubble.style.alignSelf = "flex-end";
  } else {
    bubble.style.background = "#f1f1f1";
  }

  bubble.textContent = text;
  chat.appendChild(bubble);
  chat.scrollTop = chat.scrollHeight;
}

async function tryCatch(callback: () => Promise<void>) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}