// server/index.ts
/*
import "dotenv/config";
import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import { TOOL_DEFS } from "./tools/toolDefs";
import { AIProjectClient } from "@azure/ai-projects";
import { DefaultAzureCredential } from "@azure/identity";

const app = express();
app.use(cors());
app.use(express.json());

const PROJECT_ENDPOINT = process.env.AZURE_AI_PROJECT_ENDPOINT_STRING!;
const DEPLOYMENT_NAME = process.env.DEPLOYMENT_GPT_MODEL!;
const OPENAI_API_VERSION = process.env.OPENAI_API_VERSION || "2024-10-21";
const INSTRUCTIONS_PATH = path.resolve(__dirname, "instructions.txt");
const SYSTEM_INSTRUCTIONS = fs.readFileSync(INSTRUCTIONS_PATH, "utf-8").trim();

// Validate required env vars early (helps catch undefined issues fast)
if (!PROJECT_ENDPOINT) throw new Error("Missing AZURE_AI_PROJECT_ENDPOINT_STRING");
if (!DEPLOYMENT_NAME) throw new Error("Missing DEPLOYMENT_GPT_MODEL");
if (!SYSTEM_INSTRUCTIONS) {
  throw new Error("agent-instructions.txt is empty");
}

const project = new AIProjectClient(PROJECT_ENDPOINT, new DefaultAzureCredential());

// Create the Azure OpenAI client once (recommended)
const openaiPromise = project.getAzureOpenAIClient({ apiVersion: OPENAI_API_VERSION }); // [1](https://furotmark.github.io/2025/09/23/Azure-AI-Foundry-Message-History.html)

type ChatMessage = {
  role: "system" | "user" | "assistant";
  content: string;
};

// In-memory conversations keyed by conversationId
const conversations = new Map<string, ChatMessage[]>();

// Configure your default system prompt once
const SYSTEM_PROMPT: ChatMessage = {
  role: "system",
  content: SYSTEM_INSTRUCTIONS
};


// Keep last N messages (prevents token/context explosion)
const MAX_MESSAGES = 30;

// Helper: get or initialize conversation
function getConversation(conversationId: string): ChatMessage[] {
  if (!conversations.has(conversationId)) {
    conversations.set(conversationId, [SYSTEM_PROMPT]);
  }
  return conversations.get(conversationId)!;
}

// Helper: trim message history (keep system + last N-1)
function trimConversation(history: ChatMessage[]): ChatMessage[] {
  if (history.length <= MAX_MESSAGES) return history;

  const system = history[0]?.role === "system" ? [history[0]] : [SYSTEM_PROMPT];
  const tail = history.slice(-1 * (MAX_MESSAGES - system.length));
  return [...system, ...tail];
}

// Health check
app.get("/health", (_req, res) => res.json({ ok: true }));

// Chat endpoint with memory
app.post("/api/chat", async (req, res) => {
  try {
    const prompt = String(req.body?.prompt ?? "").trim();
    if (!prompt) return res.status(400).json({ error: "Missing prompt" });

    // conversationId from client; if missing, create one
    // (client should store and reuse this value)
    let conversationId = String(req.body?.conversationId ?? "").trim();
    if (!conversationId) {
      conversationId = `conv_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    }

    // Fetch history and add user message
    const history = getConversation(conversationId);
    history.push({ role: "user", content: prompt });

    // Trim history to safe size
    const trimmed = trimConversation(history);
    conversations.set(conversationId, trimmed);

    const openai = await openaiPromise;

    // Call the model with the FULL message history
    const response = await openai.chat.completions.create({
      model: DEPLOYMENT_NAME,
      messages: trimmed,
      tools: TOOL_DEFS,
      tool_choice: { type: "function", function: { name: "read_word_body" } }, //"auto",
      temperature: 0.2
    }); // messages[] usage shown in the Azure AI Projects sample [1](https://furotmark.github.io/2025/09/23/Azure-AI-Foundry-Message-History.html)[2](https://github.com/Azure/azure-sdk-for-js/blob/main/sdk/openai/openai/README.md)

    const reply =
      response.choices?.[0]?.message?.content ?? "(No response text returned)";

      
    const msg: any = response.choices?.[0]?.message;
    console.log("MODEL MESSAGE:", JSON.stringify(msg, null, 2));

    // Store assistant message back into history
    trimmed.push({ role: "assistant", content: reply });
    conversations.set(conversationId, trimConversation(trimmed));

    // Return reply + conversationId so the client keeps using it
    res.json({ reply, conversationId });
  } catch (err: any) {
    console.error(err);
    res.status(500).json({ error: err?.message ?? String(err) });
  }
});

// Optional: reset conversation memory
app.post("/api/reset", (req, res) => {
  const conversationId = String(req.body?.conversationId ?? "").trim();
  if (!conversationId) return res.status(400).json({ error: "Missing conversationId" });

  conversations.delete(conversationId);
  res.json({ ok: true });
});

app.listen(3001, () => console.log("Backend listening on http://localhost:3001"));
*/
// server/index.ts
import "dotenv/config";
import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import { TOOL_DEFS } from "./tools/toolDefs";
import { AIProjectClient } from "@azure/ai-projects";
import { DefaultAzureCredential } from "@azure/identity";

const app = express();
app.use(cors());
app.use(express.json());

const PROJECT_ENDPOINT = process.env.AZURE_AI_PROJECT_ENDPOINT_STRING!;
const DEPLOYMENT_NAME = process.env.DEPLOYMENT_GPT_MODEL!;
const OPENAI_API_VERSION = process.env.OPENAI_API_VERSION || "2024-10-21";
const INSTRUCTIONS_PATH = path.resolve(__dirname, "instructions.txt");
const SYSTEM_INSTRUCTIONS = fs.readFileSync(INSTRUCTIONS_PATH, "utf-8").trim();

// Validate required env vars early
if (!PROJECT_ENDPOINT) throw new Error("Missing AZURE_AI_PROJECT_ENDPOINT_STRING");
if (!DEPLOYMENT_NAME) throw new Error("Missing DEPLOYMENT_GPT_MODEL");
if (!SYSTEM_INSTRUCTIONS) throw new Error("instructions.txt is empty");

const project = new AIProjectClient(PROJECT_ENDPOINT, new DefaultAzureCredential());
const openaiPromise = project.getAzureOpenAIClient({ apiVersion: OPENAI_API_VERSION });

// ---- Message types that support tool calling ----
type ToolCall = {
  id: string;
  type: "function";
  function: { name: string; arguments?: string };
};

type HistoryMessage =
  | { role: "system" | "user" | "assistant"; content: string | null; tool_calls?: ToolCall[] }
  | { role: "tool"; tool_call_id: string; content: string };

// In-memory conversations keyed by conversationId
const conversations = new Map<string, HistoryMessage[]>();

// System prompt
const SYSTEM_PROMPT: HistoryMessage = {
  role: "system",
  content: SYSTEM_INSTRUCTIONS
};

// Keep last N messages (prevents token/context explosion)
const MAX_MESSAGES = 30;

// Helper: get or initialize conversation
function getConversation(conversationId: string): HistoryMessage[] {
  if (!conversations.has(conversationId)) {
    conversations.set(conversationId, [SYSTEM_PROMPT]);
  }
  return conversations.get(conversationId)!;
}

// Helper: trim message history (keep system + last N-1)
function trimConversation(history: HistoryMessage[]): HistoryMessage[] {
  if (history.length <= MAX_MESSAGES) return history;

  const system =
    history[0]?.role === "system"
      ? [history[0]]
      : [SYSTEM_PROMPT];

  const tail = history.slice(-1 * (MAX_MESSAGES - system.length));
  return [...system, ...tail];
}

// Create a conversationId if missing
function ensureConversationId(maybeId: string | undefined | null): string {
  const id = String(maybeId ?? "").trim();
  if (id) return id;
  return `conv_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

// Utility: call model with tools enabled
async function callModelWithHistory(history: HistoryMessage[]) {
  const openai = await openaiPromise;

  return openai.chat.completions.create({
    model: DEPLOYMENT_NAME,
    messages: history as any, // SDK expects ChatCompletionMessageParam-like objects
    tools: TOOL_DEFS as any,
    tool_choice: "auto",
    temperature: 0.2
  });
}

// Health check
app.get("/health", (_req, res) => res.json({ ok: true }));

/**
 * POST /api/chat
 * Body: { prompt: string, conversationId?: string }
 * Returns:
 *  - { reply, conversationId } OR
 *  - { toolRequests, conversationId } when the model requests tools
 */
app.post("/api/chat", async (req, res) => {
  try {
    const prompt = String(req.body?.prompt ?? "").trim();
    if (!prompt) return res.status(400).json({ error: "Missing prompt" });

    const conversationId = ensureConversationId(req.body?.conversationId);

    const history = getConversation(conversationId);
    history.push({ role: "user", content: prompt });
    const trimmed = trimConversation(history);
    conversations.set(conversationId, trimmed);

    const response = await callModelWithHistory(trimmed);

    const msg = response.choices?.[0]?.message;
    console.log("MODEL MESSAGE:", JSON.stringify(msg, null, 2));

    const toolCalls: ToolCall[] = (msg as any)?.tool_calls ?? [];

    // IMPORTANT: store the assistant message that contains tool_calls
    // (content may be null when tool_calls exist)
    if (msg) {
      trimmed.push({
        role: "assistant",
        content: (msg as any).content ?? null,
        tool_calls: toolCalls.length ? toolCalls : undefined
      });
      conversations.set(conversationId, trimConversation(trimmed));
    }

    // If tools requested, return them to the client to execute
    if (toolCalls.length) {
      return res.json({ conversationId, toolRequests: toolCalls });
    }

    // Otherwise return final reply text
    const reply = (msg as any)?.content ?? "(No response text returned)";
    return res.json({ conversationId, reply });
  } catch (err: any) {
    console.error(err);
    res.status(500).json({ error: err?.message ?? String(err) });
  }
});

/**
 * POST /api/tool-result
 * Body: {
 *   conversationId: string,
 *   toolResults: Array<{ tool_call_id: string, output: string }>
 * }
 *
 * Returns:
 *  - { reply, conversationId } OR
 *  - { toolRequests, conversationId } if the model requests more tools
 */
app.post("/api/tool-result", async (req, res) => {
  try {
    const conversationId = String(req.body?.conversationId ?? "").trim();
    if (!conversationId) return res.status(400).json({ error: "Missing conversationId" });

    const toolResults = req.body?.toolResults;
    if (!Array.isArray(toolResults) || toolResults.length === 0) {
      return res.status(400).json({ error: "Missing toolResults[]" });
    }

    const history = getConversation(conversationId);

    // Append tool outputs to history
    for (const r of toolResults) {
      const tool_call_id = String(r?.tool_call_id ?? "").trim();
      const output = typeof r?.output === "string" ? r.output : JSON.stringify(r?.output ?? "");

      if (!tool_call_id) continue;

      history.push({
        role: "tool",
        tool_call_id,
        content: output
      });
    }

    const trimmed = trimConversation(history);
    conversations.set(conversationId, trimmed);

    // Call the model again so it can use tool output
    const response = await callModelWithHistory(trimmed);

    const msg = response.choices?.[0]?.message;
    console.log("MODEL MESSAGE AFTER TOOL:", JSON.stringify(msg, null, 2));

    const toolCalls: ToolCall[] = (msg as any)?.tool_calls ?? [];

    // Store assistant message (may include more tool_calls)
    if (msg) {
      trimmed.push({
        role: "assistant",
        content: (msg as any).content ?? null,
        tool_calls: toolCalls.length ? toolCalls : undefined
      });
      conversations.set(conversationId, trimConversation(trimmed));
    }

    // If model asks for more tools, return those
    if (toolCalls.length) {
      return res.json({ conversationId, toolRequests: toolCalls });
    }

    // Else final reply
    const reply = (msg as any)?.content ?? "(No response text returned)";
    return res.json({ conversationId, reply });
  } catch (err: any) {
    console.error(err);
    res.status(500).json({ error: err?.message ?? String(err) });
  }
});

// Optional: reset conversation memory
app.post("/api/reset", (req, res) => {
  const conversationId = String(req.body?.conversationId ?? "").trim();
  if (!conversationId) return res.status(400).json({ error: "Missing conversationId" });

  conversations.delete(conversationId);
  res.json({ ok: true });
});

app.listen(3001, () => console.log("Backend listening on http://localhost:3001"));