"""
LexAI Agent Service
Handles agent lifecycle, tool definitions, system prompt, and run polling.
"""

import asyncio
import json
import logging
import os
import re

from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential
from openai import AzureOpenAI

logger = logging.getLogger(__name__)

POLL_INTERVAL_S = 1.5
MAX_POLLS       = 100
TERMINAL        = {"completed", "failed", "cancelled", "expired"}

# ─── Tool definitions ──────────────────────────────────────────────────────────

LEXAI_TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "read_word_body",
            "description": (
                "Read the full plain text of the active Word document. "
                "Use this to understand the full contract before analysing a section."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "maxChars": {
                        "type": "integer",
                        "description": "Maximum characters to return (default 60000).",
                    }
                },
                "required": [],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "get_tracked_changes",
            "description": (
                "Return all tracked changes (redlines) in the active Word document as a JSON array. "
                "Each item includes: id, type, author, date, text, sectionTitle, sectionNumber, paragraphContext."
            ),
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "advance_to_next_cluster",
            "description": (
                "Call this when the user confirms they want to move on, continue, or review the next "
                "tracked-change cluster. This advances the UI to load and display the next analysis card. "
                "Use this after completing an action (accept/reject/alternative) when the user indicates "
                "they are ready to proceed, regardless of how they phrase it."
            ),
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "add_word_comment",
            "description": (
                "Insert a comment in the Word document anchored to a specific piece of text. "
                "This is the ONLY way the agent writes to the document. "
                "Use the section heading or the first distinctive phrase of the redlined text as the anchor."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "anchorText": {
                        "type": "string",
                        "description": (
                            "The exact text in the document to attach the comment to. "
                            "Use the section heading (e.g. '5.0 PAYMENT') or the first few words of the redlined text."
                        ),
                    },
                    "commentText": {
                        "type": "string",
                        "description": "The comment body. Format: 'AI Review: ACCEPT/REJECT/ALTERNATIVE — [reasoning]'",
                    },
                    "occurrence": {
                        "type": "integer",
                        "description": "Zero-based index if anchorText appears multiple times (default 0).",
                    },
                    "matchCase":      {"type": "boolean", "description": "Case-sensitive search (default false)."},
                    "matchWholeWord": {"type": "boolean", "description": "Whole-word match (default false)."},
                },
                "required": ["anchorText", "commentText"],
            },
        },
    },
]

# ─── System prompt ─────────────────────────────────────────────────────────────

with open(os.path.join(os.path.dirname(__file__), "instructions.txt"), encoding="utf-8") as _f:
    AGENT_INSTRUCTIONS = _f.read()




# ─── Agent service class ───────────────────────────────────────────────────────

class LexAIAgent:

    def __init__(self):
        self.project_endpoint = os.environ["AZURE_FOUNDRY_PROJECT_ENDPOINT"]
        self.model            = os.getenv("AZURE_MODEL_DEPLOYMENT", "gpt-4o")
        self.agent_id: str | None = None

        # Foundry agent client — used for the full agent/thread/run lifecycle
        self.client = AIProjectClient(
            endpoint=self.project_endpoint,
            credential=DefaultAzureCredential(),
        )

        # Plain AzureOpenAI client — used for stateless one-shot calls (e.g. clustering)
        # that don't need an agent thread or tool loop.
        self.openai = AzureOpenAI(
            azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
            api_key=os.environ["FOUNDRY_API_KEY"],
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview"),
        )

    # ── Knowledge base setup ───────────────────────────────────────────────────

    def setup_vector_store(self, file_paths: list[str], name: str = "LexAI-Knowledge-Base") -> str:
        """
        Upload txt files and create a vector store via the OpenAI client,
        matching the Microsoft docs pattern. Run once — persist the returned ID
        in LEXAI_VECTOR_STORE_ID so it is reused on subsequent restarts.
        """
        openai_client = self.client.get_openai_client()

        vector_store = openai_client.vector_stores.create(name=name)
        logger.info(f"Vector store created: {vector_store.id}")

        for path in file_paths:
            with open(path, "rb") as f:
                openai_client.vector_stores.files.upload_and_poll(
                    vector_store_id=vector_store.id,
                    file=f,
                )
            logger.info(f"Uploaded {path} → {vector_store.id}")

        return vector_store.id

    # ── Agent lifecycle ────────────────────────────────────────────────────────

    async def create(self, vector_store_id: str | None = None) -> None:
        """
        Create the LexAI agent. Pass a vector_store_id to attach the knowledge
        base (file_search). If omitted, falls back to LEXAI_VECTOR_STORE_ID env var.
        """
        vs_id = vector_store_id or os.getenv("LEXAI_VECTOR_STORE_ID")

        # Build tools list — prepend FileSearchTool if we have a vector store
        tools = list(LEXAI_TOOLS)
        if vs_id:
            tools = [{"type": "file_search", "file_search": {"vector_store_ids": [vs_id]}}] + tools
            logger.info(f"Attaching vector store: {vs_id}")
        else:
            logger.warning(
                "No vector store ID provided — file_search will have no knowledge base. "
                "Set LEXAI_VECTOR_STORE_ID or pass vector_store_id to create()."
            )

        logger.info("Creating LexAI agent...")
        agent = self.client.agents.create_agent(
            model=self.model,
            name="LexAI-Contract-Review",
            instructions=AGENT_INSTRUCTIONS,
            tools=tools,
        )
        self.agent_id = agent.id
        logger.info(f"Agent created: {self.agent_id}")

    async def delete(self) -> None:
        if not self.agent_id:
            return
        try:
            self.client.agents.delete_agent(self.agent_id)
            logger.info(f"Agent deleted: {self.agent_id}")
            self.agent_id = None
        except Exception as e:
            logger.warning(f"Could not delete agent: {e}")

    async def recreate(self, vector_store_id: str | None = None) -> None:
        await self.delete()
        await self.create(vector_store_id=vector_store_id)

    # ── Helpers ────────────────────────────────────────────────────────────────

    @staticmethod
    def _status(run) -> str:
        s = run.status
        return s.value.lower() if hasattr(s, "value") else str(s).lower()

    async def poll_run(self, thread_id: str, run_id: str):
        """Poll until run reaches completed or requires_action."""
        for _ in range(MAX_POLLS):
            await asyncio.sleep(POLL_INTERVAL_S)
            run = self.client.agents.runs.get(thread_id=thread_id, run_id=run_id)
            status = self._status(run)
            if status == "completed":
                return run
            if status == "requires_action":
                return run
            if status in TERMINAL:
                last_error = getattr(run, "last_error", None)
                raise RuntimeError(f"Run {run_id} ended: {status} — {last_error}")
        raise TimeoutError(f"Run {run_id} timed out after {MAX_POLLS} polls")

    async def get_last_assistant_reply(self, thread_id: str) -> str:
        messages = self.client.agents.messages.list(thread_id=thread_id)
        for msg in messages:
            if msg.role != "assistant":
                continue
            text_block = next((c for c in msg.content if c.type == "text"), None)
            if text_block:
                return text_block.text.value
        return ""

    async def build_step_response(self, thread_id: str, run) -> dict:
        if self._status(run) == "completed":
            reply = await self.get_last_assistant_reply(thread_id)
            return {"conversationId": thread_id, "reply": reply}

        tool_calls_raw = getattr(
            getattr(getattr(run, "required_action", None), "submit_tool_outputs", None),
            "tool_calls",
            [],
        ) or []

        tool_requests = []
        for tc in tool_calls_raw:
            logger.info(f"[Tool Request] thread={thread_id} | tool={tc.function.name}")
            if tc.function.arguments:
                logger.info(f"  Arguments: {tc.function.arguments}")
            tool_requests.append({
                "id": tc.id,
                "function": {
                    "name": tc.function.name,
                    "arguments": tc.function.arguments,
                },
            })

        return {"conversationId": thread_id, "toolRequests": tool_requests}

    # ── Core operations ────────────────────────────────────────────────────────

    def prime(self, content: str) -> str:
        """Create a thread and post context without running the agent. Returns thread_id."""
        thread = self.client.agents.threads.create()
        self.client.agents.messages.create(
            thread_id=thread.id, role="user", content=content
        )
        logger.info(f"[prime] thread={thread.id} | context added (no run)")
        return thread.id

    async def chat(self, prompt: str, conversation_id: str | None = None) -> dict:
        """Send a user message and run the agent. Returns reply or tool requests."""
        if not self.agent_id:
            raise RuntimeError("Agent not ready")

        thread_id = conversation_id or self.client.agents.threads.create().id
        logger.info(f"[chat] thread={thread_id} | \"{prompt[:80]}...\"")

        self.client.agents.messages.create(
            thread_id=thread_id, role="user", content=prompt
        )
        run = self.client.agents.runs.create(
            thread_id=thread_id, agent_id=self.agent_id
        )
        finished = await self.poll_run(thread_id, run.id)
        return await self.build_step_response(thread_id, finished)

    async def submit_tool_outputs(
        self, conversation_id: str, tool_results: list[dict]
    ) -> dict:
        """Submit tool outputs to a waiting run. Returns reply or next tool requests."""
        thread_id = conversation_id
        logger.info(f"[tool-result] thread={thread_id} | {len(tool_results)} result(s)")

        runs_list = self.client.agents.runs.list(thread_id=thread_id)
        active_run = None
        for r in runs_list:
            if self._status(r) in {"requires_action", "in_progress"}:
                active_run = r
                break

        if not active_run:
            raise LookupError("No active run waiting for tool outputs")

        self.client.agents.runs.submit_tool_outputs(
            thread_id=thread_id,
            run_id=active_run.id,
            tool_outputs=[
                {"tool_call_id": tr["tool_call_id"], "output": tr["output"]}
                for tr in tool_results
            ],
        )
        finished = await self.poll_run(thread_id, active_run.id)
        return await self.build_step_response(thread_id, finished)

    def delete_thread(self, conversation_id: str) -> None:
        try:
            self.client.agents.threads.delete(conversation_id)
            logger.info(f"Thread deleted: {conversation_id}")
        except Exception as e:
            logger.warning(f"Could not delete thread: {e}")

    # ── Direct model inference (no agent, no thread, no tool loop) ─────────────

    async def cluster_changes(
        self,
        section_number: str,
        section_title: str,
        changes: list[dict],
    ) -> list[dict]:
        """
        Call the model directly via AzureOpenAI to group tracked changes into
        semantically related clusters, and select a short display snippet per cluster.

        Returns a list of cluster dicts:
          [
            { "indices": [0, 2], "snippet": "...1-3 sentences containing these changes..." },
            { "indices": [1],    "snippet": "...sentence containing change 1..." },
          ]

        snippet is chosen by the model — the minimal surrounding text from
        paragraphContext that gives the reviewer enough context. Used for display
        in the review card instead of the full paragraph.

        Falls back to one-change-per-cluster with a client-side snippet on any error.
        """
        n = len(changes)
        change_list = []
        for i, c in enumerate(changes):
            para = (c.get("paragraphContext") or "")
            text = (c.get("text") or "")
            # offsetInParagraph helps the model detect that nearby changes are one atomic edit
            offset = para.lower().find(text.lower()[:40]) if text else -1
            change_list.append({
                "index": i,
                "type": c.get("type", ""),
                "author": c.get("author", ""),
                "text": text[:200],
                "paragraphContext": para[:400],
                "offsetInParagraph": offset,
            })

        prompt = (
            f'You are analysing tracked changes in contract section "{section_title}" ({section_number}).\n\n'
            f"## Task\n"
            f"1. Group the {n} changes into semantically related clusters for review.\n"
            f"2. For each cluster, extract a SHORT snippet (1-3 sentences max) from "
            f"paragraphContext that contains all the changed text in that cluster. "
            f"The snippet must be a verbatim substring of paragraphContext — do not paraphrase or invent text. "
            f"Include just enough surrounding sentence context for the reviewer to understand the change. "
            f"Do not include the entire paragraph.\n\n"
            f"## Clustering rules\n"
            f"- Changes with offsetInParagraph within ~100 characters of each other are almost certainly "
            f"one atomic edit — put them in the same cluster.\n"
            f"- Changes that affect different obligations, parties, or concepts go in separate clusters "
            f"even if they share a paragraph.\n"
            f"- Changes that only make sense together (e.g. two edits forming one proposal) belong in the same cluster.\n\n"
            f"## Response format\n"
            f"Respond with ONLY a raw JSON array. Every index 0–{n-1} must appear exactly once.\n"
            f'Example: [{{"indices":[0,2],"snippet":"sentence containing changes 0 and 2"}},{{"indices":[1],"snippet":"sentence containing change 1"}}]\n'
            f"No prose, no markdown, no explanation — just the array.\n\n"
            f"Changes:\n{json.dumps(change_list, indent=2)}"
        )

        try:
            response = self.openai.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are a contract analysis assistant. "
                            "Respond only with the raw JSON array requested. "
                            "Snippets must be verbatim substrings of the provided paragraphContext."
                        ),
                    },
                    {
                        "role": "user",
                        "content": prompt,
                    },
                ],
                max_completion_tokens=2000,
            )

            raw = response.choices[0].message.content or ""
            cleaned = raw.replace("```json", "").replace("```", "").strip()
            clusters: list[dict] = json.loads(cleaned)

            # Validate structure and index coverage
            seen: set[int] = set()
            for cluster in clusters:
                if not isinstance(cluster.get("indices"), list) or "snippet" not in cluster:
                    raise ValueError(f"Malformed cluster object: {cluster}")
                for idx in cluster["indices"]:
                    if not isinstance(idx, int) or idx < 0 or idx >= n or idx in seen:
                        raise ValueError(f"Invalid index {idx}")
                    seen.add(idx)
            if len(seen) != n:
                raise ValueError(f"Cluster response covered {len(seen)} of {n} changes")

            # Verify each snippet is a verbatim substring — fall back per-cluster if not
            for cluster in clusters:
                snippet = cluster.get("snippet", "")
                first_para = (changes[cluster["indices"][0]].get("paragraphContext") or "")
                if snippet and snippet.lower() not in first_para.lower():
                    logger.warning(f"[cluster] Snippet not verbatim, using fallback. snippet='{snippet[:60]}'")
                    cluster["snippet"] = _extract_snippet_fallback(
                        first_para, [changes[i] for i in cluster["indices"]]
                    )

            logger.info(f"[cluster] § {section_number} — {n} changes → {len(clusters)} clusters")
            return clusters

        except Exception as e:
            logger.warning(f"[cluster] Falling back to one-per-cluster: {e}")
            return [
                {
                    "indices": [i],
                    "snippet": _extract_snippet_fallback(
                        changes[i].get("paragraphContext") or "", [changes[i]]
                    ),
                }
                for i in range(n)
            ]


def _extract_snippet_fallback(paragraph: str, changes: list[dict]) -> str:
    """
    Pure-Python fallback snippet extractor. Finds the sentences in `paragraph`
    that contain any of the change texts, expands by one sentence each side for
    context, and returns the result capped at 500 chars.
    """
    if not paragraph:
        return ""

    sentences = re.split(r'(?<=[.!?])\s+', paragraph.strip())
    if len(sentences) <= 2:
        return paragraph[:500]

    hit_indices: set[int] = set()
    for c in changes:
        needle = (c.get("text") or "").strip().lower()[:60]
        if not needle:
            continue
        for si, sent in enumerate(sentences):
            if needle in sent.lower():
                hit_indices.add(si)

    if not hit_indices:
        return " ".join(sentences[:2])[:500]

    min_i = max(0, min(hit_indices) - 1)
    max_i = min(len(sentences) - 1, max(hit_indices) + 1)
    return " ".join(sentences[min_i:max_i + 1])[:500]