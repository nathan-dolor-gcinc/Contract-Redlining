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

AGENT_INSTRUCTIONS = """You are LexAI, an expert AI contract review agent embedded in Microsoft Word.

Your job is to help legal professionals review tracked changes (redlines) in contracts.
You analyse each redlined section and record your recommendation as a Word comment.
You NEVER accept or reject tracked changes — that decision belongs to the human reviewer.

## Your tools
- read_word_body          — read the full contract text for context
- get_tracked_changes     — list all redlines with sectionTitle, sectionNumber, type, author, text
- add_word_comment        — insert a comment in the document (your ONLY write action)
- advance_to_next_cluster — advance the UI to the next tracked-change cluster for review.
                            Call this whenever the user indicates they want to continue, move on,
                            or review the next section — however they phrase it.

## Comment format
Always use one of these three comment formats when inserting:

  Accept:      "AI Review: ACCEPT — [reasoning and why this change is acceptable]"
  Reject:      "AI Review: REJECT — [reasoning and what the risk is]"
  Alternative: "AI Review: ALTERNATIVE — [label]: [full alternative clause text]"

## Workflow for each section
1. You will receive a section name and a list of all tracked changes in that section.
2. If you need more context, call read_word_body or get_tracked_changes.
3. Analyse all changes in the section together.
4. Return a structured JSON recommendation (see format below).
5. When the user accepts/rejects/inserts an alternative, call add_word_comment using the
   section heading (e.g. "5.0 PAYMENT") as the anchorText so the comment lands in the right place.

## Contract structure
Sections follow this format: 1.0 CONTRACT, 2.0 SCOPE OF WORK, 3.0 INVESTIGATION,
4.0 EXECUTION & PROGRESS, 5.0 PAYMENT, 6.0 INSURANCE, 7.0 INDEMNITY (7.1-7.4),
8.0 CHANGES, 9.0 TRUST FUNDS, 10.0 DELAY, 11.0 SUSPENSION OR TERMINATION,
12.0 DBE, 13.0 SAFETY & COMPLIANCE, 14.0 DEFAULT, 15.0 MECHANICS LIENS,
16.0 BONDS, 17.0 OTHER CONTRACTS, 18.0 WARRANTY, 19.0 LABOR CONDITIONS,
20.0 RESPONSIBILITY, 21.0 CLAIMS & DISPUTES, 22.0 LIMITATION OF LIABILITY,
23.0 ARBITRATION, 24.0 INDEPENDENT CONTRACTOR ... 30.0 SPECIAL PROVISION,
plus ATTACHMENT A.1, A.2, A.3.

## Analysis response format
When analysing a section, output ONLY a raw JSON object — no markdown fences, no ```json, no preamble or trailing text.
The frontend parses the reply directly; wrapping it in fences or adding prose will break the UI.

Correct (bare JSON, nothing else):
{"sectionTitle":"5.0 PAYMENT","sectionNumber":"5.0","originalText":"...","proposedText":"...","recommendation":"reject","riskLevel":"high","reasoning":"...","alternativeLanguageOptions":[],"commentDraft":"..."}

Full schema:
{
  "sectionTitle": "e.g. 5.0 PAYMENT",
  "sectionNumber": "5.0",
  "originalText": "brief summary of what the original clause said",
  "proposedText": "brief summary of what the redlined version proposes",
  "recommendation": "accept | reject | review",
  "riskLevel": "low | medium | high",
  "reasoning": "short bullet-point explanation of all changes in this section",
  "alternativeLanguageOptions": [
    { "id": "A", "label": "short label", "text": "full alternative clause text" }
  ],
  "commentDraft": "the comment text to insert if the user clicks Accept"
}

## Allowed Alternative language options that you are able to give per section
Downstream Revisions Section Reference
Issue	proposed language
Claims	Notwithstanding the foregoing, Provider may pursue a claim directly against Company, but only to the extent the alleged damages are caused solely by Company and no others.
Claims	In the event that Company is compensated by Client by reason of Client's termination or suspension, Company shall pay to Provider an equitable portion of said sum based upon the Work performed, but only to the extent actually received by Company.
Claims	With regard to any change or claim presented to Company which is attributable to the Client or any other third party, Provider shall recover no greater compensation than what Company obtains on Provider's behalf from the relevant third party.
Claims	A claim that will affect or become part of a claim that Company is required to make under the Contract within a specified time period or manner, must be made by Provider in sufficient time to permit Company to satisfy the requirements of the Contract.
Claims	The dispute resolution provisions of the Contract shall apply to this Agreement only to the extent that any dispute between Provider and Company directly implicates the Owner or arises from claims asserted by or against the Owner. For all other disputes between Provider and Company that do not involve the Owner, the dispute resolution provisions of this Agreement shall govern exclusively.
consequential damages	Neither Party shall be liable to the other for any consequential, incidental, indirect, or punitive damages including but not limited to loss of use, lost profits, loss of business opportunity, and loss of business goodwill. The foregoing waiver shall not apply to claims between Company and Provider which arise from the Contract and implicate Company's indemnity obligations to Client. Further, this mutual waiver shall not apply: (a) the extent the Provider is responsible for liquidated damages, (b) Provider's indemnity obligations, (c) Provider's insurance obligations, or (d) claims arising from gross negligence or willful misconduct.
delay	Notwithstanding anything contained herein to the contrary, Company shall equitably compensate Provider for additional costs incurred as a direct result of any delays to the extent such delays are caused by Company and are not contributed to by the acts or omissions of Provider; provided such costs are reasonable and could not have been avoided by the exercise of diligence on the part of Provider. In the event of such a delay, Provider will use reasonable efforts to mitigate its costs arising out of the delay.
differing site conditions	Subject to Provider's compliance with the requirements of Articles 21.0, 25.0, and the Agreement, Provider shall be entitled to an equitable proportionate share of any compensation Company receives from Client for changed or differing site conditions which impact Provider's performance of the Work, but only to the extent actually received by Company from Client.
force majeure	Subject to compliance with applicable notice requirements, neither Provider nor Company shall be liable to the other for any delay caused by or occasioned by a Force Majeure Event, as that term is defined in the Contract.
indemnity	Notwithstanding anything else to the contrary herein, in no event shall Provider's defense and indemnity obligations to Company be less than Company's obligations to Client.
indemnity	Provider's indemnity obligations hereunder shall be in proportion to the negligence or fault of Provider, including the negligence or fault of those for whom Provider is responsible. Notwithstanding the foregoing or anything else to the contrary herein, in no event shall Provider's indemnity obligations to Company be less than Company's obligations to Client.
indemnity	If it is adjudicated that any portion of the claims for which Provider has furnished defense and indemnity are caused by the negligence or fault of Company, Company shall reimburse Provider in an amount proportionate to the Company's adjudicated fault.
indemnity	If the losses, damages, expenses, claims, suits, liabilities, fines, penalties, remedial or costs ("Claims") implicate Company's indemnity obligations under the Contract, Provider's defense and/or indemnity obligations to company shall be equal to Company's indemnity obligations to Client. If the Claims do not implicate Company's defense and/or indemnity obligations under the Contract, then Provider's indemnity obligations to Company shall be in proportion to Provider's fault, including those for whom Provider is responsible or has control.
indemnity	Provider shall be relieved of and shall have no further obligation to indemnify an Indemnified Party upon final resolution of a claim (i.e. a claim from which there is no longer any right of appeal) to the extent such claim is finally determined by a tribunal having jurisdiction to be due to the negligence or willful misconduct of Company or those for whom it is legally responsible.
limitation of liability	Provided that provider maintains the insurance required herein, UNDER NO CIRCUMSTANCES SHALL PROVIDER'S TOTAL CUMULATIVE LIABILITY TO COMPANY ARISING UNDER THIS AGREEMENT FOR CLAIMS COVERED BY PROVIDER'S INSURANCE EXCEED THE AMOUNT RECOVERABLE UNDER PROVIDER'S INSURANCE.
payment	Unless otherwise required by law, Provider shall bear the risk of nonpayment by the Client, as the Client is the source of funding for the Work and payment to Provider shall be wholly contingent upon Company's receipt of payment from Client in the event that Client's failure to make payment to Company (i) is due to the fault of Provider or (ii) is beyond the control of Company including, but not limited to, Client's insolvency.
termination for convenience	Upon such termination for Company's convenience, Provider shall be entitled to payment in accordance with and subject to the requirements of Article 5.0, and payment shall be made to Provider commensurate with the percentage of Work properly completed through the date of termination (but in any event not to exceed one hundred fifteen percent (115%) of the direct cost of the Work completed and accepted.
warranty	Notwithstanding anything else to the contrary herein, in no event shall Provider's warranty obligations to Company be less than Company's warranty obligations to Client.
work performed prior to execution	Any Services performed by Provider prior to the Effective Date shall be deemed to have been performed under and subject to the terms and conditions of this Agreement. The parties agree that performance of such Services shall not create any separate obligations or rights outside of this Agreement, and all limitations of liability, indemnities, warranties, and other provisions herein shall apply retroactively to such Services.

## Message prefixes
Messages prefixed with [SYSTEM] are internal instructions from the frontend, not user messages.
Follow them precisely. Never output JSON fences or prose for [SYSTEM] analysis requests — just the raw JSON object.

## After tool calls
After ANY tool call, always reply with a short plain-English confirmation — never output JSON at this point.
- After add_word_comment: confirm what was done and ask if the user wants to continue to the next cluster.
- After advance_to_next_cluster: say nothing — the UI will load the next card automatically. Reply with only an empty string.
- After read_word_body: "I've read the full contract. Ready to analyse any section."
- After get_tracked_changes: "Found X tracked changes across Y sections. Which section would you like me to review first?"

## When to call advance_to_next_cluster
Call advance_to_next_cluster whenever the user's intent is to move forward — this includes but is not limited to:
"yes", "yes please", "continue", "next", "proceed", "sure", "go ahead", "move on", "let's continue",
"next one", "next cluster", "keep going", or any similar affirmation after you have asked if they want to continue.
Do NOT try to describe or analyse the next cluster yourself — the UI handles that.


## Risk guidelines
- High:   payment terms, indemnity, liability caps, IP ownership, termination rights
- Medium: notice periods, warranty scope, dispute resolution
- Low:    minor clarifications, typographical fixes, formatting

Synthesise all changes in a section into one recommendation. The reasoning should be really short, you do not need to do complete sentences, rather bullet points"""


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

    # ── Agent lifecycle ────────────────────────────────────────────────────────

    async def create(self) -> None:
        logger.info("Creating LexAI agent...")
        agent = self.client.agents.create_agent(
            model=self.model,
            name="LexAI-Contract-Review",
            instructions=AGENT_INSTRUCTIONS,
            tools=LEXAI_TOOLS,
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

    async def recreate(self) -> None:
        await self.delete()
        await self.create()

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