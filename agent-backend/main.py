"""
LexAI — Contract Redlining API
Azure AI Foundry Agent backend (Python / FastAPI)

Run with:
    python main.py

Requires:
    pip install fastapi uvicorn python-dotenv azure-ai-projects azure-identity
    npx office-addin-dev-certs install   (one-time, same as the JS server)
"""

import logging
import os
import subprocess
from contextlib import asynccontextmanager

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from agent_service import LexAIAgent, _extract_snippet_fallback

load_dotenv()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ─── Singleton agent ───────────────────────────────────────────────────────────

_agent: LexAIAgent | None = None


def get_agent() -> LexAIAgent:
    global _agent
    if _agent is None:
        _agent = LexAIAgent()
    return _agent


# ─── Lifespan (create agent on startup, delete on shutdown) ───────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    agent = get_agent()
    await agent.create()
    yield
    await agent.delete()


# ─── App ───────────────────────────────────────────────────────────────────────

app = FastAPI(title="LexAI API", version="1.0.0", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://localhost:3000",
        "https://localhost",
    ],
    allow_origin_regex=r"https://.*\.azurewebsites\.net",
    allow_methods=["POST", "DELETE", "OPTIONS"],
    allow_headers=["*"],
)


# ─── Request / response models ─────────────────────────────────────────────────

class PrimeRequest(BaseModel):
    content: str


class ChatRequest(BaseModel):
    prompt: str
    conversationId: str | None = None


class ToolResult(BaseModel):
    tool_call_id: str
    output: str


class ToolResultRequest(BaseModel):
    conversationId: str
    toolResults: list[ToolResult]


class SessionDeleteRequest(BaseModel):
    conversationId: str | None = None


class ChangeItem(BaseModel):
    id: str = ""
    type: str = ""
    author: str = ""
    text: str = ""
    paragraphContext: str = ""


class ClusterRequest(BaseModel):
    sectionNumber: str
    sectionTitle: str
    changes: list[ChangeItem]


# ─── Routes ────────────────────────────────────────────────────────────────────

@app.get("/health")
async def health():
    return {"status": "ok", "service": "LexAI"}


@app.post("/api/prime")
async def prime(body: PrimeRequest):
    """
    Create a new thread and post context without running the agent.
    The Word add-in calls this first to seed the document content.
    """
    if not body.content:
        raise HTTPException(status_code=400, detail="content is required")
    try:
        thread_id = get_agent().prime(body.content)
        return {"conversationId": thread_id}
    except Exception as e:
        logger.exception("[prime] error")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/chat")
async def chat(body: ChatRequest):
    """
    Send a user message and run the agent.
    Returns either a text reply or a list of tool call requests for the add-in to execute.
    """
    if not body.prompt:
        raise HTTPException(status_code=400, detail="prompt is required")
    try:
        result = await get_agent().chat(
            prompt=body.prompt,
            conversation_id=body.conversationId,
        )
        return result
    except RuntimeError as e:
        status = 503 if "not ready" in str(e).lower() else 500
        raise HTTPException(status_code=status, detail=str(e))
    except Exception as e:
        logger.exception("[chat] error")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/tool-result")
async def tool_result(body: ToolResultRequest):
    """
    Submit tool outputs back to a waiting run after the Word add-in
    has executed the requested tool calls.
    """
    try:
        result = await get_agent().submit_tool_outputs(
            conversation_id=body.conversationId,
            tool_results=[tr.model_dump() for tr in body.toolResults],
        )
        return result
    except LookupError as e:
        raise HTTPException(status_code=409, detail=str(e))
    except Exception as e:
        logger.exception("[tool-result] error")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/cluster")
async def cluster(body: ClusterRequest):
    """
    Group a section's tracked changes into semantically related clusters.

    Calls the model directly as a plain chat completion — no agent thread,
    no tool loop, no run polling. Fast and stateless.

    Returns:
      {
        "clusters": [
          { "indices": [0, 2], "snippet": "...1-3 sentences containing these changes..." },
          { "indices": [1],    "snippet": "...sentence containing change 1..." }
        ]
      }

    Each cluster's snippet is chosen by the model — only the relevant sentences
    from paragraphContext, not the full paragraph.
    """
    if not body.changes:
        return {"clusters": []}
    # Single change — skip the model call, just extract a snippet directly
    if len(body.changes) == 1:
        c = body.changes[0]
        snippet = _extract_snippet_fallback(c.paragraphContext, [c.model_dump()])
        return {"clusters": [{"indices": [0], "snippet": snippet}]}
    try:
        clusters = await get_agent().cluster_changes(
            section_number=body.sectionNumber,
            section_title=body.sectionTitle,
            changes=[c.model_dump() for c in body.changes],
        )
        return {"clusters": clusters}
    except Exception as e:
        logger.exception("[cluster] error")
        raise HTTPException(status_code=500, detail=str(e))


@app.delete("/api/session")
async def delete_session(body: SessionDeleteRequest):
    """
    End a session: delete the thread, tear down the current agent, spin up a fresh one.
    """
    agent = get_agent()
    if body.conversationId:
        agent.delete_thread(body.conversationId)
    try:
        await agent.recreate()
        return {"ok": True, "message": "Session ended. New agent ready."}
    except Exception as e:
        logger.exception("[session] error recreating agent")
        raise HTTPException(status_code=500, detail=str(e))


# ─── SSL cert helpers ──────────────────────────────────────────────────────────

def _find_certs() -> tuple[str, str] | None:
    home = os.path.expanduser("~")
    cert_dir = os.path.join(home, ".office-addin-dev-certs")
    candidates = [
        ("localhost.crt", "localhost.key"),
        ("ca.crt",        "localhost.key"),
        ("server.crt",    "server.key"),
    ]
    for cert_name, key_name in candidates:
        cert = os.path.join(cert_dir, cert_name)
        key  = os.path.join(cert_dir, key_name)
        if os.path.exists(cert) and os.path.exists(key):
            logger.info(f"Found dev certs: {cert}")
            return cert, key
    return None


def _get_dev_cert_paths() -> tuple[str, str]:
    found = _find_certs()
    if found:
        return found

    logger.info("Dev certs not found — running: npx office-addin-dev-certs install")
    try:
        subprocess.run(
            ["npx", "office-addin-dev-certs", "install", "--days", "365"],
            check=True,
        )
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        raise RuntimeError(
            "Could not install dev certs. Make sure Node.js is installed, then run:\n"
            "  npx office-addin-dev-certs install\n"
            "Or set SSL_CERT and SSL_KEY env vars to point to your own cert/key files."
        ) from e

    found = _find_certs()
    if found:
        return found

    cert_dir = os.path.join(os.path.expanduser("~"), ".office-addin-dev-certs")
    files = os.listdir(cert_dir) if os.path.isdir(cert_dir) else ["(directory missing)"]
    raise FileNotFoundError(
        f"Certs installed but couldn't find expected files in {cert_dir}\n"
        f"Files present: {files}\n"
        f"Set SSL_CERT=<path> and SSL_KEY=<path> env vars to point to the correct files."
    )


# ─── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn

    PORT     = int(os.getenv("PORT", "3001"))
    ssl_cert = os.getenv("SSL_CERT")
    ssl_key  = os.getenv("SSL_KEY")

    if ssl_cert and ssl_key:
        logger.info("Using SSL certs from env vars")
    else:
        ssl_cert, ssl_key = _get_dev_cert_paths()

    logger.info(f"Starting LexAI on https://localhost:{PORT}")
    uvicorn.run(
        "main:app",
        host="localhost",
        port=PORT,
        ssl_certfile=ssl_cert,
        ssl_keyfile=ssl_key,
        reload=True,
    )