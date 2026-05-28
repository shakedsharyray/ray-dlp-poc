"""Ray DLP — classifier service (PoC).

POST /classify
  Body: ClassifyRequest (see model below)
  Returns: {"action": "block"|"allow", "reason": str}

GET /health
  Returns: {"status": "ok"}
"""

import logging
from typing import List, Literal, Optional

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("ray-dlp")

app = FastAPI(title="Ray DLP Classifier", version="0.1.0")

# PoC-permissive CORS. Tighten to specific Outlook origins for production.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "OPTIONS"],
    allow_headers=["*"],
)

BLOCK_TERMS = ["bamba", "bisli"]


class Recipient(BaseModel):
    emailAddress: Optional[str] = None
    displayName: Optional[str] = None


class Attachment(BaseModel):
    name: Optional[str] = None
    contentType: Optional[str] = None
    size: Optional[int] = None


class ClassifyRequest(BaseModel):
    sender: Optional[str] = None
    subject: Optional[str] = ""
    body: Optional[str] = ""
    to: List[Recipient] = []
    cc: List[Recipient] = []
    bcc: List[Recipient] = []
    attachments: List[Attachment] = []
    itemId: Optional[str] = None
    timestamp: Optional[str] = None


class ClassifyResponse(BaseModel):
    action: Literal["block", "allow"]
    reason: str


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/classify", response_model=ClassifyResponse)
def classify(msg: ClassifyRequest) -> ClassifyResponse:
    log.info(
        "classify sender=%s subject=%r body_chars=%d to=%d cc=%d bcc=%d attachments=%d",
        msg.sender,
        (msg.subject or "")[:80],
        len(msg.body or ""),
        len(msg.to),
        len(msg.cc),
        len(msg.bcc),
        len(msg.attachments),
    )
    for a in msg.attachments:
        log.info("  attachment name=%r size=%s contentType=%s",
                 a.name, a.size, a.contentType)

    haystack = ((msg.subject or "") + " " + (msg.body or "")).lower()
    matched = next((term for term in BLOCK_TERMS if term in haystack), None)
    if matched:
        return ClassifyResponse(
            action="block",
            reason=(
                f"Blocked by Ray DLP: this message contains the term '{matched}'. "
                "Please remove it and try again."
            ),
        )
    return ClassifyResponse(action="allow", reason="No policy violations detected.")
