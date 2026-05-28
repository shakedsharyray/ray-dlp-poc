"""Ray DLP — classifier service (PoC).

POST /classify
  Body: ClassifyRequest (see model below)
  Returns: {"action": "block"|"allow", "reason": str}

GET /health
  Returns: {"status": "ok"}
"""

import base64
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

app = FastAPI(title="Ray DLP Classifier", version="0.2.0")

# PoC-permissive CORS. Tighten to specific Outlook origins for production.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "OPTIONS"],
    allow_headers=["*"],
)

BLOCK_TERMS = ["bamba", "bisli"]

# Content types we'll try to decode as text. Anything not in this list, we
# only inspect the filename (not the content).
TEXT_CONTENT_TYPES = (
    "text/",
    "application/json",
    "application/xml",
    "application/javascript",
    "application/x-yaml",
)


class Recipient(BaseModel):
    emailAddress: Optional[str] = None
    displayName: Optional[str] = None


class Attachment(BaseModel):
    id: Optional[str] = None
    name: Optional[str] = None
    contentType: Optional[str] = None
    size: Optional[int] = None
    isInline: Optional[bool] = False
    attachmentType: Optional[str] = None
    content: Optional[str] = None  # base64 string, or None if skipped
    format: Optional[str] = None   # "base64" or "url"


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


def _looks_textual(content_type: Optional[str]) -> bool:
    if not content_type:
        return False
    ct = content_type.lower()
    return any(ct.startswith(prefix) for prefix in TEXT_CONTENT_TYPES)


def _attachment_text(att: Attachment) -> Optional[str]:
    """Decode an attachment's base64 content to text if it looks textual."""
    if not att.content or att.format != "base64":
        return None
    if not _looks_textual(att.contentType):
        # Don't try to decode binary content (.pdf, .docx, images, etc.)
        # — those need a real parser. For PoC: textual only.
        return None
    try:
        raw = base64.b64decode(att.content)
        return raw.decode("utf-8", errors="ignore")
    except Exception as e:
        log.warning("attachment decode failed name=%r err=%s", att.name, e)
        return None


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/classify", response_model=ClassifyResponse)
def classify(msg: ClassifyRequest) -> ClassifyResponse:
    log.info(
        "classify sender=%s subject=%r body_chars=%d to=%d cc=%d attachments=%d",
        msg.sender,
        (msg.subject or "")[:80],
        len(msg.body or ""),
        len(msg.to),
        len(msg.cc),
        len(msg.attachments),
    )
    for a in msg.attachments:
        log.info(
            "  attachment name=%r size=%s contentType=%s isInline=%s has_content=%s",
            a.name, a.size, a.contentType, a.isInline, bool(a.content),
        )

    # 1. Subject + body
    haystack = ((msg.subject or "") + " " + (msg.body or "")).lower()
    matched = next((t for t in BLOCK_TERMS if t in haystack), None)
    if matched:
        return ClassifyResponse(
            action="block",
            reason=(
                f"Blocked by Ray DLP: this message contains the term '{matched}'. "
                "Please remove it and try again."
            ),
        )

    # 2. Attachment filenames (always cheap to check)
    for a in msg.attachments:
        if not a.name:
            continue
        name_lower = a.name.lower()
        matched = next((t for t in BLOCK_TERMS if t in name_lower), None)
        if matched:
            return ClassifyResponse(
                action="block",
                reason=(
                    f"Blocked by Ray DLP: attachment '{a.name}' filename contains "
                    f"the term '{matched}'."
                ),
            )

    # 3. Attachment text content (for textual MIME types only)
    for a in msg.attachments:
        text = _attachment_text(a)
        if not text:
            continue
        text_lower = text.lower()
        matched = next((t for t in BLOCK_TERMS if t in text_lower), None)
        if matched:
            return ClassifyResponse(
                action="block",
                reason=(
                    f"Blocked by Ray DLP: attachment '{a.name}' contains "
                    f"the term '{matched}'."
                ),
            )

    return ClassifyResponse(action="allow", reason="No policy violations detected.")
