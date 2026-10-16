"""Read the broker / borrower email chain: .eml / .msg files in the deal folder (+ Outlook via graph_mail.py).

    parse_eml(path) / parse_msg(path) / parse_file(path)  -> EmailMsg
    clean_body(text, subject)    strip quoted history, caution banners, "Sent from my iPhone", disclaimers
                                 (a forward keeps the forwarded content - deal details often live there)
    build_chain(local, graph)    dedupe on Message-ID, order by date, clean, cap at ~60k characters
                                 (keeps the original first message and the newest ones)
    chain_text(chain)            "date | from | subject" + cleaned body per message, for the extractor
"""
from __future__ import annotations

import email
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from email import policy
from email.utils import getaddresses, parsedate_to_datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable, Optional

Log = Callable[[str], None]
MAX_CHAIN_CHARS = 60_000
MAX_BODY_CHARS = 15_000          # one very long message (e.g. a pasted rent roll) cannot crowd out the rest


@dataclass
class EmailMsg:
    message_id: Optional[str]
    in_reply_to: Optional[str]
    references: list[str]
    thread_topic: Optional[str]
    subject: str
    sender: str
    to: list[str]
    cc: list[str]
    date: Optional[datetime]         # timezone-aware
    body_text: str                   # as received (quoted history included)
    attachment_names: list[str]
    source: str = "local .eml"       # local .eml | local .msg | Graph (<mailbox>) | ... + Graph
    origin: str = ""                 # file path inside the deal folder, or the mailbox
    clean_text: str = ""             # filled by build_chain
    included: bool = True            # False when dropped by the length cap

    @property
    def key(self) -> str:
        if self.message_id:
            return self.message_id.strip().strip("<>").lower()
        when = self.date.strftime("%Y%m%d%H%M") if self.date else ""
        return f"{normalise_subject(self.subject).lower()}|{_addr(self.sender)}|{when}"

    def header_line(self) -> str:
        when = self.date.strftime("%Y-%m-%d %H:%M %Z").strip() if self.date else "date unknown"
        return f"{when} | {self.sender} | {self.subject}"

    def meta(self) -> dict:
        """Chain metadata for facts.json (no bodies)."""
        return {"date": self.date.isoformat() if self.date else None, "sender": self.sender, "to": self.to,
                "cc": self.cc, "subject": self.subject, "source": self.source, "origin": self.origin,
                "message_id": self.message_id, "attachments": self.attachment_names,
                "included_in_prompt": self.included, "clean_chars": len(self.clean_text)}


# --------------------------------------------------------------------------- helpers
def _addr(s: str) -> str:
    m = re.search(r"[\w.+-]+@[\w-]+(\.[\w-]+)+", s or "")
    return m.group(0).lower() if m else (s or "").strip().lower()


def _ids(value) -> list[str]:
    return re.findall(r"<[^<>\s]+>", str(value or ""))


def _aware(d: Optional[datetime]) -> Optional[datetime]:
    if d is None:
        return None
    return d if d.tzinfo else d.replace(tzinfo=timezone.utc)


_PREFIX = re.compile(r"^\s*((re|fw|fwd|aw|wg|tr)\s*(\[\d+\])?\s*:\s*|\[external\]\s*|external\s*:\s*)+", re.I)


def normalise_subject(subject: Optional[str]) -> str:
    """'RE: FW: [EXTERNAL] 150 E Aurora | Bridge Loan Request' -> '150 E Aurora | Bridge Loan Request'"""
    return _PREFIX.sub("", subject or "").strip()


def is_forward(subject: Optional[str]) -> bool:
    m = re.match(r"^\s*(?:\[external\]\s*)?(fw|fwd|wg|tr)\s*:", subject or "", re.I)
    return bool(m)


class _HTMLText(HTMLParser):
    BLOCK = {"p", "div", "br", "tr", "li", "h1", "h2", "h3", "h4", "h5", "h6", "table", "blockquote", "hr", "section"}
    SKIP = {"script", "style", "head", "title"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in self.SKIP:
            self.skip += 1
        elif tag in self.BLOCK:
            self.parts.append("\n")
        elif tag in ("td", "th"):
            self.parts.append(" ")

    def handle_endtag(self, tag):
        if tag in self.SKIP:
            self.skip = max(0, self.skip - 1)
        elif tag in self.BLOCK:
            self.parts.append("\n")

    def handle_data(self, data):
        if not self.skip:
            self.parts.append(data)


def html_to_text(html: str) -> str:
    p = _HTMLText()
    try:
        p.feed(html or "")
        p.close()
    except Exception:
        return re.sub(r"<[^>]+>", " ", html or "")
    text = "".join(p.parts).replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


# --------------------------------------------------------------------------- parsing
def _body_from_message(msg) -> str:
    part = msg.get_body(preferencelist=("plain", "html"))
    if part is None:
        return ""
    try:
        content = part.get_content()
    except Exception:
        content = (part.get_payload(decode=True) or b"").decode("utf-8", "replace")
    return html_to_text(content) if part.get_content_type() == "text/html" else content


def parse_eml(path: str | Path) -> EmailMsg:
    with open(path, "rb") as fh:
        msg = email.message_from_binary_file(fh, policy=policy.default)
    body = _body_from_message(msg)
    attachments = []
    for part in msg.iter_attachments():
        name = part.get_filename()
        if part.get_content_type() == "message/rfc822":
            inner = part.get_content()
            attachments.append(name or f"{inner.get('subject', 'attached message')}.eml")
            body += (f"\n\n---------- Forwarded message ----------\nFrom: {inner.get('from', '')}\n"
                     f"Date: {inner.get('date', '')}\nSubject: {inner.get('subject', '')}\n\n{_body_from_message(inner)}")
        elif name:
            attachments.append(name)
    date = None
    try:
        date = msg["date"].datetime if msg["date"] is not None else None
    except Exception:
        try:
            date = parsedate_to_datetime(str(msg.get("date")))
        except Exception:
            date = None
    addrs = lambda h: [f"{n} <{a}>" if n else a for n, a in getaddresses([str(v) for v in msg.get_all(h, [])])]  # noqa: E731
    return EmailMsg(
        message_id=(str(msg.get("message-id") or "").strip() or None),
        in_reply_to=(_ids(msg.get("in-reply-to")) or [None])[0],
        references=_ids(msg.get("references")),
        thread_topic=str(msg.get("thread-topic")) if msg.get("thread-topic") else None,
        subject=str(msg.get("subject") or ""),
        sender=", ".join(addrs("from")),
        to=addrs("to"), cc=addrs("cc"),
        date=_aware(date), body_text=body, attachment_names=attachments,
        source="local .eml", origin=str(path))


def parse_msg(path: str | Path) -> EmailMsg:
    """Outlook .msg file (needs the extract-msg package)."""
    import extract_msg

    m = extract_msg.openMsg(str(path))
    try:
        header = getattr(m, "header", None)
        hget = (lambda k: header.get(k) if header is not None else None)
        body = m.body or ""
        if not body.strip() and getattr(m, "htmlBody", None):
            raw = m.htmlBody
            body = html_to_text(raw.decode("utf-8", "replace") if isinstance(raw, bytes) else str(raw))
        split = lambda s: [x.strip() for x in re.split(r"[;,]\s*(?![^<]*>)", s or "") if x.strip()]  # noqa: E731
        date = m.date if isinstance(getattr(m, "date", None), datetime) else None
        if date is None and hget("Date"):
            try:
                date = parsedate_to_datetime(str(hget("Date")))
            except Exception:
                date = None
        names = []
        for a in getattr(m, "attachments", []) or []:
            n = getattr(a, "longFilename", None) or getattr(a, "shortFilename", None) or getattr(a, "displayName", None)
            if n:
                names.append(str(n))
        return EmailMsg(
            message_id=(str(m.messageId or hget("Message-ID") or "").strip() or None),
            in_reply_to=(_ids(getattr(m, "inReplyTo", None) or hget("In-Reply-To")) or [None])[0],
            references=_ids(hget("References")),
            thread_topic=str(hget("Thread-Topic")) if hget("Thread-Topic") else None,
            subject=str(m.subject or ""), sender=str(m.sender or ""), to=split(m.to), cc=split(m.cc),
            date=_aware(date), body_text=body, attachment_names=names, source="local .msg", origin=str(path))
    finally:
        try:
            m.close()
        except Exception:
            pass


def parse_file(path: str | Path) -> EmailMsg:
    return parse_msg(path) if str(path).lower().endswith(".msg") else parse_eml(path)


# --------------------------------------------------------------------------- cleaning
# Where the quoted history of a reply starts.
_HISTORY = [
    re.compile(r"^-{2,}\s*Original Message\s*-{2,}", re.I | re.M),
    re.compile(r"^On\b[^\n]{0,300}?(\n[^\n]{0,300}?)?\bwrote:\s*$", re.M),
    re.compile(r"^\s*\**From:\**\s.+\n(.*\n){0,3}?\s*\**(Sent|Date):\**\s", re.I | re.M),
]
# Where a forwarded message starts (kept for a forward).
_FORWARD = [
    re.compile(r"^-{2,}\s*Forwarded message\s*-{2,}", re.I | re.M),
    re.compile(r"^Begin forwarded message:", re.I | re.M),
]
_BANNER = re.compile(
    r"(CAUTION:|\[?EXTERNAL\]?\s*(EMAIL|SENDER|MESSAGE)?\s*[:\-]|External Email|This (e-?mail|message) "
    r"(originated|was sent|came) from outside|Do not click links or open attachments|"
    r"originated outside (of )?(the|your|our) organi[sz]ation)", re.I)
_DISCLAIMER = re.compile(
    r"(confidential|privileged|intended (solely |only )?for the (use|addressee)|if you (have )?received this "
    r"(e-?mail|message|communication|transmission) in error|unauthori[sz]ed (use|review|disclosure|dissemination)|"
    r"disclaimer|not an offer|wire fraud|notify the sender immediately)", re.I)
_SIGNOFF = re.compile(r"^\s*(Sent from my (iPhone|iPad|Android|Galaxy|mobile device|BlackBerry).*|Get Outlook for (iOS|Android).*|"
                      r"Sent from (Mail|Outlook) for .*|Sent via .*mobile.*)\s*$", re.I | re.M)


def _drop_paragraphs(text: str) -> str:
    keep = []
    for para in re.split(r"\n\s*\n", text):
        p = para.strip()
        if not p:
            continue
        if _BANNER.search(p) and len(p) < 700:
            continue
        if len(p) > 200 and len(_DISCLAIMER.findall(p)) >= 2:
            continue
        keep.append(para.rstrip())
    return "\n\n".join(keep)


def clean_body(text: str, subject: Optional[str] = None) -> str:
    """The new text of one message: no quoted history, banners, mobile sign-offs or long legal disclaimers.

    A forward (subject FW:/Fwd:, or a forward marker before any reply marker) keeps the forwarded message,
    because the broker's deal details are often only inside it.
    """
    t = (text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\xa0", " ")
    hist = min((m.start() for rx in _HISTORY for m in [rx.search(t)] if m), default=None)
    fwd = min((m.start() for rx in _FORWARD for m in [rx.search(t)] if m), default=None)
    forward = is_forward(subject) or (fwd is not None and (hist is None or fwd <= hist))
    if forward:
        t = re.sub(r"^>+ ?", "", t, flags=re.M)          # forwarded plain text is sometimes '>'-quoted
    else:
        t = "\n".join(line for line in t.split("\n") if not line.lstrip().startswith(">"))
        cut = min((m.start() for rx in _HISTORY for m in [rx.search(t)] if m), default=None)
        if cut is not None:
            t = t[:cut]
    t = _SIGNOFF.sub("", t)
    t = re.sub(r"^\s*_{8,}\s*$", "", t, flags=re.M)      # Outlook separator line
    t = _drop_paragraphs(t)
    t = re.sub(r"[ \t]+\n", "\n", t)
    return re.sub(r"\n{3,}", "\n\n", t).strip()


# --------------------------------------------------------------------------- chain
@dataclass
class EmailChain:
    messages: list[EmailMsg] = field(default_factory=list)   # chronological, deduped
    local_count: int = 0
    graph_count: int = 0
    duplicates: int = 0
    omitted: int = 0                                          # dropped by the length cap
    errors: list[str] = field(default_factory=list)

    @property
    def included(self) -> list[EmailMsg]:
        return [m for m in self.messages if m.included]

    def summary(self) -> str:
        return (f"Email chain: {len(self.messages)} message(s) (local files: {self.local_count}, "
                f"Graph: {self.graph_count}, duplicates removed: {self.duplicates}"
                + (f", {self.omitted} older message(s) left out of the prompt to stay under the length cap" if self.omitted else "")
                + ")")


def build_chain(local: list[EmailMsg], graph: list[EmailMsg] | None = None,
                max_chars: int = MAX_CHAIN_CHARS) -> EmailChain:
    graph = graph or []
    chain = EmailChain(local_count=len(local), graph_count=len(graph))
    by_key: dict[str, EmailMsg] = {}
    for m in list(local) + list(graph):
        k = m.key
        if k in by_key:
            chain.duplicates += 1
            old = by_key[k]
            if m.source not in old.source:
                old.source = f"{old.source} + {m.source.split(' (')[0]}"
            if not old.body_text.strip() and m.body_text.strip():
                old.body_text = m.body_text
            old.date = old.date or m.date
            continue
        by_key[k] = m
    far_future = datetime.max.replace(tzinfo=timezone.utc)
    msgs = sorted(by_key.values(), key=lambda m: m.date or far_future)
    for m in msgs:
        c = clean_body(m.body_text, m.subject)
        m.clean_text = c if len(c) <= MAX_BODY_CHARS else c[:MAX_BODY_CHARS] + "\n[... message truncated ...]"
        m.included = True
    # length cap: always keep the first (original) message, then the newest ones that still fit
    size = lambda m: len(m.header_line()) + len(m.clean_text) + 4  # noqa: E731
    if msgs and sum(size(m) for m in msgs) > max_chars:
        budget = max_chars - size(msgs[0])
        for m in reversed(msgs[1:]):
            if size(m) <= budget:
                budget -= size(m)
            else:
                m.included = False
        # once one message is dropped, drop everything older too (no gaps in the recent part of the thread)
        dropped = False
        for m in reversed(msgs[1:]):
            dropped = dropped or not m.included
            m.included = not dropped
        chain.omitted = sum(not m.included for m in msgs)
    chain.messages = msgs
    return chain


def chain_text(chain: EmailChain) -> str:
    out = []
    for i, m in enumerate(chain.included):
        if i == 1 and chain.omitted:
            out.append(f"[... {chain.omitted} message(s) in the middle of the thread omitted ...]")
        att = f"\nAttachments: {', '.join(m.attachment_names)}" if m.attachment_names else ""
        out.append(f"{m.header_line()}{att}\n{m.clean_text or '(no new text)'}")
    return "\n\n".join(out)


# --------------------------------------------------------------------------- collect (local + Graph)
def collect(storage, selection, settings, log: Log = print, conversation_id: str = "",
            use_graph: bool = False, graph=None) -> EmailChain:
    """Parse the located .eml/.msg files, optionally add the Outlook thread via Graph, and build the chain.

    Never raises: unreadable files and Graph problems are logged and recorded in chain.errors.
    """
    local, errors = [], []
    for e in (selection.files if selection is not None and selection.found else []):
        try:
            m = parse_file(storage.fetch(e))
            m.origin = e.rel
            local.append(m)
        except Exception as exc:
            errors.append(f"Could not read email {e.rel}: {exc}")
            log(f"  could not read email {e.rel}: {exc}")
    graph_msgs: list[EmailMsg] = []
    if (use_graph or conversation_id) and getattr(settings, "has_graph", False):
        try:
            from .graph_mail import GraphMail
            g = graph or GraphMail(settings, log)
            graph_msgs = g.fetch_chain(local, conversation_id=conversation_id)
        except Exception as exc:
            errors.append(f"Outlook (Graph) not read: {exc}")
            log(f"  Outlook (Graph) not read: {exc} - continuing with the local email files only")
    elif conversation_id:
        log("  Conversation ID given but Microsoft Graph is not configured in .env - using local email files only")
    chain = build_chain(local, graph_msgs, int(settings.get("emails.max_chain_chars", MAX_CHAIN_CHARS)))
    chain.errors = errors
    if chain.messages:
        log("  " + chain.summary())
    return chain
