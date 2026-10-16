"""Read the Outlook thread behind the deal's emails through Microsoft Graph (app-only, MSAL).

Setup (.env):  GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_MAILBOXES (comma-separated UPNs).
Azure app permission: Mail.Read (Application) with admin consent.

For each mailbox:
  a) find the message by internetMessageId (the Message-ID of a local .eml) -> conversationId
  b) read every message in that conversation (uniqueBody = the new text of a reply), following @odata.nextLink
  c) no Message-ID match -> $search on the normalised subject and keep results whose subject matches
A conversation ID typed in the app skips step a). Nothing here raises to the caller of fetch_chain():
problems are logged and the memo continues with the local files.
"""
from __future__ import annotations

import threading
import time
from datetime import datetime, timezone
from typing import Callable, Optional

from .emails import EmailMsg, html_to_text, is_forward, normalise_subject
from .net import httpx

Log = Callable[[str], None]
GRAPH = "https://graph.microsoft.com/v1.0"
SCOPE = ["https://graph.microsoft.com/.default"]
SELECT = ("subject,from,toRecipients,ccRecipients,receivedDateTime,body,uniqueBody,internetMessageId,"
          "hasAttachments,conversationId")
RETRY_STATUS = {429, 503, 504}

_apps: dict[tuple[str, str], object] = {}
_apps_lock = threading.Lock()


class GraphError(RuntimeError):
    pass


def _odata_str(v: str) -> str:
    return v.replace("'", "''")


def _person(r: dict | None) -> str:
    a = (r or {}).get("emailAddress") or {}
    name, addr = a.get("name") or "", a.get("address") or ""
    return f"{name} <{addr}>" if name and addr and name != addr else (addr or name)


def _when(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return None


class GraphMail:
    def __init__(self, settings, log: Log = print, client=None, token_provider: Callable[[], str] | None = None,
                 sleep: Callable[[float], None] = time.sleep, max_retries: int = 4):
        self.s = settings
        self.log = log
        self.client = client or httpx.Client(timeout=30.0)
        self._token_provider = token_provider
        self.sleep = sleep
        self.max_retries = max_retries
        self.mailboxes = [m.strip() for m in (getattr(settings, "graph_mailboxes", None) or []) if m.strip()]

    # ------------------------------------------------------------------ auth
    def token(self) -> str:
        if self._token_provider:
            return self._token_provider()
        import msal

        key = (self.s.graph_tenant_id, self.s.graph_client_id)
        with _apps_lock:
            app = _apps.get(key)
            if app is None:
                app = msal.ConfidentialClientApplication(
                    self.s.graph_client_id, authority=f"https://login.microsoftonline.com/{self.s.graph_tenant_id}",
                    client_credential=self.s.graph_client_secret)
                _apps[key] = app
        res = app.acquire_token_for_client(scopes=SCOPE)       # MSAL caches the token until it expires
        if "access_token" not in res:
            raise GraphError(f"Azure AD sign-in failed: {res.get('error')}: {res.get('error_description')}")
        return res["access_token"]

    # ------------------------------------------------------------------ HTTP
    def _get(self, url: str, params: dict | None = None, headers: dict | None = None) -> dict:
        h = {"Authorization": f"Bearer {self.token()}", "Accept": "application/json", **(headers or {})}
        for attempt in range(self.max_retries + 1):
            r = self.client.get(url, params=params, headers=h)
            if r.status_code in RETRY_STATUS and attempt < self.max_retries:
                try:
                    wait = float(r.headers.get("Retry-After", ""))
                except ValueError:
                    wait = 2.0 ** attempt
                self.log(f"  Graph busy ({r.status_code}); retrying in {wait:.0f}s")
                self.sleep(min(wait, 60.0))
                continue
            if r.status_code >= 400:
                try:
                    err = r.json().get("error", {})
                    msg = f"{err.get('code')}: {err.get('message')}"
                except Exception:
                    msg = r.text[:300]
                raise GraphError(f"Graph {r.status_code} {msg}")
            return r.json()
        raise GraphError("Graph kept answering 429/503")

    def _messages(self, mbx: str) -> str:
        return f"{GRAPH}/users/{mbx}/messages"

    # ------------------------------------------------------------------ lookups
    def find_conversation(self, mbx: str, internet_message_id: str) -> Optional[str]:
        mid = internet_message_id.strip()
        if not mid.startswith("<"):
            mid = f"<{mid}>"
        data = self._get(self._messages(mbx), {"$filter": f"internetMessageId eq '{_odata_str(mid)}'",
                                               "$select": "id,conversationId,subject"})
        vals = data.get("value", [])
        return vals[0].get("conversationId") if vals else None

    def search_subject(self, mbx: str, subject: str, limit: int = 3) -> list[str]:
        """conversationIds of messages whose normalised subject matches (fallback when no Message-ID matches)."""
        target = normalise_subject(subject)
        if not target:
            return []
        q = target.replace('"', " ").replace("\\", " ")
        data = self._get(self._messages(mbx), {"$search": f'"{q}"', "$select": "id,conversationId,subject", "$top": "25"},
                         headers={"ConsistencyLevel": "eventual"})
        out: list[str] = []
        for v in data.get("value", []):
            if normalise_subject(v.get("subject")).lower() == target.lower() and v.get("conversationId") not in out:
                out.append(v["conversationId"])
        return out[:limit]

    def conversation(self, mbx: str, cid: str) -> list[EmailMsg]:
        # $orderby together with this $filter can fail with "InefficientFilter" -> sort on our side.
        url, params = self._messages(mbx), {"$filter": f"conversationId eq '{_odata_str(cid)}'", "$select": SELECT,
                                            "$top": "50"}
        out: list[EmailMsg] = []
        while url:
            data = self._get(url, params, headers={"Prefer": 'outlook.body-content-type="text"'})
            out += [self._to_msg(v, mbx) for v in data.get("value", [])]
            url, params = data.get("@odata.nextLink"), None
        out.sort(key=lambda m: m.date or datetime.max.replace(tzinfo=timezone.utc))
        return out

    @staticmethod
    def _to_msg(v: dict, mbx: str) -> EmailMsg:
        subject = v.get("subject") or ""

        def text(b):
            if not b or not (b.get("content") or "").strip():
                return ""
            return html_to_text(b["content"]) if (b.get("contentType") or "").lower() == "html" else b["content"]

        # uniqueBody leaves out the forwarded part of a forward, so a forward keeps its full body
        body = text(v.get("body")) if is_forward(subject) else (text(v.get("uniqueBody")) or text(v.get("body")))
        return EmailMsg(
            message_id=v.get("internetMessageId"), in_reply_to=None, references=[], thread_topic=None,
            subject=subject, sender=_person(v.get("from")),
            to=[_person(r) for r in v.get("toRecipients") or []], cc=[_person(r) for r in v.get("ccRecipients") or []],
            date=_when(v.get("receivedDateTime")), body_text=body,
            attachment_names=["(has attachments)"] if v.get("hasAttachments") else [],
            source=f"Graph ({mbx})", origin=mbx)

    # ------------------------------------------------------------------ main entry
    def fetch_chain(self, local: list[EmailMsg], conversation_id: str = "") -> list[EmailMsg]:
        """Messages of the thread(s) behind the local emails, from every configured mailbox. Never raises."""
        if not self.mailboxes:
            self.log("  Graph: no GRAPH_MAILBOXES configured")
            return []
        try:
            self.token()
        except Exception as exc:
            self.log(f"  Graph not available: {exc}")
            return []
        # one seed per thread: the earliest local message of each normalised subject
        seeds: dict[str, EmailMsg] = {}
        for m in sorted(local, key=lambda m: m.date.timestamp() if m.date else float("inf")):
            seeds.setdefault(normalise_subject(m.subject).lower(), m)
        out: list[EmailMsg] = []
        for mbx in self.mailboxes:
            try:
                cids: list[str] = [conversation_id.strip()] if conversation_id.strip() else []
                if not cids:
                    for subj, m in seeds.items():
                        cid = self.find_conversation(mbx, m.message_id) if m.message_id else None
                        if not cid and m.subject:
                            found = self.search_subject(mbx, m.subject)
                            cid = found[0] if found else None
                            cids += found[1:]
                        if cid:
                            cids.append(cid)
                cids = list(dict.fromkeys(c for c in cids if c))
                n0 = len(out)
                for cid in cids:
                    out += self.conversation(mbx, cid)
                self.log(f"  Graph {mbx}: {len(out) - n0} message(s) in {len(cids)} conversation(s)")
            except Exception as exc:
                self.log(f"  Graph {mbx}: not read ({exc}) - continuing")
        return out
