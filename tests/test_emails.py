"""Email chain: .eml parsing/cleaning, locator fallback, Graph (mocked with httpx.MockTransport). All offline."""
from __future__ import annotations

import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from memo_app import emails as E  # noqa: E402
from memo_app.graph_mail import GraphMail  # noqa: E402
from memo_app.locator import locate_all  # noqa: E402
from memo_app.net import httpx  # noqa: E402
from memo_app.settings import Settings, load_yaml_resource  # noqa: E402
from memo_app.storage import LocalStorage  # noqa: E402
from test_offline import ORIGINAL_ID, REPLY_ID, _eml, make_deal, make_emails  # noqa: E402

GRAPH_SETTINGS = Settings(graph_tenant_id="t", graph_client_id="c", graph_client_secret="s",
                          graph_mailboxes=["deals@castellanre.com"], raw={})
SUBJECT = "150 E Aurora St | Bridge Loan Request"


def _locate(deal: Path):
    entries = LocalStorage().walk(str(deal))
    return locate_all(entries, load_yaml_resource("config", "sources.yaml"), deal_name=deal.name)


def _local_chain(deal: Path):
    sel = _locate(deal)["emails"]
    return sel, [E.parse_file(e.path) for e in sel.files]


# --------------------------------------------------------------------------- local .eml
def test_located_parsed_cleaned_and_ordered(tmp_path):
    deal = make_deal(tmp_path)
    sel, local = _local_chain(deal)
    assert sel.folder == "Borrower material X" and len(sel.files) == 2
    assert [m.message_id for m in local] == [REPLY_ID, ORIGINAL_ID]          # file order: newest file first
    reply = local[0]
    assert reply.in_reply_to == ORIGINAL_ID and reply.references == [ORIGINAL_ID]
    assert reply.sender == "Shaya Margareten <shaya@smhi.com>" and reply.to == ["Analyst <analyst@castellanre.com>"]
    assert reply.date.tzinfo is not None and reply.date.utcoffset() == timedelta(hours=-4)

    chain = E.build_chain(local)
    assert [m.message_id for m in chain.messages] == [ORIGINAL_ID, REPLY_ID]  # thread order: by Date header
    r = chain.messages[1].clean_text
    assert r == "The roof will be replaced in Q1 2026 and two tenants have signed LOIs."
    fwd = chain.messages[0].clean_text                                       # forward keeps the forwarded request
    assert "Please see below" in fwd and "$3,290,000 bridge loan (65% LTV)" in fwd and "iPhone" not in fwd
    text = E.chain_text(chain)
    assert text.splitlines()[0].startswith("2025-10-01 09:00") and "| Dan Broker <dan@brokerco.com> |" in text


def test_fallback_any_folder_and_shared_folder_alias(tmp_path):
    deal = tmp_path / "77 Oak Ave"
    (deal / "Emails").mkdir(parents=True)
    (deal / "Emails" / "readme.pdf").write_text("not an email")          # named folder exists but has no .eml
    _eml(deal / "Misc" / "Thread" / "loan request.eml", "Loan request", "a@b.com", "c@d.com",
         datetime(2025, 9, 1, tzinfo=timezone.utc), "Hello", "<x1@b.com>")
    sel = _locate(deal)["emails"]
    assert sel.folder == "." and [f.rel for f in sel.files] == ["Misc/Thread/loan request.eml"]
    assert "searched the whole deal folder" in sel.note

    make_emails(deal / "77 Oak Ave - Shared Folder")                       # "<deal> - Shared Folder" alias
    sel = _locate(deal)["emails"]
    assert sel.folder == "77 Oak Ave - Shared Folder" and len(sel.files) == 2


def test_clean_body_variants():
    outlook = ("Thanks, see my comments.\n\n________________________________\nFrom: Dan Broker <dan@brokerco.com>\n"
               "Sent: Wednesday, October 1, 2025 9:00 AM\nTo: Analyst\nSubject: RE: deal\n\nOld text")
    assert E.clean_body(outlook, "RE: deal") == "Thanks, see my comments."
    original = "Short note.\n\n-----Original Message-----\nFrom: x\nold"
    assert E.clean_body(original, "RE: x") == "Short note."
    # Outlook forward (subject FW:) keeps the forwarded message even though it starts with From:/Sent:
    fw = "FYI\n\nFrom: Sponsor <s@x.com>\nSent: Monday\nSubject: Request\n\nNeed $2MM for 18 months."
    assert "Need $2MM for 18 months." in E.clean_body(fw, "FW: Request")
    assert E.clean_body("[EXTERNAL] Email: be careful with links.\n\nReal text\n\nGet Outlook for iOS") == "Real text"


def test_html_to_text_and_html_only_eml(tmp_path):
    assert E.html_to_text("<p>Loan&nbsp;request</p><div>Price: <b>$3,000,000</b></div><style>p{}</style>") == \
        "Loan request\n\nPrice: $3,000,000"
    from email.message import EmailMessage
    m = EmailMessage()
    m["Subject"], m["From"], m["Message-ID"] = "html only", "a@b.com", "<h1@b.com>"
    m.set_content("<html><body><p>Hello <i>there</i></p></body></html>", subtype="html")
    p = tmp_path / "h.eml"
    p.write_bytes(bytes(m))
    msg = E.parse_eml(p)
    assert msg.body_text.strip() == "Hello there" and msg.date is None


def test_chain_cap_keeps_original_and_newest():
    base = datetime(2025, 1, 1, tzinfo=timezone.utc)
    msgs = [E.EmailMsg(f"<m{i}@x>", None, [], None, f"RE: s{i}", "a@x.com", [], [], base + timedelta(days=i),
                       ("word " * 400).strip(), []) for i in range(10)]
    chain = E.build_chain(msgs, max_chars=7000)
    kept = [m.message_id for m in chain.included]
    assert kept[0] == "<m0@x>" and kept[-1] == "<m9@x>" and chain.omitted == 10 - len(kept) > 0
    assert kept == ["<m0@x>"] + [f"<m{i}@x>" for i in range(10 - len(kept) + 1, 10)]   # no gaps among the newest
    assert "omitted" in E.chain_text(chain)


# --------------------------------------------------------------------------- Graph (mocked)
def _graph_msg(mid, when, subject, body, unique=None, sender=("Dan Broker", "dan@brokerco.com")):
    return {"internetMessageId": mid, "subject": subject, "receivedDateTime": when, "conversationId": "CID1",
            "from": {"emailAddress": {"name": sender[0], "address": sender[1]}},
            "toRecipients": [{"emailAddress": {"name": "Analyst", "address": "analyst@castellanre.com"}}],
            "ccRecipients": [], "hasAttachments": False,
            "body": {"contentType": "text", "content": body},
            "uniqueBody": {"contentType": "text", "content": unique} if unique is not None else None}


class FakeGraph:
    """Answers the three Graph calls; first call is throttled once (429 + Retry-After)."""

    def __init__(self, id_match=True):
        self.requests: list[httpx.Request] = []
        self.id_match = id_match
        self.throttled = False

    def __call__(self, req: httpx.Request) -> httpx.Response:
        self.requests.append(req)
        assert req.headers["Authorization"] == "Bearer tok"
        q = {k: v[0] for k, v in parse_qs(urlparse(str(req.url)).query).items()}
        if not self.throttled:
            self.throttled = True
            return httpx.Response(429, headers={"Retry-After": "0"}, json={"error": {"code": "TooManyRequests"}})
        f = q.get("$filter", "")
        if f.startswith("internetMessageId eq"):
            hit = self.id_match and f == f"internetMessageId eq '{ORIGINAL_ID}'"
            return httpx.Response(200, json={"value": [{"id": "1", "conversationId": "CID1", "subject": SUBJECT}] if hit else []})
        if "$search" in q:
            assert "$filter" not in q and req.headers.get("ConsistencyLevel") == "eventual"
            return httpx.Response(200, json={"value": [{"id": "9", "conversationId": "CID1", "subject": "RE: " + SUBJECT},
                                                       {"id": "8", "conversationId": "OTHER", "subject": "Other deal"}]})
        if f == "conversationId eq 'CID1'":
            assert req.headers.get("Prefer") == 'outlook.body-content-type="text"' and "$orderby" not in q
            if q.get("page") != "2":
                return httpx.Response(200, json={
                    "value": [_graph_msg("<g3@castellanre.com>", "2025-10-04T15:00:00Z", "RE: " + SUBJECT,
                                         "full body with history", unique="We can do $2,240,000 at 65% LTC.",
                                         sender=("Analyst", "analyst@castellanre.com")),
                              _graph_msg(ORIGINAL_ID, "2025-10-01T13:00:00Z", "Fwd: " + SUBJECT,
                                         "Please see below - strong sponsor.", unique="Please see below")],
                    "@odata.nextLink": str(req.url.copy_merge_params({"page": "2"}))})
            return httpx.Response(200, json={"value": [
                _graph_msg("<g2@brokerco.com>", "2025-10-02T12:00:00Z", "RE: " + SUBJECT, "x",
                           unique="Sponsor has 20 years of industrial experience.\n\nSent from my iPhone")]})
        return httpx.Response(404, json={"error": {"code": "NotFound", "message": str(req.url)}})


def _graph(fake) -> GraphMail:
    return GraphMail(GRAPH_SETTINGS, log=lambda m: None, client=httpx.Client(transport=httpx.MockTransport(fake)),
                     token_provider=lambda: "tok", sleep=lambda s: None)


def test_graph_merge_dedupe_and_order(tmp_path):
    deal = make_deal(tmp_path)
    sel, local = _local_chain(deal)
    fake = FakeGraph()
    graph_msgs = _graph(fake).fetch_chain(local)
    assert len(graph_msgs) == 3                                          # 2 pages followed
    chain = E.build_chain(local, graph_msgs)
    ids = [m.message_id for m in chain.messages]
    assert ids == [ORIGINAL_ID, "<g2@brokerco.com>", REPLY_ID, "<g3@castellanre.com>"]
    assert chain.duplicates == 1 and len(set(ids)) == 4
    assert chain.messages[0].source == "local .eml + Graph"
    assert chain.messages[1].source == "Graph (deals@castellanre.com)"
    assert chain.messages[1].clean_text == "Sponsor has 20 years of industrial experience."   # uniqueBody, cleaned
    assert chain.messages[3].clean_text == "We can do $2,240,000 at 65% LTC."
    assert "Graph: 3" in chain.summary() and "duplicates removed: 1" in chain.summary()
    assert sum(r.url.params.get("$filter", "").startswith("internetMessageId") for r in fake.requests) == 2  # 429 + retry


def test_graph_subject_search_fallback_and_conversation_id(tmp_path):
    deal = make_deal(tmp_path)
    _, local = _local_chain(deal)
    fake = FakeGraph(id_match=False)
    msgs = _graph(fake).fetch_chain(local)
    assert {m.message_id for m in msgs} == {ORIGINAL_ID, "<g2@brokerco.com>", "<g3@castellanre.com>"}
    assert any("$search" in r.url.params for r in fake.requests)

    fake = FakeGraph()
    msgs = _graph(fake).fetch_chain([], conversation_id="CID1")         # typed in the app: no Message-ID lookup
    assert len(msgs) == 3 and not any("internetMessageId" in r.url.params.get("$filter", "") for r in fake.requests)


def test_graph_failure_never_stops_the_memo(tmp_path):
    deal = make_deal(tmp_path)
    sel, _ = _local_chain(deal)

    def no_token():
        raise RuntimeError("AADSTS7000215: Invalid client secret provided.")

    logs: list[str] = []
    g = GraphMail(GRAPH_SETTINGS, log=logs.append, token_provider=no_token)
    chain = E.collect(LocalStorage(), sel, GRAPH_SETTINGS, logs.append, use_graph=True, graph=g)
    assert len(chain.messages) == 2 and chain.graph_count == 0
    assert any("AADSTS7000215" in m for m in logs)

    down = GraphMail(GRAPH_SETTINGS, log=logs.append, token_provider=lambda: "tok", sleep=lambda s: None,
                     client=httpx.Client(transport=httpx.MockTransport(lambda r: httpx.Response(503, json={}))))
    chain = E.collect(LocalStorage(), sel, GRAPH_SETTINGS, logs.append, use_graph=True, graph=down)
    assert len(chain.messages) == 2 and any("not read" in m for m in logs)
