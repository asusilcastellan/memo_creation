"""Check the Microsoft Graph (Outlook) setup without generating a memo.

    python tools/check_graph.py                          # sign in, list the 5 newest messages per mailbox
    python tools/check_graph.py --subject "150 E Aurora"  # find threads by subject (shows their Conversation IDs)
    python tools/check_graph.py --conversation-id "AAQk..."   # read one thread as the memo would see it
    python tools/check_graph.py --eml "D:\\deal\\Emails\\request.eml"   # find the thread behind a saved email
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from memo_app import emails as E  # noqa: E402
from memo_app.graph_mail import GRAPH, GraphMail  # noqa: E402
from memo_app.settings import load_settings  # noqa: E402

ap = argparse.ArgumentParser()
ap.add_argument("--subject", default="")
ap.add_argument("--conversation-id", default="")
ap.add_argument("--eml", default="", help="a saved .eml/.msg file; its Message-ID is looked up in each mailbox")
a = ap.parse_args()

s = load_settings()
for k in ("graph_tenant_id", "graph_client_id", "graph_client_secret"):
    print(f"{k.upper()}: {'set' if getattr(s, k) else 'MISSING'}")
print("GRAPH_MAILBOXES:", ", ".join(s.graph_mailboxes) or "MISSING")
if not s.has_graph:
    sys.exit("Graph is not configured in .env - see .env.example")

g = GraphMail(s)
try:
    g.token()
    print("Sign-in: OK (app-only token received)")
except Exception as exc:
    sys.exit(f"Sign-in FAILED: {exc}")

for mbx in s.graph_mailboxes:
    print(f"\n=== {mbx}")
    try:
        if a.conversation_id or a.eml:
            if a.eml:
                local = [E.parse_file(a.eml)]
                print(f"  .eml Message-ID: {local[0].message_id}")
                msgs = g.fetch_chain(local)
            else:
                msgs = g.conversation(mbx, a.conversation_id)
            chain = E.build_chain(msgs)
            print(f"  {chain.summary()}")
            for m in chain.messages:
                print(f"\n  {m.header_line()}\n    " + (m.clean_text[:300] or "(no new text)").replace("\n", "\n    "))
        elif a.subject:
            data = g._get(f"{GRAPH}/users/{mbx}/messages", {"$search": f'"{a.subject}"', "$top": "10",
                          "$select": "subject,receivedDateTime,conversationId"}, headers={"ConsistencyLevel": "eventual"})
            for m in data.get("value", []):
                print(f"  {m.get('receivedDateTime')} | {m.get('subject')}\n      Conversation ID: {m.get('conversationId')}")
            if not data.get("value"):
                print("  no messages found")
        else:
            data = g._get(f"{GRAPH}/users/{mbx}/messages", {"$top": "5", "$select": "subject,receivedDateTime,conversationId"})
            for m in data.get("value", []):
                print(f"  {m.get('receivedDateTime')} | {m.get('subject')}")
            print("  mailbox readable: OK")
    except Exception as exc:
        print(f"  FAILED: {exc}")
