"""One-time helper: turn your Dropbox App key + secret into a long-lived refresh token.

    python tools/dropbox_get_refresh_token.py

1. Opens an authorisation URL - sign in with the Castellan Dropbox account and click Allow.
2. Paste the code shown by Dropbox back here.
3. Copy the printed DROPBOX_REFRESH_TOKEN into your .env file.

The Dropbox app needs these permissions (App Console -> Permissions tab), then re-run this tool:
    files.metadata.read, files.content.read, account_info.read
    (team spaces / team folders may also need: team_data.member if it is a team-scoped app)
"""
import os
import sys
import webbrowser
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from dotenv import load_dotenv  # noqa: E402
from dropbox import DropboxOAuth2FlowNoRedirect  # noqa: E402

load_dotenv(Path(__file__).resolve().parent.parent / ".env")
key = os.getenv("DROPBOX_APP_KEY") or input("Dropbox App key: ").strip()
secret = os.getenv("DROPBOX_APP_SECRET") or input("Dropbox App secret: ").strip()

flow = DropboxOAuth2FlowNoRedirect(key, secret, token_access_type="offline")
url = flow.start()
print("\n1) Open this URL and click 'Allow':\n   " + url)
try:
    webbrowser.open(url)
except Exception:
    pass
code = input("\n2) Paste the authorisation code here: ").strip()
res = flow.finish(code)
print("\n3) Add these lines to your .env file:\n")
print(f"DROPBOX_APP_KEY={key}")
print(f"DROPBOX_APP_SECRET={secret}")
print(f"DROPBOX_REFRESH_TOKEN={res.refresh_token}")
