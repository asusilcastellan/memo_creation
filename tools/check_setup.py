"""Quick connectivity check: OpenAI key + models, Dropbox access, and an optional folder listing.

    python tools/check_setup.py
    python tools/check_setup.py "/Bridge Lending/Current Portfolio/150 E Aurora St"
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from memo_app.settings import load_settings  # noqa: E402

s = load_settings()
print("OpenAI key:", "set" if s.openai_api_key else "MISSING")
if s.openai_api_key:
    from openai import OpenAI
    try:
        ids = {m.id for m in OpenAI(api_key=s.openai_api_key).models.list()}
        for role, m in (("extract", s.extract_model), ("write", s.write_model)):
            print(f"  {role} model {m}: {'available' if m in ids else 'NOT FOUND - pick one from the list below'}")
        if s.extract_model not in ids or s.write_model not in ids:
            print("  available gpt models:", ", ".join(sorted(i for i in ids if i.startswith("gpt"))))
    except Exception as exc:
        print("  OpenAI error:", exc)

print("Dropbox:", "configured" if s.has_dropbox else "not configured")
if s.has_dropbox:
    from memo_app.storage import DropboxStorage
    try:
        st = DropboxStorage(s)
        acct = st.dbx.users_get_current_account()
        print("  signed in as", acct.name.display_name, "<" + acct.email + ">")
        if len(sys.argv) > 1:
            entries = st.walk(sys.argv[1])
            print(f"  {sys.argv[1]}: {len(entries)} items")
            for e in sorted(entries, key=lambda e: e.rel)[:60]:
                print("   ", "[D]" if e.is_dir else "   ", e.rel)
    except Exception as exc:
        print("  Dropbox error:", exc)
