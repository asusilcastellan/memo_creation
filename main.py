"""Castellan Memo Builder - entry point.

GUI (default):        python main.py
Command line:         python main.py --cli --path "/Bridge Lending/Current Portfolio/150 E Aurora St" --loan-type acquisition
Scan sources only:    python main.py --cli --scan-only --path "..." --loan-type acquisition
"""
from __future__ import annotations

import argparse
import sys


def cli(argv: list[str]) -> int:
    from memo_app.facts import UserInputs
    from memo_app.pipeline import generate, scan
    from memo_app.settings import load_settings

    ap = argparse.ArgumentParser(description="Generate a Castellan loan investment memo draft.")
    ap.add_argument("--cli", action="store_true")
    ap.add_argument("--path", required=True, help="Deal folder: Dropbox path (/Bridge Lending/...) or local folder")
    ap.add_argument("--loan-type", required=True, choices=["acquisition", "construction", "refinance", "note_on_note"])
    ap.add_argument("--property-type", default="")
    ap.add_argument("--exit", default="", help="Exit strategy, e.g. 'Permanent refinance'")
    ap.add_argument("--notes-file", help="Text file with pasted broker/borrower emails or analyst notes")
    ap.add_argument("--index", default="", help="Index name overriding the term sheet, e.g. '1M Term SOFR', "
                    "'3M Term SOFR', '30-day Avg SOFR', 'AMERIBOR', 'WSJ Prime', 'Fixed rate'")
    ap.add_argument("--index-rate", type=float, default=None, help="Current index value in percent, e.g. 4.31")
    ap.add_argument("--prior-relationship", action="store_true")
    ap.add_argument("--graph", action="store_true", help="Also read the email thread from Outlook (Microsoft Graph, see .env)")
    ap.add_argument("--conversation-id", default="", help="Outlook conversation ID to read through Graph (implies --graph)")
    ap.add_argument("--web-search", action="store_true", help="Run the online adverse-media search")
    ap.add_argument("--out", default="", help="Output folder")
    ap.add_argument("--scan-only", action="store_true", help="Only show which files would be used")
    a = ap.parse_args(argv)

    notes = open(a.notes_file, encoding="utf-8").read() if a.notes_file else ""
    ui = UserInputs(deal_path=a.path, loan_type=a.loan_type, property_type=a.property_type.lower(),
                    exit_strategy=a.exit, prior_relationship=a.prior_relationship, notes=notes,
                    index_name_override=a.index, index_rate_pct=a.index_rate,
                    run_web_search=a.web_search, output_dir=a.out,
                    read_graph=a.graph or bool(a.conversation_id), conversation_id=a.conversation_id)
    settings = load_settings()
    from memo_app import cache
    from memo_app.progress import Reporter, UserFacingError, console_ui

    rep = Reporter(console_ui, log_path=cache.new_run_dir(a.path, "cli") / "debug.log")
    try:
        if a.scan_only:
            plan = scan(ui, settings, rep)
            print("\nSources:")
            for r in plan.rows.values():
                print(f"  {r.label}:\n      {r.files_text()}")
            for d in plan.prior_deals:
                print(f"  possible earlier deal with this sponsor: {d}")
        else:
            out = generate(ui, settings, rep=rep)
            print(f"\nMemo saved:        {out.memo}")
            print(f"Review list saved: {out.review}" if out.review else
                  "The memo was saved, but the review list could not be created.")
    except UserFacingError as exc:
        print(f"\n{exc}")
        return 1
    finally:
        print(f"(diagnostics: {rep.log_path})")
        rep.close()
    return 0


if __name__ == "__main__":
    if "--ui-screenshot" in sys.argv:          # hidden: window with sample values -> PNG (visual check only)
        from memo_app.gui import ui_screenshot
        i = sys.argv.index("--ui-screenshot")
        sys.exit(ui_screenshot(sys.argv[i + 1] if i + 1 < len(sys.argv) else "memo_builder_ui.png"))
    if "--cli" in sys.argv:
        sys.exit(cli(sys.argv[1:]))
    from memo_app.gui import main
    main()
