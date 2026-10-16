# CLAUDE.md – Castellan Memo Builder

This file is context for Claude Code. Read README.md for the user-facing description.

## What this is
This is a Tkinter desktop app, packaged with PyInstaller. It drafts a lender's Loan Investment Memorandum (.docx) from a Dropbox deal folder. OpenAI (Responses API) handles document extraction and the prose sections. Everything else is deterministic.

## Pipeline (memo_app/pipeline.py)
```
pipeline.scan          -> storage.open_storage + ONE walk of the deal folder (LocalStorage | DropboxStorage; downloads in cache/downloads)
  locator.locate_all   -> one Selection per source, rules from resources/config/sources.yaml (ordered / fallback_any_folder / <deal>)
  _promote_signed_term_sheet -> the term-sheet PDF with signature evidence wins
  sources_plan.build_plan    -> SourcesPlan (rows the analyst edits in the GUI); plan.apply_saved() re-applies cache/plans/<deal>.json
  auto_one_pager_sheets      -> sources_plan.pick_sheets: only one-pager sheets about this address
  relationship.find_prior_deals -> names from PFS/resume FILE NAMES, other deals 2 levels deep
pipeline.generate(ui, settings, plan, rep)   -> uses the plan as given; NEVER rescans
  emails.collect       -> .eml/.msg + graph_mail.GraphMail (Outlook thread) -> EmailChain (dedupe, date order, 60k cap)
  extract.Extractor.run -> parallel LLM.extract(...) per source into pydantic schemas (memo_app/schemas.py); progress per step
  facts.index_quote_for -> rates.resolve: index value (analyst > NY Fed / FRED / settings > model > none); may fetch
  facts.build_facts    -> DealFacts: source priority, structure/participation, ratios, ALL business rules, flags
  writer.Writer        -> 5 AI sections: plan (SectionPlan) -> write -> check (language.py) -> one repair
  review.finish_flags  -> extraction errors / sources-check notes become flags; flags get their source files
  docx_builder.build_memo -> boilerplate.yaml (Jinja) + AI sections + tables; numbers the red remarks R1, R2, ...
  review_workbook.build_review_workbook -> "<memo> - Review.xlsx" from the SAME f.flags (build it AFTER the docx)
  write_run_files      -> facts.json + sources report in cache/runs/<deal>/<ts>/ (deleted after success unless debug.keep_run_files)
```
The output folder receives exactly two files: the .docx and "<memo> - Review.xlsx". If the workbook fails the
docx is kept and the analyst sees "The memo was saved, but the review list could not be created.".

## Design rules – keep these
- **The LLM never writes standard wording.** The green language from the team's templates lives in `resources/templates/boilerplate.yaml`. Only these sections go to the writer model: transaction_overview, business_background, submarket_overview, business_plan, sponsor_experience.
- **The LLM never decides business rules or does arithmetic.** Thresholds and conditional sentences (Mezz, pari passu, liquidity > $2MM, equity > 30%, PG / carve-out / completion guaranty, env findings → review, strong credit ≥ 700, LTV sanity 10–100%, value choice) belong in `facts.py` or in the `when:` expressions in boilerplate.yaml. Values are formatted before they reach the prompt (`writer._fmt_facts`).
- **Markers**, defined in `render.py`:
  - `⟦…⟧` = tool-filled value, `⟬…⟭` = AI text: tracked internally, printed as **normal text**
  - `⟪\u2063uid\u2063⟫` = red remark pointing at a `facts.Flag`: the ONLY highlight in the memo (red highlight, white bold), printed as `[REVIEW R3: title]`

  Jinja's `finalize` wraps every `{{ }}` automatically. Templates see the facts through `render.FactsView`, so a `None` field becomes `Missing(name)` and prints as `[MISSING Rn: <plain name>]` (names/sources in `facts.FIELD_LABELS`). Grammar words must use `{% if %}s{% endif %}`, never `{{ 's' if x }}`. `docx_builder.MemoDoc.add_rich` turns the markers into Word formatting. `**bold**` is also supported.
- **Every analyst item is a `facts.Flag`** - created with `f.flag(level, details, title=, section=, source=, key=)` in `facts.py`, the writer, `review.finish_flags` or the docx builder. Fields: `id` (R1, R2, ... assigned by `build_memo` in document order; flags not in the body get the next IDs), `level` (review | missing | check), `memo_section`, `title` (≤ 15 words, the inline remark text), `details` (full text), `source_key` (sources.yaml key), `source_files` (relative paths), `key` (link for `f.remark('<key>')`), `inline`. The docx and the workbook use the SAME objects, so never copy flags or create a remark without a flag: use `f.remark(key)`, `f.new_remark(...)`, `f.marker(flag)` or the `|flag` filter.
- **The review workbook** (`review_workbook.py`) has no token counts, model names or paths outside the deal folder.
- **Structured-output schemas:** every field is required-but-nullable (`Optional[...]`) with **no default values**, which is a strict-mode requirement. Enumerations use `Literal[...]` (`PropertyType`, `Structure`). Check with:
  `python -c "from openai.lib._pydantic import to_strict_json_schema as t; from memo_app import schemas as s; t(s.Appraisal)"`
- **Source priority** (from the analyst walkthrough):
  - signed term sheet > model for loan terms; a term sheet issued by another lender / pari passu → `f.loan` = Castellan's share, `f.total_loan` = whole loan, LTV in the memo = total-loan LTV (`f.memo_ltv`)
  - executed PSA > model for price
  - appraisal for values, size, comps and market; but the one pager's comps average / loan per SF win in the text so it matches the tables
  - PFS for financials; Eaglescreen for background (only sponsors / guarantors / key principals are named)
  - narrative (purpose, plan, exit): analyst Notes box > email chain (EmailInsights) > term sheet / model. Requested amounts or LTV/LTC in emails never replace term-sheet numbers; differences become "check" flags in `facts._apply_email_insights`.
- **Writer prompts never label sources in a way the model will echo.** Emails, the term sheet and the model are described as "background facts - do not cite"; the only citations allowed are a playbook's `allowed_attributions`. The check step rejects `banned_phrases`.
- **Analyst-facing text is plain English.** Pipeline code reports with `progress.Reporter`: `step()` / `done(ok, note)` for the UI (fixed labels in `progress.py`, notes like "Couldn't read the appraisal - ..."), `detail()` for debug.log only. Anything passed as `log` (storage, extract, writer, emails, graph, rates) is the detail channel. Never put paths, model names, token counts or tracebacks in step/done; `progress.sanitize` is a safety net, not the plan. Show `UserFacingError` messages as-is; anything else → generic message + traceback in debug.log.
- **Index rate:** GUI/CLI index override > term sheet index; "Fixed rate" in either means no index. The phrase in the memo comes from `rates.DISPLAY` via `f.index_display`. `effective_rate = max(index + spread, floor)` is computed in `rates.effective_rate`. Fetches (NY Fed, FRED) use `memo_app/net.py` (`httpx`, or `httpx2` with openai>=3), a 10s timeout, and must never raise. Tests stub them (`test_offline.offline_rates`).
- **Prior relationship** is the analyst's checkbox; `relationship.find_prior_deals` only suggests (GUI auto-ticks once per deal). Never walk the whole portfolio recursively through the Dropbox API.
- **Graph never fails the memo:** `GraphMail.fetch_chain` and `emails.collect` log and continue with local files.
- **Paths:** always use `paths.resource(...)` for resources (editable next to the frozen exe) and `paths.cache_dir()` / `cache.py` for anything temporary. Never write next to the memo. Never hard-code paths.
- **Windows:** no `%-d` in strftime (use `facts.memo_date`); open files with `os.startfile`.
- **GUI look:** pure tkinter/ttk ("clam" theme), no extra dependencies. All styling lives in `ui_theme.py`; `gui.py` uses
  `grid` only. Pixel sizes (window, column minimums) go through `App.px()` because DPI awareness is on. Inputs in a card
  share one grid (labels col 0, fields col 1 `sticky="ew"`, buttons col 2) so the right edges line up -
  `tests/test_gui_smoke.py` checks it.

## Commands
```bash
python -m pytest -q                          # offline tests (stub LLM, fake deal, temp cache) – run after any change
python main.py                               # GUI
python main.py --ui-screenshot ui.png        # hidden: window with demo values -> PNG (visual check of the layout)
python main.py --cli --scan-only --path "<deal>" --loan-type acquisition
python main.py --cli --path "<deal>" --loan-type acquisition --notes-file notes.txt --index "1M Term SOFR" --index-rate 4.3
python main.py --cli --path "<deal>" --loan-type acquisition --graph [--conversation-id "<id>"]
python tools/check_setup.py ["<dropbox path>"]
python tools/check_graph.py [--subject "..."] [--conversation-id "..."] [--eml file.eml]
python tools/harvest_memos.py --root "/Bridge Lending/Current Portfolio" --dropbox
pyinstaller memo_app.spec --noconfirm         # or build.bat
```

## Where to change things
| Change | File |
|---|---|
| Which folder/file feeds a section | `resources/config/sources.yaml` (`display_name` / `used_for` / `memo_sections` for the Sources table and the workbook; name defaults in `sources_plan.DEFAULT_NAMES`) |
| Sources table, Choose-source dialog, status area, Help menu, layout | `gui.py` (`App`, `ChooseSourceDialog`, `DoneDialog`; grid only, cards from `ui_theme.card`); data model `sources_plan.py` |
| Colours, fonts, button / field / table styles, DPI awareness, title bar | `memo_app/ui_theme.py` (palette constants, `font()`, `apply_theme`). Never hard-code a font or colour in gui.py - read them from ui_theme |
| Saved per-deal choices, one-pager sheet selection by address | `sources_plan.py` (`SourcesPlan.save/apply_saved`, `pick_sheets`, `subject_street`) |
| Progress steps and their labels | `progress.py` (`GENERATE_STEPS`, `SCAN_STEPS`); grouping of sources into steps `extract.GROUPS` |
| Cache location, retention, cleanup | `paths.cache_dir`, `cache.py`; `debug.keep_run_files`, `debug.cache_retention_days` in settings.yaml |
| Standard sentences, risk/highlight order and conditions, structure openers | `resources/templates/boilerplate.yaml` |
| Review list workbook (sheets, columns, fills, Status values, key figures) | `review_workbook.py` (`key_figures`, `STATUSES`, `FILLS`) |
| Analyst items: titles, sections, keys; red remark format | `facts.py` (`Flag`, `DealFacts.flag/remark/new_remark/missing_remark`, `FIELD_LABELS`, section constants); `render.py` |
| Online search findings | `schemas.AdverseSearch` / `AdverseFinding`; prompt in `extract.Extractor.adverse_search`; one Review flag per person in `facts.build_facts` |
| Excel cell display (number formats), one-pager number tidy-up | `text_extract.format_cell`; `extract.tidy_table_numbers` |
| Voice / accuracy rules | `resources/style/style_guide.md` |
| Per-section structure, variants, allowed attributions, banned phrases, length | `resources/style/playbooks/<section>.yaml` |
| Few-shot examples and how they are picked | `resources/style/examples/<loan_type>/<section>/*.txt` (YAML front matter); scoring in `examples.py` (`WEIGHTS`, `GOLD_BONUS`) |
| Plan / check rules (numbers, banned phrases, length, markdown, "we") | `language.py`; orchestration `writer.Writer.write_section` |
| Thresholds, models, insurance agent | `resources/config/settings.yaml` (`rules.strong_credit_min`, `rules.ltv_sanity_min/max`, `rules.comps_mismatch_pct`, ...) |
| Participations / pari passu, value choice, LTV sanity, credit wording, environmental Phase II, city/state, property phrase | `facts.py` (`_apply_structure`, value block, `_filter_one_pager_tables`, `city_state_from`, `property_phrase`); "share not found" opener `openers.structure.participation_no_share` in boilerplate.yaml |
| One-pager sheets: tables vs facts | `sources_plan.pick_sheets` (`sheets` = tables shown, `excluded_sheets` = other deals, never read); `extract.Extractor._one_pager` |
| New extracted field | `schemas.py` → `extract.INSTRUCTIONS` → `facts.py` → boilerplate/writer (+ `CANNED` in `tests/test_offline.py`) |
| Word layout (headings, tables, bullets, review notes) | `docx_builder.py`; base styles/logo/header in `resources/templates/memo_base.docx` |
| Add a GUI input | `gui.py` (`_build`, `_inputs`) + `facts.UserInputs` + `main.py` CLI args |
| Index names, memo phrase per index, fetch sources | `memo_app/rates.py` (`CHOICES`, `DISPLAY`, `classify`, `resolve`); Prime fallback `rules.prime_rate_pct`, `rates.auto_fetch` in settings.yaml |
| Prior-deal search (which sub-folders, stopwords for names) | `memo_app/relationship.py`; on/off `memo.prior_relationship_autodetect` |
| Email folders / file types | `resources/config/sources.yaml` → `emails` (`fallback_any_folder`, `<deal>` placeholder) |
| Email cleaning (banners, quoted history, disclaimers), chain cap | `memo_app/emails.py` (`clean_body`, `build_chain`); `emails.max_chain_chars` |
| Outlook / Graph lookup | `memo_app/graph_mail.py`; secrets `GRAPH_*` in `.env` (see `.env.example`) |
| Email-vs-term-sheet conflict rules | `facts._apply_email_insights`; tolerances `rules.email_amount_tolerance_pct`, `rules.email_ratio_tolerance_pts` |

## Open items
- Auto-detect loan type from the term sheet and model, and pre-select it in the GUI.
- Construction: the plan & cost review firm (`DealFacts.pcr_firm`) has no source yet; add a GUI field or a sources.yaml entry.
- Optionally upload the finished memo back to the deal's `Memo/` folder (needs the Dropbox `files.content.write` scope).
- Graph: find threads by property address when the deal folder has no .eml (today: Message-ID / subject of local emails, or a pasted Conversation ID).
- `.msg` parsing (`emails.parse_msg`) is not covered by the offline tests (no stdlib way to write a .msg fixture).
- Evaluation harness: regenerate memos for past deals and diff them against the final memos in each `Memo/` folder to tune prompts, playbooks and examples (and promote good harvested examples to `quality: gold`).
