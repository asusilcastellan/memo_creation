# Castellan Memo Builder

This tool drafts a Castellan **Loan Investment Memorandum** (.docx) from a deal folder in Dropbox.

You give it a deal folder path and a loan type: Acquisition, Construction, Refinance or Note-on-Note. It then:

1. Finds the right file for each memo section, using the rules from Vasist's walkthrough, and shows them in an editable **Sources** table so you can change any pick before generating.
2. Reads each file with OpenAI and pulls out structured data.
3. Computes LTV, LTC, loan basis, equity % and the other ratios, and applies the template rules (Mezz, pari passu participations, liquidity > $2MM, equity > 30%, PG vs carve-out, environmental review, credit wording, LTV sanity).
4. Fills in Castellan's standard wording, and has the AI write only the free-text sections in the team's voice (plan → write → check).
   It also reads the broker / borrower **email chain** (`.eml` / `.msg` files in the deal folder and, optionally, the Outlook thread through Microsoft Graph), computes the **effective rate at closing** for any index (SOFR, Term SOFR, Prime, AMERIBOR, fixed), and **suggests a prior relationship** when the sponsor's name appears in other deal folders.
5. Saves **two files**: the draft memo (`Memo - <address> (DRAFT <date>).docx`) and the analyst's **review list** (`Memo - <address> (DRAFT <date>) - Review.xlsx`). Short red remarks in the memo carry IDs (R1, R2, ...) that match the rows of the review list, where the details are. Diagnostics stay in a cache folder that cleans itself.

## 1. Setup (Windows)

```bat
cd castellan_memo
py -3.12 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env      &  REM then edit .env
python tools\check_setup.py  &  REM verifies the OpenAI key, model names and Dropbox access
python main.py               &  REM starts the app
```

### OpenAI
Put `OPENAI_API_KEY` in `.env`. Default models are set in `resources/config/settings.yaml`:

| Role | Default | What it does |
|---|---|---|
| `extract_model` | `gpt-5.4-mini` | Reads term sheets, appraisals and other documents into JSON, and plans each AI section. Cheap and fast. |
| `write_model` | `gpt-5.5` | Writes the prose sections in Castellan's voice. |

If `check_setup.py` says a model is not found, use one from the list it prints. You can override either model in `.env` with `OPENAI_MODEL_EXTRACT` / `OPENAI_MODEL_WRITE`. A memo makes roughly 25–30 calls (one plan + one write per AI section, plus a repair call when a section fails the style check).

### Dropbox
The **deal folder** box accepts either kind of path:
- **A folder on this PC.** If the Dropbox desktop app syncs the folder, browse to it. No API is needed.
- **A Dropbox path**, such as `/Bridge Lending/Current Portfolio/150 E Aurora St`. This uses the Dropbox API.

For the API, use a **refresh token**. Access tokens generated in the App Console expire after about 4 hours.

1. In the App Console, open your app's *Permissions* tab and tick `files.metadata.read`, `files.content.read` and `account_info.read`. Then click *Submit*.
2. Put `DROPBOX_APP_KEY` and `DROPBOX_APP_SECRET` in `.env`.
3. Run `python tools\dropbox_get_refresh_token.py` and paste the printed `DROPBOX_REFRESH_TOKEN` into `.env`.
4. Run `python tools\check_setup.py "/Bridge Lending/Current Portfolio/<a deal>"` to confirm the tool can list the folder.

`DROPBOX_USE_TEAM_ROOT=true` makes paths start from the team space, the same way the desktop app shows them.

### Outlook email chain (Microsoft Graph, optional)
Without this, the tool still reads every `.eml` / `.msg` file saved in the deal folder. With it, the tool also pulls the whole Outlook thread behind those emails.

1. Azure portal → *App registrations* → *New registration* (single tenant).
2. *API permissions* → *Add a permission* → *Microsoft Graph* → **Application permissions** → **`Mail.Read`** → *Grant admin consent* (a Microsoft 365 admin must click this).
   - Optional but recommended: limit the app to the deal mailboxes with an Exchange [application access policy](https://learn.microsoft.com/graph/auth-limit-mailbox-access).
3. *Certificates & secrets* → *New client secret*; copy the **Value**.
4. In `.env` set `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET` and `GRAPH_MAILBOXES` (comma-separated mailboxes to search, e.g. `deals@castellanre.com,analyst@castellanre.com`).
5. Run `python tools\check_graph.py` (sign-in + newest messages), `--subject "150 E Aurora"` (find a thread and its Conversation ID) or `--conversation-id "..."` (read a thread as the memo would).

How the thread is found, for each mailbox: the Message-ID of each local `.eml` is looked up (`internetMessageId`), then every message of that conversation is read. If no Message-ID matches, the tool searches for the subject (without RE:/FW:). If you paste a **Conversation ID** in the app, that conversation is read directly. Graph and local messages are merged, duplicates removed, and ordered by date. If Graph is down or the secret is wrong, the reason (e.g. the Azure AD error) goes to the diagnostics log and the memo continues with the local files.

### Index rates
- **SOFR / 30-day average SOFR**: NY Fed public API, no key needed.
- **1M / 3M Term SOFR (CME)**: no free API. Type the value; if you leave it empty, overnight SOFR is used as a proxy and the review list flags it.
- **WSJ Prime**: `rules.prime_rate_pct` in `resources/config/settings.yaml` (**update it when the Fed moves rates**), or the live FRED series DPRIME if `FRED_API_KEY` is in `.env` (free key from FRED).
- **AMERIBOR**: manual only.
- `rates.auto_fetch: false` in settings.yaml stops all rate fetching. A failed fetch never stops the memo: the floor rate is used and flagged.

### Prior relationship detection
After **1. Check sources**, the tool takes the names in the PFS / resume **file names** (e.g. `PFS Shaya Margareten.pdf`) and looks in the **other deal folders next to this one** (e.g. `/Bridge Lending/Current Portfolio`) for file names containing the same surname (4+ letters). Only each deal's root and its PFS / Resume / Term Sheet / Background folders are listed, so it stays fast over the Dropbox API. Matches tick "Castellan has done a loan with this sponsor before" and are listed under it; untick it if it is a different person. Turn it off with `memo.prior_relationship_autodetect: false`.

## 2. Using the app

### The window

| Area | What is there |
|---|---|
| Header (navy) | The two steps; the orange chip shows where you are: ① Check sources, then ② Generate memo |
| **Deal** card | Deal folder, loan type, property type, exit strategy, where to save the memo |
| **Rates & sources** card | Index and current index % (**Fetch**), Outlook conversation ID, and the three options: prior relationship, online adverse-media search, read the email chain from Outlook |
| **Notes for the memo** | Loan purpose, business plan, exit - this wins over everything else |
| Buttons | **Open output folder** · **1. Check sources** (blue) · **2. Generate memo** (orange) |
| **Sources** card | The Sources table (see below) |
| Footer | Orange progress bar and the current step; warnings in red, "✓ Memo and review list saved" in green when done |

Keyboard: **Enter** in the deal-folder box checks the sources; **Ctrl+G** generates the memo once the sources are checked. Tab moves through the fields in the order they appear.

**Screenshot for a visual check** (not shown in the app): `python main.py --ui-screenshot ui.png` opens the window with sample values and a demo Sources table (no network, no deal folder needed), saves a PNG and exits. It uses Pillow's `ImageGrab` when Pillow is installed and a built-in Windows capture otherwise.

### Step by step

1. Paste (or browse to) the deal folder and choose the loan type. Optionally fill in the property type, exit strategy, **Index** + **current index rate %** (**Fetch** fills SOFR or Prime), the Outlook options and the **Notes** box. Notes win over everything else for the narrative (purpose, plan, exit).
2. Click **1. Check sources**. The **Sources** table shows one row per document type in three columns: **Document**, **Used in memo section(s)** and **File(s) used**. Hover the document name to see what it is used for, and any cut-off cell to see its full text; columns can be resized and the table scrolls sideways. Rows in red say "Not found - double-click to choose". Which memo sections each document feeds is set in `sources.yaml` (`memo_sections`).
3. To change a row, double-click it (or select it and click **Change…**). In the **Choose source** dialog pick the folder, then the file(s) (newest first; tick *Show all file types* for other extensions). For the **One pager**, choose the sheet(s) about this property: the sheets that mention the deal's street number and name are pre-selected, and sheets about another address are marked "(other property)". Buttons: **Use selected**, **None (skip this document)**, **Reset to automatic**, **Cancel**.
   Your choices are saved per deal and re-applied the next time you check the same deal (shown as "(saved choice)") as long as the files still exist.
   The one-pager sheet choice only decides which **tables** go into Loan Financials. Facts (Castellan's share, total debt, lead lender, rate terms, loan metrics) are read from every sheet of the one pager and the model, except sheets that clearly belong to another deal.
4. Click **2. Generate memo** (enabled after checking sources; changing the deal folder clears the table). The status area shows plain steps with a progress bar: Checking the deal folder · Reading the term sheet · Reading the appraisal · Reading sponsor documents · Reading third-party reports · Reading emails · Checking for earlier deals with this sponsor · Writing the memo sections · Building the Word document · Done. A warning icon with a short note means a document could not be read; the memo marks that part for review.
5. When it finishes: **"Memo and review list saved."** with **Open memo**, **Open review list**, **Open folder** and **Close**. If the review list could not be created the memo is still saved and the app says so.

**In the memo** only short red remarks remain, each with an ID, e.g. `[REVIEW R3: adverse items in background reports]`, `[MISSING R7: appraiser]`, `[MISSING R12: insert map]`. Fix each one in Word and remove the highlight before circulating.

**The review list** (`... - Review.xlsx`):

| Sheet | What it holds |
|---|---|
| **Review items** | One row per item: ID (same as the red remark; items that are not in the memo text get the next IDs), Priority (Review = red, Missing = amber, Check = grey), Memo section, Item, Details, Source document and file(s), **Status**, **Analyst comment**. Sorted Review → Missing → Check. |
| Online search | Each finding of the adverse-media search with person, category, severity and a clickable link to the source |
| Background checks | Each person in the background reports, their role in the deal (people who are not sponsors / guarantors are marked and never named in the memo), credit score and records |
| Sources used | Which file fed which memo sections, its folder and date, and whether it was picked automatically, by you, or from a saved choice |
| Key figures | Loan, share, total loan, price, budget, values, LTV/LTC, $/SF, equity, liquidity, credit scores, index/spread/floor/effective rate - as real Excel numbers |
| Email chain | The emails read (date, sender, subject, local file or Outlook) and the conflicts with the term sheet |

How to use it: work down **Review items**, fix each item in the memo (find it with Ctrl+F on its ID), write what you did in **Analyst comment** and set **Status** to *Resolved* or *Not applicable* with the dropdown (new rows start as *Open*). Filter the Status column on *Open* to see what is left; the list is done when nothing is Open. Send the workbook with the memo to whoever reviews your draft.

**Help menu** (for IT): *Open diagnostics folder* opens the cache folder with each run's `debug.log`; *Clear cache* deletes downloads and logs (saved source choices are kept).

Command line, useful for testing and for Claude Code:
```bat
python main.py --cli --scan-only --path "/Bridge Lending/Current Portfolio/150 E Aurora St" --loan-type acquisition
python main.py --cli --path "..." --loan-type construction --notes-file emails.txt --index "1M Term SOFR" --index-rate 4.31 --web-search --out C:\Memos
python main.py --cli --path "..." --loan-type acquisition --graph --conversation-id "AAQkAD..."
```
The CLI uses the saved source choices of the deal and prints the same plain steps; it prints where the diagnostics log is.

## 3. Where each section comes from

The rules are in `resources/config/sources.yaml` and can be edited without changing code; the **Sources** table overrides them per deal.

| Memo item | Folder | File rule |
|---|---|---|
| Loan amount, junior loan, rate, floor, recourse, reserves, participation | Term Sheet | Signed PDF wins (signature field, DocuSign stamp or /s/; the model also checks the signature page). `Archive` is ignored. Castellan loan = blended amount − junior loan; for a pari passu / participated loan (or a term sheet issued by another lender) Castellan's share. If the share is nowhere to be found, the memo says "Castellan Capital will participate in a $X senior loan led by ..." with a red `[MISSING Rn: Castellan's share]` and never presents the whole loan as Castellan's. |
| Purchase price | PSA | "Executed PSA"; cross-checked against the model and flagged if they differ |
| Capex, sources & uses, property facts | `*model*` / `*uw*` / `*underwriting*` / `*pro forma*` / `*sizing*` .xlsx/.xlsm in the deal root, then `Model` / `Underwriting` folders, then `Memo` (one pager excluded) | Newest modified |
| Loan Financials tables | Memo → `one pager*.xlsx` | Tables only from the sheets chosen in the Sources table (facts from all sheets of this property); tables that mention another address are dropped (review item); a table repeated side by side is output once. Numbers are read as Excel displays them ($869, 48.95%); stray decimals are rounded ($/SF columns to whole dollars, other $ to cents, % to 2 decimals). Its comps average and loan per SF are used in the highlights so the text matches the tables. |
| Business background, sponsor experience | Resume | All files |
| Personal financials | PFS & SREO | PFS files (REO schedules excluded) |
| Background & credit | Background Authorisation & Credit Report | "Final Results…" (consent and disclosure forms excluded). Only sponsors / guarantors / key principals are named; "strong credit" only if every score ≥ 700 (`rules.strong_credit_min`). |
| Submarket, values, comps, building SF, appraiser | Appraisal | "Final Appraisal" (engagement letter and invoice excluded); long reports are trimmed to the relevant pages. Condo construction uses the gross sellout value. An LTV above 100% or below 10% is never printed. |
| Environmental risk | Environmental / Phase 1 | Phase I **and** Phase II reports. Only a recommendation for further investigation/remediation is an "issue"; a Phase II that clears the Phase I uses its own standard sentence. |
| Insurance risk | Insurance | Evidence / ACORD / certificates |
| Pictures | Site Visit, or the photo embedded in the model | — |
| Loan purpose, business plan, exit (with the Notes box) | Emails / Email / Correspondence / Borrower material X / `<deal> - Shared Folder`, else anywhere | All `.eml` / `.msg` (up to 30), plus the Outlook thread via Graph; ordered by date, capped at ~60k characters |
| Map | — | Manual (red placeholder) |

**Recommended naming conventions for the team.** Vasist offered to standardise, and these make every pick exact:
- `<Address> Model.xlsx` in the deal root, with old versions moved to `Archive`
- The signed term sheet as a PDF, ideally with `signed` in the name
- `Executed PSA.pdf`
- The appraisal file containing `Final Appraisal`
- One name for the environmental folder: `Environmental`

## 4. Tuning the language

There are four levers, from safest to most powerful.

1. **Standard wording** (`resources/templates/boilerplate.yaml`). This holds the green text from the AI templates: risks, personal financials, background and credit, highlights, and the Transaction Overview openers (per loan type, and per structure for pari passu / participations). The AI never rewrites it. You can edit sentences freely as long as the `{{ }}` / `{% %}` parts stay intact. Red remarks are `{{ f.remark('<item key>') }}` (linked to an item the code raised) or `{{ 'REVIEW: short text'|flag }}`; keep them short - the details go to the review list.
2. **Section playbooks** (`resources/style/playbooks/<section>.yaml`): purpose, required elements in order, guidance per loan type / structure / exit, the only allowed ways to cite a source, banned phrases (e.g. "email chain", "term sheet", "robust"), target length and Vasist's notes. The global voice and accuracy rules stay in `resources/style/style_guide.md`.
3. **Few-shot examples** (`resources/style/examples/<loan_type>/<section>/*.txt`). Each file starts with a YAML header:
   ```
   ---
   loan_type: acquisition
   property_type: industrial
   structure: mezz          # senior | pari_passu | participation | mezz | note_on_note
   exit: refinance          # refinance | sale | sellout
   quality: gold            # gold (hand-picked) | silver (harvested, unreviewed)
   source_memo: Memo - 150 E Aurora St.docx
   ---
   ```
   The writer picks the top `memo.few_shot_examples_per_section` examples of the same section by score: +3 same loan type, +3 same structure, +2 same property type, +1 same exit, +2 gold. Collect examples from past memos (they are tagged automatically as `silver`):
   ```bat
   python tools\harvest_memos.py --root "/Bridge Lending/Current Portfolio" --dropbox
   ```
   Then **review the folder**: delete weak examples and change the best ones to `quality: gold`.
4. **How each AI section is produced**: (1) *plan* - the model lists each sentence with the facts it uses; any sentence with a number that is not in the facts/sources is dropped (review note); (2) *write* - the writer model writes from the approved plan, the playbook and the examples; (3) *check* - code checks banned phrases, length (target ±30%), unknown numbers, markdown and "we" (only "We believe" is allowed). One repair call fixes the listed problems; if it still fails, the text is kept with a short red remark and the problems go to the review list. All prompts and outputs are in the run's `debug.log`.

Rules that depend on numbers (thresholds, insurance agent name) are in `resources/config/settings.yaml`.

After a build, the `resources/` folder sits next to `CastellanMemo.exe`, so the team can change wording without rebuilding.

## 5. Cache and diagnostics

Only the `.docx` and its `- Review.xlsx` are written to the output folder. Everything else lives in the cache folder: `%LOCALAPPDATA%\CastellanMemo\cache` on Windows (`~/.cache/castellan_memo` elsewhere; `CASTELLAN_CACHE_DIR` overrides it):

| Folder | Content |
|---|---|
| `downloads/` | Dropbox files fetched for reading |
| `runs/<deal>/<timestamp>-check` / `-memo` | `debug.log` of each run (file picks, prompts, model outputs, errors with tracebacks). `facts.json` and `sources report.md` are written too but deleted after the memo is saved, unless `debug.keep_run_files: true`. |
| `plans/<deal>.json` | Your source choices for the deal (re-applied on the next check) |

On start the app deletes downloads and run folders older than `debug.cache_retention_days` (default 7). **Help → Clear cache** deletes them now.

## 6. Build the .exe

```bat
.venv\Scripts\activate
build.bat
```
The output is `dist\CastellanMemo\` (one folder). Put `.env` next to `CastellanMemo.exe` and share the whole folder.

The exe has no console window. Problems are in the run's `debug.log` (Help → Open diagnostics folder). To see tracebacks in a console while debugging, set `console=True` in `memo_app.spec`, or run `python main.py`.

## 7. Tests
```bat
python -m pytest -q
```
All tests are offline (stub model, no keys, no network; the cache goes to a temporary folder):
- `tests/test_offline.py` - a fake deal folder shaped like 150 E Aurora (including two `.eml` files) through the whole pipeline into a real .docx.
- `tests/test_sources_plan.py` - Sources table overrides, saved choices, memo sections, one-pager sheet filtering, generate without a rescan.
- `tests/test_progress.py` - the analyst only sees plain steps (no paths / model names); details go to debug.log.
- `tests/test_output.py` - exactly the .docx + review .xlsx are written, remark IDs match the workbook rows, Status dropdown, hyperlinks, numeric key figures, only red highlights, cache cleanup.
- `tests/test_onepager_numbers.py` - one-pager facts from unselected sheets, participations without a known share, Excel number formats.
- `tests/test_accuracy.py` - pari passu, LTV sanity, value choice, credit wording, Phase II, one-pager consistency, city/state, model discovery.
- `tests/test_language.py` - playbooks, example selection, plan/write/check with one repair.
- `tests/test_rates.py`, `tests/test_relationship.py`, `tests/test_emails.py` - index rates, prior-deal search, email chain (Graph mocked).

Run them whenever you change rules, templates or the Word layout.

## Known limits / next steps
- Loan type is chosen by the user. The extracted term sheet and model could suggest it automatically.
- Construction Completion Risk needs the plan-and-cost-review firm. It is not in the folder list from the walkthrough, so it shows as a red placeholder.
- The one-pager tables are rebuilt as Word tables, not pasted as Excel pictures.
- Outlook threads are found from the Message-IDs / subjects of the `.eml` files in the deal folder (or a pasted Conversation ID); there is no search by property address yet.
- `.msg` parsing depends on the `extract-msg` package; the offline tests cover `.eml` only.
- Scanned PDFs are sent to the model as page images (up to `documents.max_image_pages`), which works but costs more tokens.
