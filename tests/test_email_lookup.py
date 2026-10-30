"""Deal -> Outlook conversation ID lookup: folder-name parsing, fuzzy matching, workbook reading."""
import openpyxl

from memo_app import email_lookup as el
from memo_app.email_lookup import LookupConfig, LookupRow


def _rows(*pairs):
    return [LookupRow(conversation_id=cid, values={"Name": name}) for name, cid in pairs]


def test_deal_name_from_path():
    assert el.deal_name_from_path(r"X\Y\42-18 147th Street") == "42-18 147th Street"
    assert el.deal_name_from_path("/Bridge Lending/Current Portfolio/170 Broadway/") == "170 Broadway"
    assert el.deal_name_from_path('  "C:\\deals\\88 Rose Way\\"  ') == "88 Rose Way"
    assert el.deal_name_from_path("170 Broadway") == "170 Broadway"
    assert el.deal_name_from_path("") == ""


def test_partial_ratio_is_substring_aware():
    # a street-only folder name still matches the full "street, city, STATE zip" address
    assert el.partial_ratio("42-18 147th Street", "42-18 147th Street, Flushing, NY 11355") == 1.0
    assert el.partial_ratio("170 Broadway", "170 Broadway, New York, NY 10038") == 1.0
    assert el.partial_ratio("", "anything") == 0.0
    assert el.partial_ratio("newsletter", "42-18 147th Street, Flushing, NY") < 0.8


def test_score_row_takes_best_column():
    cfg = LookupConfig(match_columns=["Name", "Property Address"])
    row = LookupRow(conversation_id="c", values={"Name": "Deal 1", "Property Address": "170 Broadway, New York, NY"})
    assert el.score_row("170 Broadway", row, cfg.match_columns) == 100


def test_match_threshold_and_order():
    cfg = LookupConfig(match_columns=["Name"], threshold=80)
    rows = _rows(
        ("170 Broadway, New York, NY 10038", "cid-broadway"),
        ("88 Rose Way, Somewhere, CA 90000", "cid-rose"),
        ("Quarterly investor newsletter", "cid-junk"),
    )
    m = el.match(rows, "170 Broadway", cfg)
    assert m and m[0].conversation_id == "cid-broadway" and m[0].score >= 80
    assert all(r.score >= 80 for r in m)
    assert "cid-junk" not in [r.conversation_id for r in m]


def test_match_empty_query_returns_nothing():
    cfg = LookupConfig(match_columns=["Name"])
    assert el.match(_rows(("170 Broadway", "c")), "   ", cfg) == []


def test_match_below_threshold_falls_back_to_closest():
    cfg = LookupConfig(match_columns=["Name"], threshold=80)
    rows = _rows(("Completely unrelated subject line", "c1"), ("Another unrelated row", "c2"))
    m = el.match(rows, "170 Broadway", cfg)
    assert m and all(r.score < 80 for r in m)          # nothing qualifies, but closest rows are still pickable


def test_load_rows_reads_workbook_and_skips_blank_ids(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "board_items"
    ws.append(["Name", "Broker", "Loan Request", "Loan Type", "Asset Type", "Conversation ID", "Property Address"])
    ws.append(["170 Broadway", "Jane", 2500000, "acquisition", "Office", "CID-1", "170 Broadway, New York, NY 10038"])
    ws.append(["No conv row", "Joe", 100, "refinance", "Retail", "", "1 Nowhere St"])   # no id -> skipped
    path = tmp_path / "board.xlsx"
    wb.save(path)

    cfg = LookupConfig(workbook=str(path))

    class S:
        has_dropbox = False

        def get(self, dotted, default=None):
            return default

    rows = el.load_rows(S(), cfg=cfg)
    assert len(rows) == 1
    r = rows[0]
    assert r.conversation_id == "CID-1"
    assert r.values["Name"] == "170 Broadway"
    assert r.values["Loan Request"] == "2500000"       # integer kept clean (no trailing .0)
    assert el.match(rows, "170 Broadway", cfg)[0].conversation_id == "CID-1"


def test_config_from_settings_defaults_and_overrides():
    class S:
        def get(self, dotted, default=None):
            return {"workbook": "board.xlsx", "threshold": 90} if dotted == "email_lookup" else default

    cfg = LookupConfig.from_settings(S())
    assert cfg.workbook == "board.xlsx"
    assert cfg.threshold == 90
    assert cfg.id_column == "Conversation ID"          # default kept
    assert cfg.match_columns == ["Name", "Property Address"]
