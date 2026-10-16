"""Language system (Part E): playbooks, tagged examples, plan -> write -> check. Offline."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from memo_app import schemas as S  # noqa: E402
from memo_app.examples import Example, parse_example, select  # noqa: E402
from memo_app.extract import Extracted  # noqa: E402
from memo_app.facts import UserInputs, build_facts  # noqa: E402
from memo_app.language import allowed_numbers, check_text  # noqa: E402
from memo_app.render import FLAG_O, strip_markers  # noqa: E402
from memo_app.settings import Settings, load_yaml_resource  # noqa: E402
from memo_app.writer import StylePack, Writer  # noqa: E402
from test_offline import CANNED, STUB_TEXT  # noqa: E402

SETTINGS = Settings(raw={"memo": {"few_shot_examples_per_section": 2}})
GOOD_PLAN = STUB_TEXT["business_plan"]
BAD_PLAN = GOOD_PLAN.replace("The borrower is acquiring", "As described in the email chain, the borrower is acquiring")


class ScriptedLLM:
    """extract() returns a plan; write() returns the scripted texts in order."""

    def __init__(self, texts: list[str], plan: S.SectionPlan | None = None):
        self.texts, self.plan, self.writes = list(texts), plan or CANNED["SectionPlan"], []

    def extract(self, schema, instructions, text=None, images_b64=None):
        assert schema is S.SectionPlan
        return self.plan

    def write(self, system, user):
        self.writes.append(user)
        return self.texts.pop(0)


def _deal():
    x = Extracted(term_sheet=CANNED["TermSheet"], model=CANNED["UWModel"], appraisal=CANNED["Appraisal"],
                  pfs=CANNED["PFSResult"], resume=CANNED["ResumeFacts"])
    f = build_facts(x, UserInputs(deal_path=".", loan_type="acquisition", exit_strategy="Lease-up then refinance"),
                    Settings(raw={}))
    return f, x


def _writer(llm) -> Writer:
    return Writer(llm, SETTINGS, load_yaml_resource("templates", "boilerplate.yaml"), log=lambda m: None)


def test_check_catches_email_chain_and_one_repair_fixes_it():
    f, x = _deal()
    llm = ScriptedLLM([BAD_PLAN, GOOD_PLAN])
    w = _writer(llm)
    out = w.write_section("business_plan", f, x)
    assert len(llm.writes) == 2                                                  # draft + ONE repair
    assert 'Remove the phrase "email chain"' in llm.writes[1] and "YOUR DRAFT:" in llm.writes[1]
    assert strip_markers(out) == GOOD_PLAN and FLAG_O not in out
    assert not [fl for fl in f.flags if "style check" in fl.title]


def test_still_failing_after_repair_keeps_text_with_short_red_remark():
    f, x = _deal()
    llm = ScriptedLLM([BAD_PLAN, BAD_PLAN])
    w = _writer(llm)
    out = w.write_section("business_plan", f, x)
    assert len(llm.writes) == 2 and "[REVIEW: wording needs a look (style check)]" in strip_markers(out, f)
    fl = [fl for fl in f.flags if "style check" in fl.title][0]
    assert "email chain" in fl.details and fl.memo_section == "Business Plan" and fl.source_key == "model"


def test_plan_sentence_with_unknown_number_is_dropped():
    f, x = _deal()
    plan = S.SectionPlan(sentences=[
        S.PlannedSentence(purpose="purchase price", facts_used=["$3,000,000 purchase price"]),
        S.PlannedSentence(purpose="invented rent", facts_used=["rents of $47.50/SF", "occupancy of 93%"])])
    llm = ScriptedLLM([GOOD_PLAN], plan)
    w = _writer(llm)
    w.write_section("business_plan", f, x)
    assert "1. purchase price [facts: $3,000,000 purchase price]" in llm.writes[0]
    assert "invented rent" not in llm.writes[0]
    assert any("invented rent" in fl.details and "47.5" in fl.details and "93" in fl.details for fl in f.flags)


def test_check_text_rules():
    pb = StylePack().playbook("business_background")
    allowed = allowed_numbers("$2,240,000 loan; 25 years; 2025")
    ok = ("As per the sponsor’s bio shared with Castellan, Mr. Smith has 25 years of experience and arranged a "
          "$2.24MM loan in 2025. " + "He manages a portfolio of industrial properties across the region. " * 6)
    assert check_text(ok, pb, allowed) == []                     # allowed attribution contains "per the"
    bad = "**We** think this robust sponsor has 31 years. Per the term sheet he is great."
    problems = " | ".join(check_text(bad, pb, allowed))
    for expected in ('"robust"', '"term sheet"', '"per the"', "Too short", "31", "markdown", 'Do not write "we"'):
        assert expected in problems, expected
    assert check_text("We believe " + ok, pb, allowed) == []
    assert any("Begin exactly" in p for p in check_text(ok, pb, allowed, opening="Castellan Capital will provide"))


def test_playbook_system_prompt_uses_variants_and_never_labels_emails():
    f, x = _deal()
    system = StylePack().system("transaction_overview", f)
    assert "Purchase price; WITH renovation/capex" in system            # acquisition variant
    assert "Mention the junior loan" in system                          # structure = mezz
    assert "Never use these words or phrases" in system and '"email chain"' in system
    assert "Length: 55-102 words." in system
    w = _writer(ScriptedLLM([]))
    f.email_insights = {"loan_purpose": "acquisition"}
    src = w._sources_for("transaction_overview", f, x)
    assert "background facts" in src.lower() and "email" not in src.lower()


def test_example_tags_and_selection_score():
    tags, text = parse_example("# property_type: industrial\n# source: Memo - X.docx\nBody text")
    assert tags == {"property_type": "industrial", "source_memo": "Memo - X.docx"} and text == "Body text"
    tags, text = parse_example("---\nloan_type: Construction\nstructure: pari_passu\nquality: gold\n"
                               "source_memo: Memo - Y.docx\n---\nBody")
    assert tags["loan_type"] == "construction" and tags["source_memo"] == "Memo - Y.docx" and text == "Body"
    ex = lambda name, **t: Example(Path(name), "transaction_overview", name, t)  # noqa: E731
    pool = [ex("a", loan_type="acquisition", structure="senior", property_type="retail", quality="silver"),
            ex("b", loan_type="construction", structure="pari_passu", property_type="condominium", quality="silver"),
            ex("c", loan_type="acquisition", structure="pari_passu", property_type="condominium", exit="sellout", quality="gold"),
            ex("d", loan_type="acquisition", structure="pari_passu", property_type="office", quality="silver")]
    deal = {"loan_type": "acquisition", "structure": "pari_passu", "property_type": "condominium", "exit": "sellout"}
    assert [e.text for e in select(pool, deal, 3)] == ["c", "d", "b"]     # 11, 6, 5 (a = 3)
    seeds = StylePack().examples("transaction_overview", _deal()[0], 4)
    assert seeds[0][0].startswith("acquisition / mezz / industrial")      # same loan type + structure + type first
