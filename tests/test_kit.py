import os, sys

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
TASK = os.path.join(ROOT, "examples", "board-memo")

from aspera_grader_kit import Check, Result, grade, inspect  # noqa: E402
from aspera_grader_kit import rubric as R  # noqa: E402


@pytest.fixture(scope="module")
def samples(tmp_path_factory):
    sys.path.insert(0, TASK)
    import make_samples
    d = tmp_path_factory.mktemp("samples")
    s, w = str(d / "strong.docx"), str(d / "weak.docx")
    make_samples.strong(s); make_samples.weak(w)
    return s, w


def test_strong_memo_scores_full_marks(samples):
    rep = grade(samples[0], TASK)
    assert rep["total"] == 100.0
    assert all(c["status"] == "pass" for c in rep["checks"].values())


def test_weak_memo_fails_the_right_checks(samples):
    rep = grade(samples[1], TASK)
    assert rep["total"] < 40
    failed = {k for k, c in rep["checks"].items() if c["status"] == "fail"}
    assert {"R2", "R4", "C1", "C2", "C4"} <= failed
    r4 = rep["checks"]["R4"]["elements"]
    assert [e["item"] for e in r4 if e["status"] == "fail"] == ["Operating expenses", "EBITDA"]


def test_missing_file_is_gated_to_zero():
    rep = grade(os.path.join(TASK, "does_not_exist.docx"), TASK)
    assert rep["total"] == 0.0 and rep["gates"][0]["id"] == "missing"


def test_unreadable_file_is_gated_to_zero(tmp_path):
    bad = tmp_path / "board_memo.docx"; bad.write_text("not a zip")
    rep = grade(str(bad), TASK)
    assert rep["total"] == 0.0 and rep["gates"][0]["id"] == "unreadable"


def test_items_count_one_unit_each():
    r = Result.items([("a", True, ""), ("b", False, ""), ("c", True, "")])
    assert r.status == "fail" and r.units() == (2, 3)


def test_empty_category_hands_points_to_the_rest():
    rub = {"categories": {"req": {"points": 60, "per_item": True}, "craft": {"points": 40}}}
    checks = [Check("R1", "req", "", None)]
    out = R.score({"R1": Result.items([("x", True, ""), ("y", False, "")])}, checks, rub)
    assert out["categories"]["req"]["possible"] == 100.0 and out["total"] == 50.0


def test_cap_gate():
    rub = {"categories": {"req": {"points": 100}}, "gates": {"fabricated": {"cap": 40}}}
    out = R.score({"R1": Result.ok()}, [Check("R1", "req", "", None)], rub, gates=[{"id": "fabricated", "evidence": ""}])
    assert out["total"] == 40.0


def test_pptx_and_xlsx_inspectors(tmp_path):
    from openpyxl import Workbook
    from pptx import Presentation
    from pptx.util import Inches, Pt
    prs = Presentation(); s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Revenue grew 8%"
    tb = s.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1)).text_frame
    tb.text = "Revenue: 48.2"; tb.paragraphs[0].runs[0].font.size = Pt(9)
    p = str(tmp_path / "d.pptx"); prs.save(p)
    f = inspect(p)
    assert f["slides"][0]["title"] == "Revenue grew 8%" and 48.2 in f["numbers"]
    assert f["slides"][0]["min_font_pt"] == 9

    wb = Workbook(); ws = wb.active
    ws["A1"], ws["A2"], ws["A3"] = 10, 20, "=A1*1.07+A2"
    x = str(tmp_path / "m.xlsx"); wb.save(x)
    f = inspect(x, with_recalc=False)
    sh = f["sheets"][0]
    assert sh["formulas"] == 1 and sh["n_constants_in_formulas"] == 1


def test_broken_check_becomes_error_not_crash():
    from aspera_grader_kit.checks import Context, run
    boom = Check("X", "req", "", lambda ctx: 1 / 0)
    out = run([boom], Context("p", {}))
    assert out["X"].status == "error" and "ZeroDivisionError" in out["X"].evidence
