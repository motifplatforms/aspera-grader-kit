"""Build two sample outputs for the board-memo task: a strong one and a weak one.

    python examples/board-memo/make_samples.py

The weak memo makes the mistakes graders exist to catch: a wrong EBITDA figure, bold Normal text faked as headings,
risks typed with hyphens instead of a list, a table missing rows, no source line and a second font.
"""
import os

from docx import Document

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "samples")
ROWS = [("Revenue", "USD 48.2m", "+8.1%"), ("Gross margin", "61.5%", "+1.7 pts"),
        ("Operating expenses", "USD 22.9m", "+3.6%"), ("EBITDA", "USD 6.6m", "+43.5%")]


def strong(path):
    d = Document()
    d.add_heading("Q3 2026 results", level=0)
    d.add_heading("Summary", level=1)
    d.add_paragraph("Revenue reached USD 48.2m, gross margin widened to 61.5% and EBITDA rose to USD 6.6m, "
                    "the best quarter this year.")
    t = d.add_table(rows=1, cols=3); t.style = "Table Grid"
    for c, h in zip(t.rows[0].cells, ("Metric", "Q3 2026", "Change vs Q2")): c.text = h
    for r in ROWS:
        cells = t.add_row().cells
        for c, v in zip(cells, r): c.text = v
    d.add_paragraph("Source: q3_results.csv (management accounts).")
    d.add_heading("Risks", level=1)
    d.add_paragraph("Freight costs could reverse part of the margin gain in Q4.", style="List Bullet")
    d.add_paragraph("Two retail partners renew contracts in December.", style="List Bullet")
    d.save(path)


def weak(path):
    d = Document()
    d.add_paragraph("Q3 2026 results")
    for name, body in (("Summary", "Revenue reached USD 48.2m and EBITDA was USD 5.9m. Margins improved."),):
        p = d.add_paragraph(); p.add_run(name).bold = True
        d.add_paragraph(body)
    t = d.add_table(rows=1, cols=3)
    for c, h in zip(t.rows[0].cells, ("Metric", "Q3", "Change")): c.text = h
    for r in ROWS[:2]:
        cells = t.add_row().cells
        for c, v in zip(cells, r): c.text = v
    p = d.add_paragraph(); p.add_run("Risks").bold = True
    r = d.add_paragraph().add_run("- Freight costs\n- Contract renewals"); r.font.name = "Comic Sans MS"
    d.save(path)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    strong(os.path.join(OUT, "board_memo_strong.docx"))
    weak(os.path.join(OUT, "board_memo_weak.docx"))
    print("wrote", OUT)
