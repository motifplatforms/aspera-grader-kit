"""Checks for the board-memo example. Each tests one thing and says why it passed or failed."""
import re

from aspera_grader_kit import Result, check

T = "board-memo"
HEADING = re.compile(r"^(heading \d|title)$", re.I)


def paragraphs(ctx): return [b for b in ctx.facts["blocks"] if b["type"] == "p"]


def tables(ctx): return [b for b in ctx.facts["blocks"] if b["type"] == "table"]


def section_text(ctx, name):
    """Text from the heading called `name` up to the next heading."""
    out, inside = [], False
    for b in ctx.facts["blocks"]:
        if b["type"] == "p" and HEADING.match(b["style"] or ""):
            inside = b["text"].strip().lower() == name.lower()
            continue
        if inside:
            out.append(b["text"] if b["type"] == "p" else " ".join(c["text"] for r in b["rows"] for c in r["cells"]))
    return "\n".join(out)


def has_number(text, value):
    return re.search(r"(?<![\d.])%s(?![\d])" % re.escape(value), text) is not None


# ---- Task requirements (counted item by item) ----

@check("R1", "requirements", "Memo is titled as asked", registry=T)
def title(ctx):
    first = next((p for p in paragraphs(ctx) if p["text"].strip()), None)
    ok = first is not None and first["text"].strip().lower() == ctx.key["title"].lower()
    return Result.when(ok, "first line: %r" % (first["text"][:60] if first else None))


@check("R2", "requirements", "Required sections exist as real headings", registry=T)
def sections(ctx):
    heads = {p["text"].strip().lower(): p["style"] for p in paragraphs(ctx) if HEADING.match(p["style"] or "")}
    return Result.items((s, s.lower() in heads, "heading style" if s.lower() in heads else "no heading called %s" % s)
                        for s in ctx.key["sections"])


@check("R3", "requirements", "Summary states each headline value", registry=T)
def summary_values(ctx):
    s = section_text(ctx, "Summary")
    return Result.items((v["item"], has_number(s, v["value"]), v["value"]) for v in ctx.key["summary_values"])


@check("R4", "requirements", "Metrics table has every metric with its value and change", registry=T)
def metrics_table(ctx):
    rows = [" | ".join(c["text"] for c in r["cells"]) for t in tables(ctx) for r in t["rows"]]
    def found(k):
        row = next((r for r in rows if k["item"].lower() in r.lower()), None)
        if row is None: return False, "no row"
        return has_number(row, k["value"]) and has_number(row, k["change"]), row[:80]
    return Result.items((k["item"],) + found(k) for k in ctx.key["table_rows"])


@check("R5", "requirements", "Risks section lists enough risks as bullets", registry=T)
def risks(ctx):
    items, inside = [], False
    for p in paragraphs(ctx):
        if HEADING.match(p["style"] or ""):
            inside = p["text"].strip().lower() == "risks"; continue
        if inside and p["text"].strip(): items.append(p)
    # a real list: direct numbering on the paragraph, or a list style (List Bullet, List Paragraph ...)
    bullets = [p for p in items if p["num"] is not None or (p["style"] or "").startswith("List")]
    return Result.when(len(bullets) >= ctx.key["min_risks"], "%d bulleted risks, %d plain paragraphs" % (len(bullets), len(items) - len(bullets)))


# ---- Craft (each check counts once) ----

@check("C1", "craft", "Source file is cited", registry=T)
def source(ctx):
    return Result.when(ctx.key["source"].lower() in ctx.text.lower(), "looked for %s" % ctx.key["source"])


@check("C2", "craft", "No faked headings (bold Normal text standing in for a heading)", registry=T)
def no_fake_headings(ctx):
    fake = [p["text"][:40] for p in paragraphs(ctx)
            if (p["style"] or "").lower() == "normal" and p["runs"] and all(r["b"] for r in p["runs"]) and len(p["text"]) < 60]
    return Result.when(not fake, "fake headings: %s" % fake if fake else "none")


@check("C3", "craft", "No tracked changes or comments left in the file", registry=T)
def clean(ctx):
    n = len(ctx.facts["ins"]) + len(ctx.facts["del"]) + len(ctx.facts["comments"])
    return Result.when(n == 0, "%d tracked changes/comments" % n)


@check("C4", "craft", "One body font throughout", registry=T)
def one_font(ctx):
    d = ctx.facts["defaults"]
    base = d.get("normal_font") or d.get("font")
    odd = sorted({r["font"] for p in paragraphs(ctx) for r in p["runs"] if r["font"] and r["font"] != base})
    return Result.when(not odd, "body font %s; other fonts: %s" % (base or "theme default", odd or "none"))
