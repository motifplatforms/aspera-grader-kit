"""Read a .docx into plain facts a check can test: body blocks in order (paragraphs with style,
numbering and run formatting; tables with header-repeat, shading, cell margins, alignment), tracked changes, comments,
content controls, protection, fields (PAGE, NUMPAGES, TOC, SEQ), headers/footers, footnotes, images, hyperlinks."""
import re, zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W, "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main", "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
q = lambda t: "{%s}%s" % (NS[t.split(":")[0]], t.split(":")[1])
wv = lambda el, path: (el.find(path, NS).get(q("w:val")) if el is not None and el.find(path, NS) is not None else None)


def text_of(el, deleted=False):
    out = []
    for t in el.iter(q("w:t"), q("w:delText"), q("w:tab"), q("w:br")):
        if t.tag == q("w:t"): out.append(t.text or "")
        elif t.tag == q("w:delText") and deleted: out.append(t.text or "")
        elif t.tag == q("w:tab"): out.append("\t")
    return "".join(out)


def styles(z):
    if "word/styles.xml" not in z.namelist(): return {}
    root = etree.fromstring(z.read("word/styles.xml"))
    return {s.get(q("w:styleId")): (s.find("w:name", NS).get(q("w:val")) if s.find("w:name", NS) is not None else s.get(q("w:styleId")))
            for s in root.iter(q("w:style"))}


def fields_in(el):
    """Field codes (simple and complex) in an element, upper-cased first word: PAGE, NUMPAGES, TOC, SEQ, REF ..."""
    codes = [f.get(q("w:instr")) or "" for f in el.iter(q("w:fldSimple"))]
    codes += [(t.text or "") for t in el.iter(q("w:instrText"))]
    return [c.strip() for c in codes if c and c.strip()]


def para_facts(p, S):
    ppr = p.find("w:pPr", NS)
    sid = wv(ppr, "w:pStyle") if ppr is not None else None
    num = ppr.find("w:numPr", NS) if ppr is not None else None
    runs = []
    for r in p.iter(q("w:r")):
        rpr = r.find("w:rPr", NS); t = text_of(r)
        if not t.strip(): continue
        b = rpr is not None and rpr.find("w:b", NS) is not None and (rpr.find("w:b", NS).get(q("w:val")) not in ("0", "false"))
        sz = wv(rpr, "w:sz"); col = wv(rpr, "w:color"); font = rpr.find("w:rFonts", NS).get(q("w:ascii")) if rpr is not None and rpr.find("w:rFonts", NS) is not None else None
        sup = wv(rpr, "w:vertAlign") == "superscript" if rpr is not None else False
        runs.append({"t": t[:200], "b": b, "sz": int(sz) / 2 if sz and sz.isdigit() else None, "color": col, "font": font, "sup": sup})
    return {"type": "p", "style": S.get(sid, sid) if sid else "Normal", "text": text_of(p)[:1200],
            "num": {"numId": wv(num, "w:numId"), "ilvl": wv(num, "w:ilvl")} if num is not None else None,
            "align": wv(ppr, "w:jc") if ppr is not None else None, "runs": runs[:40],
            "fields": fields_in(p), "has_drawing": p.find(".//w:drawing", NS) is not None or p.find(".//w:pict", NS) is not None,
            "in_del": text_of(p, deleted=True) != text_of(p)}


def table_facts(t, S):
    rows = []
    for tr in t.findall("w:tr", NS):
        trpr = tr.find("w:trPr", NS)
        cells = []
        for tc in tr.findall("w:tc", NS):
            tcpr = tc.find("w:tcPr", NS)
            shd = tcpr.find("w:shd", NS) if tcpr is not None else None
            mar = tcpr.find("w:tcMar", NS) if tcpr is not None else None
            ps = tc.findall("w:p", NS)
            bold = any(r.find("w:rPr/w:b", NS) is not None for p in ps for r in p.iter(q("w:r")) if text_of(r).strip())
            colors = sorted({wv(r.find("w:rPr", NS), "w:color") for p in ps for r in p.iter(q("w:r")) if text_of(r).strip() and r.find("w:rPr", NS) is not None} - {None})
            cells.append({"text": " ".join(text_of(p) for p in ps).strip()[:200], "fill": shd.get(q("w:fill")) if shd is not None else None,
                          "align": next((wv(p.find("w:pPr", NS), "w:jc") for p in ps if p.find("w:pPr/w:jc", NS) is not None), None),
                          "margins": {e.tag.split("}")[1]: e.get(q("w:w")) for e in mar} if mar is not None else None, "bold": bold, "colors": colors})
        spans = sum(int(wv(tc.find("w:tcPr", NS), "w:gridSpan") or 1) if tc.find("w:tcPr", NS) is not None else 1 for tc in tr.findall("w:tc", NS))
        sizes = [int(sz.get(q("w:val"))) / 2 for sz in tr.iter(q("w:sz")) if sz.get(q("w:val"), "").isdigit() and sz.getparent().tag == q("w:rPr")]
        rows.append({"header": trpr is not None and trpr.find("w:tblHeader", NS) is not None, "cells": cells, "span": spans, "min_pt": min(sizes) if sizes else None})
    tpr = t.find("w:tblPr", NS)
    tmar = tpr.find("w:tblCellMar", NS) if tpr is not None else None
    grid = [int(g.get(q("w:w")) or 0) for g in t.findall("w:tblGrid/w:gridCol", NS)]
    return {"type": "table", "grid_w_twips": sum(grid), "style": S.get(wv(tpr, "w:tblStyle"), wv(tpr, "w:tblStyle")) if tpr is not None else None,
            "n_rows": len(rows), "n_cols": max((len(r["cells"]) for r in rows), default=0),
            "table_cell_margins": {e.tag.split("}")[1]: e.get(q("w:w")) for e in tmar} if tmar is not None else None,
            "rows": rows[:60]}


def inspect(path):
    z = zipfile.ZipFile(path); names = z.namelist(); S = styles(z)
    doc = etree.fromstring(z.read("word/document.xml")); body = doc.find("w:body", NS)
    blocks = []
    for el in body:
        if el.tag == q("w:p"): blocks.append(para_facts(el, S))
        elif el.tag == q("w:tbl"): blocks.append(table_facts(el, S))
        elif el.tag == q("w:sdt"):
            for p in el.iter(q("w:p")): blocks.append(dict(para_facts(p, S), sdt_block=True))
    ins = [{"author": e.get(q("w:author")), "text": text_of(e)[:300]} for e in doc.iter(q("w:ins"))]
    dels = [{"author": e.get(q("w:author")), "text": "".join((t.text or "") for t in e.iter(q("w:delText")))[:300]} for e in doc.iter(q("w:del"))]
    comments = []
    if "word/comments.xml" in names:
        for c in etree.fromstring(z.read("word/comments.xml")).iter(q("w:comment")):
            comments.append({"id": c.get(q("w:id")), "author": c.get(q("w:author")), "text": text_of(c)[:500]})
    anchors = {}
    for p in doc.iter(q("w:p")):
        for s in p.iter(q("w:commentRangeStart"), q("w:commentReference")):
            anchors.setdefault(s.get(q("w:id")), text_of(p)[:160])
    for c in comments: c["anchored_in"] = anchors.get(c["id"])
    sdts = []
    for s in doc.iter(q("w:sdt")):
        pr = s.find("w:sdtPr", NS)
        if pr is None: continue
        kind = ("dropdown" if pr.find("w:dropDownList", NS) is not None else "combobox" if pr.find("w:comboBox", NS) is not None
                else "date" if pr.find("w:date", NS) is not None else "checkbox" if pr.find("w14:checkbox", NS) is not None
                else "richtext" if pr.find("w:richText", NS) is not None else "picture" if pr.find("w:picture", NS) is not None
                else "text" if pr.find("w:text", NS) is not None else "docpart" if pr.find("w:docPartObj", NS) is not None else "other")
        items = [li.get(q("w:displayText")) or li.get(q("w:value")) for li in pr.iter(q("w:listItem"))]
        df = pr.find("w:date/w:dateFormat", NS)
        sdts.append({"kind": kind, "tag": wv(pr, "w:tag"), "alias": wv(pr, "w:alias"), "placeholder": pr.find("w:showingPlcHdr", NS) is not None or pr.find("w:placeholder", NS) is not None,
                     "items": items, "date_format": df.get(q("w:val")) if df is not None else None,
                     "text": text_of(s.find("w:sdtContent", NS))[:120] if s.find("w:sdtContent", NS) is not None else ""})
    prot = None
    if "word/settings.xml" in names:
        dp = etree.fromstring(z.read("word/settings.xml")).find("w:documentProtection", NS)
        if dp is not None: prot = {"edit": dp.get(q("w:edit")), "enforcement": dp.get(q("w:enforcement")), "password": any("hash" in k.lower() or "salt" in k.lower() for k in dp.attrib)}
    hf = {}
    for n in names:
        m = re.match(r"word/(header|footer)\d+\.xml$", n)
        if m:
            root = etree.fromstring(z.read(n)); hf.setdefault(m.group(1), []).append({"text": text_of(root)[:300], "fields": fields_in(root)})
    sect = body.find("w:sectPr", NS)
    pg = {}
    if sect is not None:
        ps, pm = sect.find("w:pgSz", NS), sect.find("w:pgMar", NS)
        if ps is not None: pg = {"w_twips": int(ps.get(q("w:w")) or 0), "h_twips": int(ps.get(q("w:h")) or 0)}
        if pm is not None: pg.update({k: int(pm.get(q("w:" + k)) or 0) for k in ("left", "right", "top", "bottom")})
    foot = []
    if "word/footnotes.xml" in names:
        for f in etree.fromstring(z.read("word/footnotes.xml")).iter(q("w:footnote")):
            if f.get(q("w:type")) in (None, "normal"): foot.append(text_of(f)[:200])
    links = []
    rels = {}
    if "word/_rels/document.xml.rels" in names:
        for r in etree.fromstring(z.read("word/_rels/document.xml.rels")).iter(REL):
            rels[r.get("Id")] = (r.get("Target"), r.get("TargetMode"))
    for h in doc.iter(q("w:hyperlink")):
        rid = h.get(q("r:id")); links.append({"text": text_of(h)[:120], "target": rels.get(rid, (None,))[0] if rid else None, "anchor": h.get(q("w:anchor"))})
    fields = fields_in(doc)
    images = sum(1 for _ in doc.iter(q("w:drawing"))) + sum(1 for _ in doc.iter(q("w:pict")))
    dd = {}
    if "word/styles.xml" in names:
        sr = etree.fromstring(z.read("word/styles.xml"))
        rp = sr.find("w:docDefaults/w:rPrDefault/w:rPr", NS)
        if rp is not None:
            dd["sz"] = int(wv(rp, "w:sz")) / 2 if wv(rp, "w:sz") and wv(rp, "w:sz").isdigit() else None
            f = rp.find("w:rFonts", NS); dd["font"] = f.get(q("w:ascii")) if f is not None else None
        for st in sr.iter(q("w:style")):
            if st.get(q("w:styleId")) == "Normal":
                rp = st.find("w:rPr", NS)
                if rp is not None and wv(rp, "w:sz"): dd["normal_sz"] = int(wv(rp, "w:sz")) / 2
                if rp is not None and rp.find("w:rFonts", NS) is not None: dd["normal_font"] = rp.find("w:rFonts", NS).get(q("w:ascii"))
    num_xml = z.read("word/numbering.xml").decode("utf-8", "ignore") if "word/numbering.xml" in names else ""
    return {"blocks": blocks, "ins": ins, "del": dels, "comments": comments, "sdts": sdts, "protection": prot, "headers_footers": hf,
            "page": pg, "footnotes": foot, "hyperlinks": links, "fields": fields, "images": images,
            "has_numbering_part": bool(num_xml), "defaults": dd, "styles_used": sorted({b["style"] for b in blocks if b["type"] == "p"}),
            "text": "\n".join(b["text"] if b["type"] == "p" else "\n".join(" | ".join(c["text"] for c in r["cells"]) for r in b["rows"]) for b in blocks)}
