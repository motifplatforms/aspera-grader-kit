"""Read a .pptx into plain facts a check can test: slide size, slides in presentation order, titles, text, fonts and
sizes (body vs footer zone), charts (and whether each keeps an embedded workbook), tables, pictures, off-canvas shapes,
source lines and speaker notes. Reads the XML directly, so it needs no PowerPoint and never renders anything."""
import hashlib, os, re, zipfile
from lxml import etree

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
EMU_IN = 914400
NUM = re.compile(r"-?\(?\$?\d[\d,]*\.?\d*\)?%?")
FOOTER_TEXT = re.compile(r"^\s*(sources?|notes?|footnotes?|confidential|©|\*|\d+\s*(/|of)\s*\d+\s*$)", re.I)


def _top_y(el):
    """y offset of the outermost group holding el (child offsets inside a group are in group space)."""
    y, cur = None, el
    while cur is not None:
        if cur.tag == "{%s}grpSp" % NS["p"]:
            off = cur.find("./p:grpSpPr/a:xfrm/a:off", NS)
        elif cur.tag == "{%s}graphicFrame" % NS["p"]:
            off = cur.find("./p:xfrm/a:off", NS)
        elif cur.tag == "{%s}sp" % NS["p"]:
            off = cur.find("./p:spPr/a:xfrm/a:off", NS)
        else:
            off = None
        if off is not None: y = int(off.get("y"))
        cur = cur.getparent()
    return y


def font_roles(root, slide_h):
    """Split run sizes into body and footer. Footer = footer/date/slide-number placeholder, a shape in the bottom 15% of
    the slide, or text that starts like a source line, a note or a page number."""
    body, foot = [], []
    for shp in list(root.iter("{%s}sp" % NS["p"])) + list(root.iter("{%s}graphicFrame" % NS["p"])):
        sizes = [int(r.get("sz")) / 100 for r in shp.iter("{%s}rPr" % NS["a"]) if r.get("sz")]
        if not sizes: continue
        ph = shp.find(".//p:nvPr/p:ph", NS)
        y = _top_y(shp)
        is_foot = (ph is not None and ph.get("type") in ("ftr", "sldNum", "dt")) or (y is not None and y >= 0.85 * slide_h) \
            or bool(FOOTER_TEXT.match("".join(shp.itertext()).strip()))
        (foot if is_foot else body).extend(sizes)
    return body, foot


def slide_parts(z):
    """Slide parts in presentation order (sldIdLst). File names keep their numbers when slides are reordered, so name
    order is only a fallback."""
    s = [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    by_name = sorted(s, key=lambda n: int(re.findall(r"\d+", n)[-1]))
    try:
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
        rl = {r.get("Id"): r.get("Target") for r in etree.fromstring(z.read("ppt/_rels/presentation.xml.rels")).iter(REL)}
        order = []
        for sid in pres.iter("{%s}sldId" % NS["p"]):
            t = rl.get(sid.get("{%s}id" % NS["r"]))
            if t: order.append(os.path.normpath(os.path.join("ppt", t)).replace("\\", "/"))
        order = [o for o in order if o in s]
        return order if len(order) == len(s) else by_name
    except Exception:
        return by_name


def rels(z, part):
    d, f = os.path.split(part); rp = d + "/_rels/" + f + ".rels"
    if rp not in z.namelist(): return {}
    out = {}
    for r in etree.fromstring(z.read(rp)).iter(REL):
        if r.get("TargetMode") == "External": continue
        out[r.get("Id")] = (os.path.normpath(os.path.join(d, r.get("Target"))).replace("\\", "/"), r.get("Type", "").rsplit("/", 1)[-1])
    return out


def numbers_in(text):
    """Numbers written in text: 1,250 / (3.5) / -4% / $12.0 -> floats; brackets and minus signs mean negative."""
    vals = []
    for m in NUM.findall(text):
        neg = m.startswith("-") or (m.startswith("(") and m.endswith(")"))
        s = re.sub(r"[^\d.]", "", m)
        if not s or s == ".": continue
        try: v = float(s)
        except ValueError: continue
        vals.append(-v if neg else v)
    return vals


def inspect(path):
    z = zipfile.ZipFile(path)
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    sz = pres.find("p:sldSz", NS); W, H = int(sz.get("cx")), int(sz.get("cy"))
    slides, all_text, all_nums, hashes = [], [], [], []
    for i, sp in enumerate(slide_parts(z), 1):
        xml = z.read(sp); root = etree.fromstring(xml); rl = rels(z, sp)
        hashes.append(hashlib.sha1(xml).hexdigest())
        texts = ["".join(t.itertext()) for t in root.iter("{%s}p" % NS["a"])]
        text = "\n".join(t for t in texts if t.strip())
        charts = embedded = pics = big_pics = tables = offcanvas = 0
        chart_nums = []
        for gf in root.iter("{%s}graphicFrame" % NS["p"]):
            ch = gf.find(".//c:chart", NS)
            if ch is not None:
                charts += 1
                tgt = rl.get(ch.get("{%s}id" % NS["r"]), (None,))[0]
                if tgt and tgt in z.namelist():
                    if any(t in ("package", "oleObject") for _, t in rels(z, tgt).values()): embedded += 1
                    cx = etree.fromstring(z.read(tgt))
                    chart_nums += [float(v.text) for v in cx.iter("{%s}v" % NS["c"]) if re.match(r"^-?\d+\.?\d*(E-?\d+)?$", v.text or "")]
            if gf.find(".//a:tbl", NS) is not None: tables += 1
        for pic in root.iter("{%s}pic" % NS["p"]):
            pics += 1
            ext = pic.find(".//a:xfrm/a:ext", NS)
            if ext is not None and int(ext.get("cx")) * int(ext.get("cy")) > 0.08 * W * H: big_pics += 1
        for off in root.iter("{%s}off" % NS["a"]):
            x, y = int(off.get("x")), int(off.get("y"))
            if x < -EMU_IN * 0.05 or y < -EMU_IN * 0.05 or x > W or y > H: offcanvas += 1
        sizes = [int(r.get("sz")) / 100 for r in root.iter("{%s}rPr" % NS["a"]) if r.get("sz")]
        body_sz, foot_sz = font_roles(root, H)
        slides.append({"slide": i, "title": texts[0][:120] if texts else "", "text": text[:2000],
                       "charts": charts, "charts_with_workbook": embedded, "tables": tables,
                       "pictures": pics, "large_pictures": big_pics, "offcanvas_shapes": offcanvas,
                       "fonts": sorted({e.get("typeface") for e in root.iter("{%s}latin" % NS["a"]) if e.get("typeface")}),
                       "min_font_pt": min(sizes) if sizes else None,
                       "min_body_font_pt": min(body_sz) if body_sz else None,
                       "min_footer_font_pt": min(foot_sz) if foot_sz else None,
                       "has_source_line": bool(re.search(r"\bsources?\b|\bsource:", text, re.I)),
                       "has_notes": any(t == "notesSlide" for _, t in rl.values())})
        all_text.append(text); all_nums += numbers_in(text) + chart_nums
    return {"slide_w_in": round(W / EMU_IN, 2), "slide_h_in": round(H / EMU_IN, 2), "slides": slides,
            "text": "\n".join(all_text), "numbers": all_nums, "slide_hashes": hashes}
