"""Read an .xlsx into plain facts a check can test: sheets, cells (formula, cached value,
number format, input styling), errors, hard-coded constants in formulas, charts and their ranges, names, validation,
comments, freeze panes, merged cells, external links; plus a LibreOffice recalculation to compare cached values."""
import io, os, re, shutil, subprocess, tempfile, zipfile
import openpyxl
from openpyxl.utils import get_column_letter

ERR = re.compile(r"^#(REF!|DIV/0!|VALUE!|NAME\?|N/A|NUM!|NULL!|SPILL!|CALC!)")
CONST = re.compile(r"(?<![A-Za-z_$:\d.])(-?\d+(?:\.\d+)?)(?![\d.]*[:A-Za-z(])")
BENIGN = {"0", "1", "2", "3", "4", "5", "6", "7", "12", "100", "1000", "-1", "0.5", "365", "4"}


def recalc(path):
    """LibreOffice opens the workbook, recalculates and saves a copy; returns {sheet: {cell: value}} or None."""
    tmp = tempfile.mkdtemp(prefix="xr-")
    try:
        src = os.path.join(tmp, "in.xlsx"); shutil.copy(path, src); out = os.path.join(tmp, "out")
        os.makedirs(out)
        subprocess.run(["soffice", "--headless", "--norestore", "-env:UserInstallation=file://" + os.path.join(tmp, "p"),
                        "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, src], capture_output=True, timeout=600)
        f = os.path.join(out, "in.xlsx")
        if not os.path.exists(f): return None
        wb = openpyxl.load_workbook(f, data_only=True)
        return {ws.title: {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None} for ws in wb.worksheets}
    except Exception:
        return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def inspect(path, with_recalc=True):
    wbf = openpyxl.load_workbook(path, data_only=False)
    wbv = openpyxl.load_workbook(path, data_only=True)
    z = zipfile.ZipFile(path)
    sheets = []
    for ws in wbf.worksheets:
        wv = wbv[ws.title]
        cells, n_formula, n_num, errs, consts, inputs_styled = [], 0, 0, [], [], 0
        for row in ws.iter_rows():
            for c in row:
                if c.value is None: continue
                v = wv[c.coordinate].value
                isf = isinstance(c.value, str) and c.value.startswith("=")
                if isf:
                    n_formula += 1
                    k = [x for x in CONST.findall(re.sub(r"\"[^\"]*\"", "", c.value[1:])) if x not in BENIGN]
                    if k: consts.append({"cell": c.coordinate, "formula": c.value[:120], "constants": k[:4]})
                elif isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                    n_num += 1
                    fill = c.fill.fgColor.rgb if c.fill and c.fill.fill_type == "solid" and c.fill.fgColor is not None else None
                    fc = c.font.color.rgb if c.font and c.font.color is not None and c.font.color.type == "rgb" else None
                    if (fill and fill not in ("00000000", "FFFFFFFF")) or (fc and fc not in ("FF000000", None)): inputs_styled += 1
                if isinstance(v, str) and ERR.match(v): errs.append(c.coordinate)
                if len(cells) < 400:
                    cells.append({"c": c.coordinate, "f": c.value[:160] if isf else None, "v": v if not isinstance(v, str) else v[:120],
                                  "fmt": c.number_format if c.number_format != "General" else None})
        charts = []
        for ch in getattr(ws, "_charts", []):
            refs = []
            try:
                for s in ch.series:
                    if s.val and s.val.numRef: refs.append(s.val.numRef.f)
            except Exception: pass
            def _axt(ax):
                try: return bool(ax is not None and ax.title is not None)
                except Exception: return False
            pie = "Pie" in type(ch).__name__ or "Doughnut" in type(ch).__name__
            charts.append({"type": type(ch).__name__, "title": _title(ch), "refs": refs,
                           "axis_titles": None if pie else (_axt(getattr(ch, "x_axis", None)) and _axt(getattr(ch, "y_axis", None)))})
        dv = [{"sqref": str(d.sqref), "type": d.type, "formula": d.formula1} for d in ws.data_validations.dataValidation] if ws.data_validations else []
        comments = [{"cell": c.coordinate, "text": c.comment.text[:200]} for row in ws.iter_rows() for c in row if c.comment]
        sheets.append({"name": ws.title, "dims": ws.dimensions, "max_row": ws.max_row, "max_col": ws.max_column, "formulas": n_formula,
                       "numbers": n_num, "inputs_styled": inputs_styled, "errors": errs[:50], "n_errors": len(errs),
                       "constants_in_formulas": consts[:40], "n_constants_in_formulas": len(consts), "charts": charts,
                       "freeze": ws.freeze_panes, "merged": [str(m) for m in ws.merged_cells.ranges][:30], "validation": dv[:20],
                       "comments": comments[:30], "cells": cells, "col_widths": {k: v.width for k, v in list(ws.column_dimensions.items())[:40] if v.width}})
    ext = [n for n in z.namelist() if n.startswith("xl/externalLinks/")]
    names = []
    try: names = [{"name": n, "ref": str(d.attr_text)[:120]} for n, d in wbf.defined_names.items()]
    except Exception: pass
    out = {"sheets": sheets, "sheet_names": [s["name"] for s in sheets], "external_links": len(ext), "defined_names": names}
    rc = recalc(path) if with_recalc else None
    # _grid: every non-empty value (recalculated when available), for value checks against an answer key
    out["_grid"] = {ws.title: {c.coordinate: (rc.get(ws.title, {}).get(c.coordinate) if rc is not None and isinstance(c.value, str) and c.value.startswith("=") else wbv[ws.title][c.coordinate].value)
                               for row in ws.iter_rows() for c in row if c.value is not None} for ws in wbf.worksheets}
    if with_recalc:
        out["recalc_ok"] = rc is not None
        if rc is not None:
            diffs, rc_err = [], []
            for ws in wbf.worksheets:
                wv = wbv[ws.title]; R = rc.get(ws.title, {})
                for row in ws.iter_rows():
                    for c in row:
                        if not (isinstance(c.value, str) and c.value.startswith("=")): continue
                        a, b = wv[c.coordinate].value, R.get(c.coordinate)
                        if isinstance(b, str) and ERR.match(b): rc_err.append("%s!%s" % (ws.title, c.coordinate))
                        if isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) > max(1e-6, abs(b) * 1e-6):
                            diffs.append({"cell": "%s!%s" % (ws.title, c.coordinate), "cached": a, "recalc": b})
                        elif a is None and b is not None: diffs.append({"cell": "%s!%s" % (ws.title, c.coordinate), "cached": None, "recalc": b if not isinstance(b, str) else b[:40]})
            out["recalc_diffs"] = diffs[:40]; out["n_recalc_diffs"] = len(diffs); out["recalc_errors"] = rc_err[:40]
    return out


def _title(ch):
    try:
        t = ch.title
        if t is None: return None
        return " ".join(r.t for p in t.tx.rich.p for r in (p.r or []))[:100]
    except Exception:
        return None
