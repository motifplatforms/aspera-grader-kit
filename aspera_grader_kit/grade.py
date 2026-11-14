"""Grade one output file against one task folder.

A task folder holds:
    task.md       the brief the agent was given
    task.json     {"id", "format", "output"}   (format: docx | pptx | xlsx)
    key.json      the answer key (values, required items); never shown to the agent
    rubric.json   categories, points and gates
    checks.py     the task's checks, registered with @check(..., registry=<task id>)
"""
import importlib.util, json, os

from . import checks as C
from . import rubric as R
from .inspect import inspect


def _load_json(task_dir, name, default=None):
    p = os.path.join(task_dir, name)
    if not os.path.exists(p): return default
    with open(p) as f: return json.load(f)


def load_task_checks(task_dir, task_id):
    C.clear(task_id)
    spec = importlib.util.spec_from_file_location("task_checks_%s" % task_id.replace("-", "_"), os.path.join(task_dir, "checks.py"))
    mod = importlib.util.module_from_spec(spec); spec.loader.exec_module(mod)
    return C.registered(task_id)


def grade(output_path, task_dir):
    task = _load_json(task_dir, "task.json"); key = _load_json(task_dir, "key.json", {})
    rubric = R.load(os.path.join(task_dir, "rubric.json"))
    checks = load_task_checks(task_dir, task["id"])
    report = {"task": task["id"], "output": os.path.basename(output_path) if output_path else None, "checks": {}}
    gates = []
    if not output_path or not os.path.exists(output_path):
        gates.append({"id": "missing", "evidence": "no output file"})
    else:
        try:
            facts = inspect(output_path)
        except Exception as e:
            gates.append({"id": "unreadable", "evidence": "%s: %s" % (type(e).__name__, e)})
        else:
            results = C.run(checks, C.Context(output_path, facts, task, key))
            report["checks"] = {c.id: dict(results[c.id].to_dict(), category=c.category, description=c.description) for c in checks}
            report.update(R.score(results, checks, rubric, gates))
            return report
    report.update(R.score({}, checks, rubric, gates))   # nothing to check: every category is empty, so the total is 0
    report["gates"] = gates
    return report


def render(report):
    """Plain-text summary of a report for the terminal."""
    lines = ["%s  %s  ->  %.1f / 100" % (report["task"], report.get("output"), report["total"])]
    for g in report.get("gates", []): lines.append("  GATE %s: %s" % (g["id"], g["evidence"]))
    for k, c in report.get("categories", {}).items():
        if c["possible"]: lines.append("  %-20s %5.1f / %5.1f" % (c["label"], c["earned"], c["possible"]))
    for cid, r in report.get("checks", {}).items():
        lines.append("  [%s] %-4s %s  (%s)" % (r["status"].upper()[:4].ljust(4), cid, r["description"], r["evidence"]))
        for e in r.get("elements") or []:
            if e["status"] != "pass": lines.append("           - missed: %s (%s)" % (e["item"], e["evidence"]))
    return "\n".join(lines)
