"""Rubric: turn check results into a score out of 100.

A rubric is plain JSON, kept next to the task so the weights are visible and versioned:

    {
      "version": "example-1",
      "categories": {
        "requirements": {"label": "Task requirements", "points": 60, "per_item": true},
        "craft":        {"label": "Craft",             "points": 40}
      },
      "gates": {"missing": "zero", "unreadable": "zero"}
    }

Rules, deliberately simple so a reader can predict a score by hand:
- Each category earns points x (share of its units passed).
- per_item: every element of a check counts as one unit (ten required values weigh ten times, not once).
  Otherwise each check counts once, scaled by its weight.
- A category with no applicable checks hands its points to the others, so a task is never scored on what it
  did not ask for.
- Gates override the sum: "zero" sets the total to 0; {"cap": N} caps it at N.
These are example mechanics for the kit; tune them for your own tasks.
"""
import json


def load(path):
    with open(path) as f:
        return json.load(f)


def score(results, checks, rubric, gates=()):
    cats = {k: {"label": v.get("label", k), "points": float(v["points"]), "per_item": bool(v.get("per_item")),
                "num": 0.0, "den": 0.0, "checks": 0}
            for k, v in rubric["categories"].items()}
    for c in checks:
        r = results.get(c.id)
        u = r.units() if r is not None else None
        if u is None or c.category not in cats: continue
        p, n = u; x = cats[c.category]; x["checks"] += 1
        if x["per_item"]:
            x["num"] += p; x["den"] += n
        else:
            x["num"] += c.weight * p / n; x["den"] += c.weight
    live = {k: x for k, x in cats.items() if x["den"] > 0}
    total_points = sum(x["points"] for x in live.values()) or 1.0
    out_cats, total = {}, 0.0
    for k, x in cats.items():
        possible = x["points"] * 100.0 / total_points if k in live else 0.0
        share = x["num"] / x["den"] if k in live else None
        earned = possible * share if share is not None else 0.0
        total += earned
        out_cats[k] = {"label": x["label"], "possible": round(possible, 2), "earned": round(earned, 2),
                       "share": None if share is None else round(share, 3), "checks": x["checks"]}
    applied = []
    for g in gates:
        rule = rubric.get("gates", {}).get(g["id"])
        if rule == "zero":
            total = 0.0; applied.append(g)
        elif isinstance(rule, dict) and "cap" in rule:
            total = min(total, float(rule["cap"])); applied.append(g)
    return {"total": round(total, 1), "categories": out_cats, "gates": applied, "rubric_version": rubric.get("version")}
