"""Checks: small functions that each test one thing about an output and say why.

A check gets a Context (the inspected facts, the task and its answer key) and returns a Result. A check that tests a
list of things (five key values, three required sections) returns one element per item, so the rubric can count every
item instead of treating the whole list as one pass or fail.

    from aspera_grader_kit import check, Result

    @check("R1", "requirements", "Every key value appears in the memo")
    def key_values(ctx):
        return Result.items([(v["item"], str(v["value"]) in ctx.text, v["value"]) for v in ctx.key["values"]])
"""
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional


@dataclass
class Result:
    status: str                      # pass | fail | na | error
    evidence: str = ""
    elements: Optional[List[dict]] = None

    @classmethod
    def ok(cls, evidence=""): return cls("pass", evidence)

    @classmethod
    def fail(cls, evidence=""): return cls("fail", evidence)

    @classmethod
    def na(cls, evidence="does not apply"): return cls("na", evidence)

    @classmethod
    def when(cls, condition, evidence=""):
        return cls("pass" if condition else "fail", evidence)

    @classmethod
    def items(cls, triples):
        """One element per (item, passed, evidence). Passes only when every item passes."""
        els = [{"item": str(i), "status": "pass" if ok else "fail", "evidence": str(e)} for i, ok, e in triples]
        if not els: return cls.na("no items to test")
        n = sum(e["status"] == "pass" for e in els)
        return cls("pass" if n == len(els) else "fail", "%d of %d items met" % (n, len(els)), els)

    def units(self):
        """(passed, total) units for scoring, or None when the check does not count."""
        if self.status not in ("pass", "fail"): return None
        if self.elements: return sum(e["status"] == "pass" for e in self.elements), len(self.elements)
        return (1, 1) if self.status == "pass" else (0, 1)

    def to_dict(self):
        d = {"status": self.status, "evidence": self.evidence}
        if self.elements is not None: d["elements"] = self.elements
        return d


@dataclass
class Check:
    id: str
    category: str
    description: str
    fn: Callable
    weight: float = 1.0


@dataclass
class Context:
    path: str
    facts: dict
    task: dict = field(default_factory=dict)
    key: dict = field(default_factory=dict)

    @property
    def text(self): return self.facts.get("text", "")


_REGISTRY: Dict[str, List[Check]] = {}


def check(id, category, description, weight=1.0, registry="default"):
    """Decorator that registers a check function under a registry name (one per task, usually)."""
    def wrap(fn):
        _REGISTRY.setdefault(registry, []).append(Check(id, category, description, fn, weight))
        return fn
    return wrap


def registered(registry="default"):
    return list(_REGISTRY.get(registry, []))


def clear(registry="default"):
    _REGISTRY.pop(registry, None)


def run(checks, ctx):
    """Run every check; an exception inside a check becomes status 'error' with the message as evidence."""
    out = {}
    for c in checks:
        try:
            r = c.fn(ctx)
            if not isinstance(r, Result):
                raise TypeError("check %s returned %r, not a Result" % (c.id, type(r).__name__))
        except Exception as e:  # a broken check must never crash the grade
            r = Result("error", "%s: %s" % (type(e).__name__, e))
        out[c.id] = r
    return out
