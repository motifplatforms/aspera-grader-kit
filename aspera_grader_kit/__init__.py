"""aspera-grader-kit: a small pattern for grading the files knowledge-work agents produce.

Inspect the file into facts, run checks that each test one thing and explain themselves, then score with a rubric
whose weights sit in plain JSON next to the task.
"""
from .checks import Check, Context, Result, check
from .grade import grade, render
from .inspect import inspect

__all__ = ["Check", "Context", "Result", "check", "grade", "render", "inspect"]
__version__ = "0.1.1"
