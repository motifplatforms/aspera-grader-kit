"""File inspectors: turn an Office file into a plain dict of facts. Checks test facts, never raw XML."""
import os

from . import docx, pptx, xlsx

INSPECTORS = {".docx": docx.inspect, ".pptx": pptx.inspect, ".xlsx": xlsx.inspect}


def inspect(path, **kw):
    """Facts for one file, picked by extension. Raises ValueError for a format with no inspector."""
    ext = os.path.splitext(path)[1].lower()
    if ext not in INSPECTORS:
        raise ValueError("no inspector for %s files" % ext)
    return INSPECTORS[ext](path, **kw)
