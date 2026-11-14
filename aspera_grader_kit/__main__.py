"""python -m aspera_grader_kit <task_dir> <output_file> [--json]"""
import argparse, json, sys

from .grade import grade, render


def main(argv=None):
    ap = argparse.ArgumentParser(prog="aspera_grader_kit", description="Grade one output file against one task folder.")
    ap.add_argument("task_dir"); ap.add_argument("output")
    ap.add_argument("--json", action="store_true", help="print the full report as JSON")
    a = ap.parse_args(argv)
    rep = grade(a.output, a.task_dir)
    print(json.dumps(rep, indent=2, default=str) if a.json else render(rep))
    return 0


if __name__ == "__main__":
    sys.exit(main())
