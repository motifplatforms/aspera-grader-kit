# aspera-grader-kit

[![tests](https://github.com/motifplatforms/aspera-grader-kit/actions/workflows/tests.yml/badge.svg)](https://github.com/motifplatforms/aspera-grader-kit/actions/workflows/tests.yml) [![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE) [![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](pyproject.toml)

A small, readable pattern for grading the files that knowledge-work agents produce: Word memos, PowerPoint decks and
Excel models.

Most agent evals grade the chat transcript. The client never sees the transcript. They open the file. A memo with
its headings faked in bold, a deck whose chart is a pasted picture, or a model with hard-coded numbers inside its
formulas all *look* correct. They are still wrong, and a grader that only reads text will reward them.

This kit grades the file itself:

```
output file ──► inspect ──► facts ──► checks ──► results ──► rubric ──► score + reasons
               (docx/pptx/xlsx)        (one thing each,       (weights in plain JSON,
                                        says why)              next to the task)
```

## Quick start

```bash
pip install -e ".[test]"
python examples/board-memo/make_samples.py
python -m aspera_grader_kit examples/board-memo examples/board-memo/samples/board_memo_weak.docx
```

```
board-memo  board_memo_weak.docx  ->  26.4 / 100
  Task requirements     16.4 /  60.0
  Craft                 10.0 /  40.0
  [FAIL] R2   Required sections exist as real headings  (0 of 2 items met)
           - missed: Summary (no heading called Summary)
  [FAIL] R4   Metrics table has every metric with its value and change  (2 of 4 items met)
           - missed: Operating expenses (no row)
           - missed: EBITDA (no row)
  [FAIL] C2   No faked headings (bold Normal text standing in for a heading)  (fake headings: ['Summary', 'Risks'])
  ...
```

The strong sample scores 100. Every point lost comes with the check that took it and the evidence, so a disagreement
with the grader is a conversation about one line, not about a number.

## What's inside

| Part | What it does |
|---|---|
| `inspect/docx.py` | Body blocks in order (style, numbering, run formatting), tables (header rows, shading, margins), tracked changes, comments, content controls, fields, headers/footers, footnotes, hyperlinks |
| `inspect/pptx.py` | Slides in presentation order, titles, text and numbers, chart count and whether each keeps its workbook, tables, pictures, off-canvas shapes, body vs footer font floors, source lines, notes |
| `inspect/xlsx.py` | Formulas vs typed numbers, hard-coded constants inside formulas, error cells, charts and their ranges, validation, names, external links, and a LibreOffice recalculation to catch stale cached values |
| `checks.py` | `@check` decorator, `Result.items()` for per-item checks, runner that turns a crashing check into an `error` result instead of a crashed grade |
| `rubric.py` | Scores from a JSON rubric: categories, per-item counting, empty categories hand their points on, gates (`zero`, `cap`) |
| `examples/board-memo` | A complete task: brief, input CSV, answer key, rubric, nine checks and a strong/weak pair of outputs |

## Design rules

1. **Grade the artifact, not the transcript.** Inspect the file the way the client will open it.
2. **One check, one variable.** A check that tests three things can't tell you which one failed.
3. **Every check explains itself.** Evidence is part of the result, not a log line.
4. **Count requirements item by item.** A task with ten required values should weigh them ten times, not once.
5. **The answer key never reaches the agent.** Keep `key.json` out of anything the agent can read.
6. **Gates for the unforgivable.** A missing or unopenable file scores zero, however good the reasoning was.
7. **Weights live in data.** The rubric is JSON in the task folder, versioned with it, readable by a reviewer.

## Writing your own task

Copy `examples/board-memo`, then edit:

- `task.md`: the brief the agent sees.
- `key.json`: what a correct answer contains.
- `rubric.json`: categories, points, gates.
- `checks.py`: checks registered with `@check("ID", "category", "description", registry="<task id>")`.

## Where this comes from

This is the open, generic core of the harness we use at [Motif Platforms](https://www.motifplatforms.com) to grade AI
agents on real business deliverables. Our production tasks, answer keys, weights and judge prompts stay private. The
reasoning behind the pattern is in the series
[Teaching machines to do real work](https://www.motifplatforms.com/insights):

- [Reward hacking explained: how one bad grader teaches AI a thousand bad lessons](https://www.motifplatforms.com/insights/reward-hacking-explained-one-bad-grader)
- [How to design training exercises for AI agents](https://www.motifplatforms.com/insights/how-to-design-training-exercises-for-ai-agents)
- [Ten ways a grader lies in AI post-training](https://www.motifplatforms.com/insights/ten-ways-an-ai-grader-lies)
- [The correct-looking file: why AI agents fail and what experts offer](https://www.motifplatforms.com/insights/why-ai-agents-fail-at-spreadsheets-decks-documents)

Benchmark results from the full harness: [motifplatforms/mosaic-benchmarks](https://github.com/motifplatforms/mosaic-benchmarks).

Need graders or exercises built for your own agents? [dsaran@motifplatforms.com](mailto:dsaran@motifplatforms.com)

## License

MIT
