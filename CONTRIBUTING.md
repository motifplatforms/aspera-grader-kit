# Contributing

Thanks for your interest. The kit is small on purpose, so the most useful contributions are:

- **New inspector facts** that a check needs and the inspectors don't expose yet (for example chart axis labels, slide masters, Excel tables).
- **Example tasks** with a brief, an answer key, a rubric, checks and a strong/weak pair of outputs. Keep them fictional: no client data.
- **Bug reports** with the smallest file that reproduces the problem.

## Before you open a pull request

1. `pip install -e ".[test]"`
2. `pytest -q` passes.
3. Each check tests one thing and returns evidence a reviewer can read.
4. No answer keys, real client files or secrets in the change.

## Questions

Open an issue, or write to dsaran@motifplatforms.com.
