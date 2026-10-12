# CLAUDE.md

The agent instructions for this repository live in **[AGENTS.md](AGENTS.md)**. Read that first: it
has the commands, the repo map, the invariants that must not be broken, and how to verify your work.

This file exists only because several tools look for `CLAUDE.md` by name. It is deliberately a
pointer — duplicating the content here would let the two drift apart, and a stale instruction file
is worse than none.

If you read nothing else, these three are the ones that cause silent damage:

1. **Restore Excel snapshots from `formulas`, never `values`.** Assigning `values` to a cell that
   held a formula writes the number and destroys the formula.
2. **Snapshot before every mutating write, and refuse the write if the snapshot fails.**
3. **The certificate exception is loopback-and-own-port only, and must fail closed.**

Two things you cannot verify on your own — say so rather than claiming success:

- **Office.js behaviour** needs real Excel. Ask the user to run the Diagnostics tab.
- **The look of the UI** needs eyes. Use the smoke test's screenshot, and look at the image.

Further reading: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) ·
[docs/DECISIONS.md](docs/DECISIONS.md) · [app/README.md](app/README.md)
