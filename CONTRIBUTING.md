# Contributing

Thanks for looking. This is a small project with an unusually strict set of internal rules, and
they exist because each one was learned the hard way.

**Read [AGENTS.md](AGENTS.md) first.** It is the contributor contract: the commands, the repo map,
and 22 invariants that must not be broken. Each links to the decision in
[docs/DECISIONS.md](docs/DECISIONS.md) that explains what went wrong when it was ignored.

## Getting set up

You need **Node.js 24 or newer**: the build scripts are TypeScript run directly by `node`.

```bash
npm install
npm run desktop        # builds and opens the window
npm run trust-cert     # once; macOS will ask for your password
```

Then insert the add-in in Excel: **Insert → Add-ins → My Add-ins → AI Spreadsheet Automation**. You
need a DeepSeek API key to exercise the model, but **not** to run the tests.

## Before you open a pull request

```bash
npm test                 # 67 checks, offline: no Excel, no network, no API key
npm run typecheck:ui     # the React app is typechecked separately
```

Both must pass. `npm test` builds first, so it also catches compile errors.

If you changed `app/web/manifest.template.xml`, also run:

```bash
npm run validate:manifest
```

It is not part of `npm test` because it downloads a tool and talks to Microsoft — but it once found
two errors in a manifest Excel had been loading happily for weeks.

## The three things most likely to trip you up

1. **Restore Excel snapshots from `formulas`, never `values`.** Assigning `values` to a cell that
   held a formula writes the number and destroys the formula.
2. **Snapshot before every mutating write, and refuse the write if the snapshot fails.** A change
   that cannot be undone is worse than one that is refused.
3. **The certificate exception is loopback-and-own-port only, and fails closed.** A mistake there
   means trusting any certificate.

## Adding a new Excel tool

A tool is not done until it can be undone. The steps are:

1. A schema in `app/tools.ts` — the schema **is** the runtime validator, so describe arguments the
   way you want them rejected.
2. An executor in `app/web/executor.ts`, under the same name.
3. An undo plan. If the change is a range of cells and stays inside the address it was given, the
   existing range snapshot covers it — add the tool and a test. If it creates an object or escapes
   the address, it needs a new `UndoAction` kind, and the journal's format is on disk, so add a kind
   rather than reshaping an existing one.
4. A test in `app/selftest.ts` proving the snapshot happens *before* the write and that undo
   restores the state.
5. If it is new Office.js behaviour, a check in `app/web/diagnostics.ts`. **You cannot verify
   Office.js without Excel** — say so in the pull request and ask for a Diagnostics run rather than
   claiming it works.

## Tests must never touch the live setup

The suite runs on a real machine, so four guards are structural and must stay:

- Test servers use port **3499**, never 3443 — the taskpane polls 3443, so a test server there is
  reachable by a real workbook.
- Every test gets its own journal file in a temp directory. The default is the user's real
  `~/.ai-excel/journal.json`, which is how they undo.
- Certificate and manifest work uses temp directories with trusting off
  (`AI_EXCEL_CERT_DIR`, `AI_EXCEL_WEF_DIR`, `AI_EXCEL_NO_KEYCHAIN=1`).
- `--smoke` never uses a real API key.

## Style

- **Comment the *why*.** Several places here look wrong until you know the reason. Those comments
  are the point, not decoration.
- TypeScript everywhere, `strict` on. The server and the taskpane have **zero runtime
  dependencies**; adding one is a decision worth recording, not a default.
- Errors are written for the model to act on: Excel's own messages name nothing useful, so name the
  sizes and the offending argument instead.

## Licence

By contributing you agree your work is licensed under the [MIT licence](LICENSE).
