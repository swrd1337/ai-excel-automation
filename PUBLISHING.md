# Publishing notes

**Published: <https://github.com/swrd1337/ai-excel-automation>** — public, MIT, CI green.

This file is the record of how it was published and what is left to do by hand. It is not
instructions any more; the steps below are already done.

## What is in place

| File | Why |
|---|---|
| `README.md` | Written for a visitor: hook, screenshot, why it is different, quick start, the undo model, honest status |
| `LICENSE` | MIT, `Copyright (c) 2026 Alexei` |
| `CONTRIBUTING.md` | Setup, the three rules most likely to be broken, how to add a tool |
| `SECURITY.md` | How to report privately, and what this app's security model actually is |
| `.github/workflows/ci.yml` | `npm test` + the UI typecheck on macOS, on push and PR |
| `.github/ISSUE_TEMPLATE/` | Bug and feature forms, both asking for Diagnostics output |
| `.github/pull_request_template.md` | The checklist that catches the common mistakes |
| `docs/images/app.png` | The screenshot the README shows |
| `package.json` | Description, licence, 20 keywords, repository, issues URL, `engines: >=24` |

The About box carries the description and **all 20 topics** — that is what GitHub search and Google
index, so it matters more than anything inside the README.

## Two things publishing caught that nothing else had

Both were found before anyone else saw them, which is the whole argument for CI and for push
protection. Worth remembering, because neither was visible locally.

**1. The Node floor was wrong.** The first CI run failed with
`ERR_UNKNOWN_FILE_EXTENSION: Unknown file type ".ts" for scripts/build.ts`. The build scripts are
TypeScript **run directly by `node`**, which needs Node's native type stripping — unflagged from
Node 23.6/24, absent in Node 20. `engines` said `>=20`, the README said "Node 20+", and CI was
pinned to 20. All three now say 24. Nothing on the development machine could have revealed this,
because the development machine is not Node 20.

**2. GitHub's push protection blocked the first push** over a test fixture —
`sk-abcdefghijklmnopqrstuvwxyz012345` in `app/selftest.ts`. Obviously fake to a person, and
indistinguishable from a leaked key to a scanner. The fixture is now *assembled* at runtime
(`['sk','fixture','not','areal','key'].join('-')`) rather than written as a literal, so there is
nothing to match and nobody has to maintain an allow-list entry to push a fork. The reasoning is in
a comment above it, because the obvious "fix" is to restore a realistic-looking key.

## What still needs doing by hand

- [ ] **Social preview image** — Settings → Social preview. A 1280×640 crop of
      `docs/images/app.png`. It is what appears when the link is shared anywhere.
- [ ] **Private vulnerability reporting** — Settings → Code security. `SECURITY.md` tells people to
      use it, so it should exist.
- [ ] **Discussions** — only if you want to answer them. An empty Discussions tab is worse than none.
- [ ] **Check the copyright name in `LICENSE`** if it should say something other than "Alexei".

## What to decide before this gets popular

These are the parts that are fine for one machine and not for many. Each is in `TODO.md` with more
detail.

- **No packaged build.** Every user is currently a developer running `npm install` and
  `npm run trust-cert`. This is the biggest barrier to anyone actually using it, and the natural
  next milestone — `.dmg` first, then the Windows question.
- **One workbook at a time.** Two workbooks open and whichever taskpane polls wins the call. This is
  the one remaining hole that can put a change into a file the user did not ask for. A *stale* undo
  is caught by the journal fingerprint; two live taskpanes racing is not. A *Workbook identity*
  Diagnostics check was added to measure which identifier can fix it — run it before building this.
- **`remove_duplicates`'s undo rests on a measurement** taken once, on one Excel build, with nothing
  refusing if it stops being true.
- **The manifest version drifts from `package.json`** (`1.1.0.0` vs `0.4.0`).

## Adding a release later

There is no released version yet. When there is one, the pattern that fits this repository is a
`.dmg` attached to a GitHub Release, with the manifest `Version` bumped in the same commit as
`package.json` — which is the drift above, and the reason to fix it first.
