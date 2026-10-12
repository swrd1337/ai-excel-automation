# Publishing this to GitHub

Everything is prepared; this is the exact sequence. Replace nothing — the URLs already point at
`github.com/swrd1337/ai-excel-automation`.

## What is already in place

| File | Why it is there |
|---|---|
| `README.md` | Rewritten for a visitor: what it is, why it is different, how to run it, honest status |
| `LICENSE` | MIT, `Copyright (c) 2026 Alexei` — **check the name before you push** |
| `CONTRIBUTING.md` | Setup, the three rules most likely to be broken, how to add a tool |
| `SECURITY.md` | How to report privately, and what this app's security model actually is |
| `.github/workflows/ci.yml` | Runs `npm test` and the UI typecheck on macOS, on push and PR |
| `.github/ISSUE_TEMPLATE/` | Bug and feature forms, both asking for the Diagnostics output |
| `.github/pull_request_template.md` | The checklist that catches the common mistakes |
| `docs/images/app.png` | The screenshot the README shows |
| `package.json` | Description, licence, keywords, repository, issues URL, `engines` |

`private: true` was removed from `package.json` on purpose — it reads as "not meant to be used".

## 1. Feel the state of the repo

```bash
cd /Users/swrd1337/Develop/ai-excel-automation
npm test              # 67 checks, and it also compiles everything
npm run build         # full build, including the UI bundle
```

Both must be green before anything becomes public.

## 2. Initialise and commit

There is no git repository here yet, so this is the first one.

```bash
git init -b main
git add -A
git status --short | wc -l     # expect a few hundred files, NOT tens of thousands
git commit -m "AI Spreadsheet Automation: natural-language control of installed Excel"
```

If `git status` lists anything under `node_modules/`, `dist/` or `.npm-cache/`, stop — the
`.gitignore` is not being applied and you are about to commit a few hundred megabytes.

## 3. Create the repository

With the GitHub CLI:

```bash
gh repo create ai-excel-automation --public --source=. --remote=origin --push
```

Or through the web UI: create **ai-excel-automation**, public, **without** a README, licence or
`.gitignore` (this repo already has all three), then:

```bash
git remote add origin https://github.com/swrd1337/ai-excel-automation.git
git push -u origin main
```

## 4. Set the About box — this is the discoverability

The description and topics are what GitHub search and Google actually index, so they matter more
than anything inside the README.

**Description** (paste into the About field):

> Drive your installed Microsoft Excel from plain English — with a real undo for every change.
> Local-first Electron + Office.js add-in, powered by DeepSeek.

**Website**: leave empty for now, or point at the repository.

**Topics** — paste these one at a time in the About box, or all at once:

```bash
gh repo edit --add-topic excel,microsoft-excel,excel-addin,office-js,office-addins,ai,llm,ai-agent,agent,natural-language,spreadsheet,spreadsheet-automation,automation,deepseek,electron,typescript,react,macos,local-first,privacy
```

Those 20 are chosen so that the repo appears for both kinds of search: someone looking for
`excel add-in office-js` (the technology) and someone looking for `ai spreadsheet automation`
(the job). `excel-addin` and `office-js` are the two highest-intent.

## 5. After the first push

- **Watch the CI run.** It is the first time this has ever run anywhere but your Mac. If it fails
  for an environment reason rather than a real one, delete the badge line in `README.md` and the
  workflow, rather than leaving a red badge on a public repo.
- **Add a social preview image**: Settings → Social preview. A 1280×640 crop of
  `docs/images/app.png` is the honest choice, and it is what shows up when the link is shared.
- **Enable private vulnerability reporting**: Settings → Code security → Private vulnerability
  reporting, so `SECURITY.md`'s instruction actually works.
- **Turn on Discussions** only if you want to answer them. An empty Discussions tab is worse than
  none.

## 6. Things to decide before this gets popular

Written down here because they are the parts that are fine for one machine and not for many:

- **The manifest version drifts from `package.json`** (`1.1.0.0` vs `0.4.0`). Two numbers that move
  independently will disagree. See `TODO.md` §0.
- **`remove_duplicates`'s undo rests on a measurement** taken once, on one Excel build, with nothing
  refusing if it stops being true. `TODO.md` §0 has the detail.
- **One workbook at a time.** Two workbooks open and whichever taskpane polls wins the call — the
  one remaining hole that can put a change in a file the user did not ask for. §1.
- **No packaged build**, so every user is currently a developer running `npm install`. That is the
  biggest barrier to anyone actually using this, and it is the natural next milestone.
