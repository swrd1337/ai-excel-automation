# AGENTS.md

Orientation for AI agents and developers working on this repository.

Read this before changing anything. Most of the rules below exist because a reasonable-looking
choice was tried and was wrong — the reasoning is in [docs/DECISIONS.md](docs/DECISIONS.md).

---

## What this is

A desktop app that carries out natural-language tasks in the user's **installed Microsoft Excel**
(macOS today; Windows should work through the same add-in, untested).

**Excel is the only target, deliberately.** LibreOffice and Google Sheets were dropped
([D-26](docs/DECISIONS.md)): one add-in already covers macOS and Windows Excel, and a second host
would need the entire undo model — typed actions, range snapshots, fingerprints — re-derived
rather than reused. Do not add them back as "just one more thing". Only DeepSeek is wired up
today; **Claude is the last piece of work**, and its key arrives at the end.

- **Excel** is driven through an **Office.js add-in taskpane**. Excel is the only place Office.js
  can run, so all workbook access happens there.
- **The chat window** is Electron + React + shadcn/ui.
- **A local Node server** mediates: it holds the model, the API keys and the safety journal, and
  dispatches every workbook operation to the connected taskpane.

The full picture is in [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

---

## Commands

**Everything is TypeScript, and every generated file lives in `dist/`.** `app/` and `ui/` contain
sources only; there is no `.js` anywhere in them. `scripts/build.ts` cleans `dist/`, compiles with
`tsc`, copies the static assets, and bundles the UI — run it after changing any `.ts`:

```
app/**/*.ts  ui/src/**/*.tsx  scripts/*.ts        (sources — edit these)
        │
        └─ node scripts/build.ts ─▶ dist/app/**/*.js, dist/app/web/**, dist/scripts/*.js
```

**Never edit anything under `dist/`** — it is deleted and rebuilt every time.

```bash
npm install
npx office-addin-dev-certs install      # once: Excel only loads a taskpane over trusted HTTPS
npm run setkey                          # once: DeepSeek key -> ~/.ai-excel/config.json
npm run setkey -- --search              # optional: Tavily/Brave key, enables web_search

npm run build:server                    # compile + copy assets, skip the UI bundle
npm run typecheck:ui                    # the React app, no emit
npm run build                           # clean, compile, copy assets, bundle the UI
npm test                                # 67 checks. Builds first. No Excel, no network, no key.
npm run validate:manifest               # Microsoft's schema validator, on the rendered manifest
npm run trust-cert                      # once: trust the local certificate (asks for your password)
npm run desktop                         # build, then open the window
npm run preview                         # scripted model + fake Excel, for UI work
npm run desktop -- --smoke --no-sandbox # hidden window, fails on console errors, screenshots
```

Run `npm run validate:manifest` after **any** change to `app/web/manifest.template.xml`. It is not
part of `npm test` because it downloads a tool and talks to Microsoft, and the suite is offline —
but it found two errors in a manifest Excel had been loading happily for weeks, which is exactly the
kind of thing that otherwise ships. The offline subset of those checks lives in `app/selftest.ts`.

`npm run desktop -- --smoke` is the way to verify the UI without seeing it. It exits non-zero on
any renderer console error, asserts the app rendered, and writes a screenshot to
`SMOKE_SHOT` (default `/tmp/ai-excel-ui.png`).

| Variable | Effect |
|---|---|
| `SMOKE_THEME=light\|dark` | Force a colour scheme |
| `SMOKE_SETTLE_MS` | How long to let the page settle before probing |
| `SMOKE_OPEN=details\|settings` | Open the Details panel or the API-key dialog first |
| `SMOKE_CLICK='<selector>'` | Click a control first (waits for it to exist *and* be enabled) |
| `SMOKE_SEND="…"` | Type into the composer and press Send |
| `SMOKE_SEND_WAIT` | How long after Send to capture — set it low to catch a reply mid-stream |

`--no-sandbox` is only needed when running inside another sandbox that blocks Chromium's own.

---

## Repo map

```
SOURCES — there is no .js anywhere in here
  ui/src/                 React + Tailwind + shadcn/ui for the desktop window
    App.tsx                 layout: header, transcript, approvals, composer
    lib/api.ts              ALL server communication, including the NDJSON stream reader
    components/SettingsDialog.tsx  where the API keys are set
    components/SetupPanel.tsx      the Excel steps the app could not do for you
    components/ui/          shadcn-style primitives
  app/
    server.ts             HTTP(S) server, executor channel, tool loop, usage accounting
    ai.ts                 DeepSeek client — streaming, injectable fetch, so it can be scripted
    tools.ts              tool catalog — the single source of truth for tool schemas
    pricing.ts            DeepSeek prices; the ONLY place to edit them
    search.ts             Tavily + Brave, server-side only
    config.ts             ~/.ai-excel/config.json
    setkey.ts             key setup CLI
    journal.ts            the undo journal on disk (~/.ai-excel/journal.json)
    tls.ts                the self-signed localhost certificate, and its trust state
    setup.ts              writes the add-in manifest into Excel's folder
    selftest.ts           the 67 checks
    electron/main.ts      desktop window, smoke test
    electron/cert-policy.ts  the one scoped certificate exception
    web/                  served statically; .ts compiles to .js in dist
      taskpane.html         the Excel worker shell        (copied verbatim)
      manifest.template.xml {{BASE_URL}} per request      (copied verbatim)
      assets/               add-in icons                 (copied verbatim)
      taskpane.ts           long-poll loop + status
      executor.ts           Office.js tool execution
      diagnostics.ts        capability probe
  scripts/
    build.ts              the one build entry point
    ui-preview.ts         scripted server + fake Excel for UI work
    probe-search.ts       is DeepSeek's native web_search available? (~$0.0005 to re-test)
    validate-manifest.ts  renders the manifest and runs Microsoft's schema validator
    trust-cert.ts         the one setup step macOS will not let the app do silently

BUILD CONFIG
  tsconfig.json           Node + taskpane + scripts  -> dist/
  tsconfig.ui.json        the React app, typecheck only (no emit)
  vite.config.ts          bundles ui/ -> dist/app/web/ui/ with base "/ui/"

GENERATED — deleted and rebuilt by `npm run build`
  dist/                   dist/app/**, dist/app/web/**, dist/scripts/**
```

## Invariants

These are load-bearing. Breaking one produces a bug that is silent, or that corrupts user data.

1. **Restore snapshots from `formulas`, never `values`.** Assigning `values` to a cell that held a
   formula writes the computed number and **destroys the formula**. `formulas` round-trips both.
   (`app/web/executor.ts`)

2. **Snapshot before every mutating write, and refuse the write if the snapshot fails.** An
   unundoable change is worse than a refused one. (`app/server.ts`)

3. **Never dispatch a tool the connected taskpane has not advertised.** The taskpane sends its
   `capabilities` on every poll; `__snapshot`, `__fingerprint`, `__restore` and
   `__workbook_context` are server-internal (double-underscore prefix) and are never offered to
   the model.

4. **Tool schemas live in `app/tools.ts`; tool execution lives in `app/web/executor.ts`.** They
   must agree on names and arguments. The schema *is* the runtime validator — do not add a second
   source of truth.

5. **The certificate exception must stay loopback + own-port scoped, and fail closed.** It exists
   because Chromium refuses the local certificate that WebKit (Excel) accepts — even though macOS
   trusts it, which it does. [D-30](docs/DECISIONS.md) corrected the original explanation for this.
   A mistake here means trusting *any* certificate. Tested in `app/selftest.ts`.
   (`app/electron/cert-policy.ts`)

6. **The build id hashes taskpane files only**, not the whole `web/` directory. Otherwise every UI
   rebuild would make an open taskpane reload itself.

7. **The window is served from `/ui/`**, so its API calls must be absolute (`/api/...`). A relative
   `api/status` resolves to `/ui/api/status`, which silently 404s.

8. **Everything under `dist/` is generated.** Sources are `app/**/*.ts`, `ui/src/**` and
   `scripts/*.ts`; there must be no `.js` in the source tree. `npm run build` owns `dist/` — but
   note that `build:server` deliberately preserves `dist/app/web/ui`, because it sits inside the
   tree being cleaned and deleting it leaves the server serving a window with no bundle.

9. **`SESSION_ID = 'default'` is a wire contract** between `ui/src/lib/api.ts` and
   `app/web/taskpane.ts`. Change it in both, and note the one-workbook-at-a-time limitation.

10. **The taskpane stays dependency-free.** Its source is TypeScript, but it compiles to plain
    JS with no runtime dependencies and no bundler, because it runs inside Excel's WebView.

11. **The server has zero runtime dependencies** (Node built-ins only). Adding one is a decision
    worth an entry in the decision log.

12. **Approval polling in the window is unconditional**, not gated on `busy`. A pending approval
    must surface whenever it exists, or the turn it belongs to sits blocked with nothing on screen
    to answer it.

13. **The chat reply is a stream, not a JSON body.** `POST /api/chat` returns NDJSON events
    (`assistant_start`, `delta`, `tool`, `search`, `approval`, `done`, `error`). Anything that
    awaits `res.json()` will hang or throw. Request validation happens *before* the stream opens,
    so malformed input is a real 400 rather than an error event inside a 200.

14. **Web search is opt-in per request.** A configured search key makes search *possible*, not
    agreed: the request carries `search: true` only when the user turned the toggle on, and an
    absent flag means off. The tool list and the system prompt both read that one flag, so they
    can never disagree. A server tool that was not offered is refused explicitly — it is still in
    the catalog, so a hallucinated call would otherwise execute.

15. **A tool's outcome is reported after it is known.** Emitting `{type:'tool', ok:true}` before
    the dispatch told users a call had succeeded when Excel had rejected it. Do not reintroduce a
    pre-dispatch success event.

16. **An undo must prove the cells are still the ones it changed.** The journal outlives the
    process now, so it can be far older than the workbook: every write that could overwrite cell
    state is fingerprinted just after it happens, and `__restore` re-reads and compares before
    writing anything. A mismatch refuses and says so — writing an old snapshot over newer work is
    the data loss the journal exists to prevent. A missing fingerprint refuses too. Tested in
    `app/selftest.ts`. (`app/web/executor.ts`, `app/journal.ts`)

17. **An API key travels one way.** `POST /api/config` takes a key; `GET /api/config` returns
    `{set, masked, source}` and never the value. A key on the window's side of the wire is a key in
    the renderer's memory, in devtools, and in every screenshot of this UI, which would make the
    config file's `0600` pointless. That is why the settings fields are empty on open and an empty
    field means "leave it alone". Tested in `app/selftest.ts`.
    (`app/server.ts`, `ui/src/components/SettingsDialog.tsx`)

18. **The server listens on loopback only, and a busy port is fatal.** `listen(port)` with no host
    binds every interface, which made this app reachable from the local network — while its
    certificate exception, its same-origin check and its "nothing leaves this machine" promise all
    assume otherwise. Both loopback families are bound, but a family that fails with `EADDRINUSE`
    must abort the whole listen: tolerating it lets a second instance bind only `::1` and silently
    shadow the first for clients that prefer IPv6. A half-bound server is worse than no server.
    Tested in `app/selftest.ts`. (`app/server.ts`)

19. **The server reports; the app acts.** Setup — the certificate and the manifest — runs in the
    app's entry point *before* the server exists, so `GET /api/setup` is read-only and no server,
    including one a test started with a stub model, can change anything about the machine. Do not
    move setup back inside `createApp`/`startServer`; that is what invariant 19 used to have to
    guard against. (`app/setup.ts`, `app/electron/main.ts`, `app/server.ts`)

20. **A certificate in the keychain is not a trusted certificate, and Excel only accepts the
    second.** `security verify-cert -p basic` is the check — `-p ssl` is useless here, because it
    demands Certificate Transparency timestamps that no locally-issued certificate will ever have,
    so it fails on every certificate and distinguishes nothing. `npm run trust-cert` is the one
    command that grants it. Tested in `app/selftest.ts`. (`app/tls.ts`, `app/setup.ts`)

21. **Never call `security add-trusted-cert` unattended.** It raises an authorization prompt and
    blocks until someone answers — measured, not assumed — so a startup that calls it can hang
    behind a password dialog nobody expected, or silently succeed on a machine that happens not to
    ask. Only `scripts/trust-cert.ts` runs it, because a person asked for it by name. Storing a
    certificate (`add-certificates`) and deleting one are silent and fine.
    (`app/tls.ts`, `scripts/trust-cert.ts`)

22. **A timeout is not a refusal.** When an approval expires unanswered the model is told exactly
    that — never that the user declined. "The user declined this change" makes the model stop and
    ask what they would prefer, for a change the user most likely never saw, because they were away
    from the machine. Silence and rejection call for different replies, and this message is the only
    thing that tells them apart. The timeout is also read per call rather than captured in a
    constant, so a test can shorten it: an expiry path that can never be exercised is one that does
    not work. Tested in `app/selftest.ts`. (`app/server.ts`)

---

## Verifying your work

**Always run `npm test` before claiming anything works.** 67 checks, no Excel, no network, no key.
A scripted model stands in for DeepSeek and an in-memory sheet for Excel, but the HTTP layer, the
executor channel, the tool loop and the safety journal are the real code.

**For UI changes, run the smoke test and look at the screenshot.** A hidden window does not
repaint, so the smoke test calls `webContents.invalidate()` before capturing — without it you get
the frame from *before* the last state update, which is misleading in exactly the way that matters.
This has already produced two false "verified" results.

```bash
npm run preview &                                        # seeds a turn + a pending approval
npm run desktop -- --smoke --no-sandbox
```

**Tests must never touch the live setup.** This was learned the hard way: a preview failed to
start with `EADDRINUSE`, the wait-loop timed out silently, and the smoke then started its own
server on **3443 with the user's real API key** — where the user's Excel taskpane, which polls 3443
because that is in the sideloaded manifest, promptly connected to it and executed a turn.

Four structural guards now exist, and all of them should stay:

- **Test servers use port 3499, never 3443.** The taskpane only ever polls 3443, so a test server
  there is reachable by a real workbook and one anywhere else is not.
- **`--smoke` never uses a real API key.** If it has to start a server itself it installs a stub
  model that refuses to answer, so a stray test can neither spend money nor write to a workbook.
- **Every test gets its own journal file in a temp directory** (`tempJournal()` in
  `app/selftest.ts`). The default is the user's real `~/.ai-excel/journal.json`, and a test that
  wrote there would destroy the journal they are relying on to undo. A test may pass an explicit
  `journalFile`; it may never fall through to the real one.
- **Certificate and manifest work is pointed at temp directories, with trusting off**
  (`AI_EXCEL_CERT_DIR`, `AI_EXCEL_WEF_DIR`, `AI_EXCEL_NO_KEYCHAIN=1`). These are the only side
  effects in this codebase that are not a file in a directory the app owns — and the test suite
  runs on the machine whose keychain it would otherwise be writing to.

Also gate on evidence rather than on a sleep: confirm `seeded turn 2` appears in the preview log
and abort if it does not. A wait-loop that times out silently is how the above happened.

**You cannot verify Office.js behaviour without Excel.** If you change `executor.ts`, say so
explicitly and ask the user to run the Diagnostics tab in Excel. Do not claim it works.

**Kill background servers properly.** `kill $!` after `npm run <script>` kills `npm`, not the
`node` child, which keeps holding port 3443 and produces confusing results on the next run. Use
`pkill -f ui-preview` (or check `lsof -nP -iTCP:3443 -sTCP:LISTEN`).

---

## Conventions

- **Comment the *why*.** Several places in this codebase look wrong until you know the reason
  (`formulas` vs `values`, the cert exception, the build-id scope). Those comments are the point.
- **Zero runtime dependencies on the server**, and that is load-bearing (invariant 11). The UI has
  a small, deliberate set — React, Radix, Tailwind, `react-markdown` — and adding to it is a
  decision worth recording, not a default. See D-23 for the tradeoff made there.
- **TypeScript everywhere, and it is not decorative.** `@types/office-js` is why the taskpane's
  Excel calls are checked at compile time; `tsc` caught three real errors on migration. Keep
  `strict` on, and keep the two tsconfigs separate: the UI must not see Node or Office globals.
- **Types only where they carry meaning.** Boundary shapes and exported functions are typed;
  do not annotate every local.
- **Errors the model can act on.** Excel's own messages ("The number of rows or columns … doesn't
  match") tell a model nothing; name the sizes so it can correct itself.
- **Never ask the user to paste data the app could fetch** — that was the original motivation for
  `web_search`.

---

## Maintaining these docs

These are living documents. Update them **in the same change** that makes them stale:

| File | Update when |
|---|---|
| `AGENTS.md` | Commands, invariants, or the repo map change |
| `docs/ARCHITECTURE.md` | A component, protocol or data flow changes |
| `docs/DECISIONS.md` | You make a non-obvious choice or discover a trap — append a dated entry |
| `TODO.md` | You finish something, or find something new that should be done |
| `app/README.md` | User-facing setup or usage changes |
| `app/pricing.ts` | DeepSeek changes prices (also bump `PRICING_AS_OF`) |

If you add a new invariant, add a test for it. An invariant without a test is a comment.

---

## Known gaps

**The list lives in [TODO.md](TODO.md)** — one place, so it cannot drift from this file. Read it
before starting work to see what is already known and what is deliberately deferred.
