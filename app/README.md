# AI Spreadsheet Automation

Describe a task in plain language; the assistant carries it out in your **open, installed Excel**.
It can also search the web, so you never have to paste data it could look up itself.

```
You:  "find the current top 3 bestselling EVs and put them in a new sheet"
      → searches the web, writes the table, formats the header, tells you what it changed
```

## Architecture

The chat lives **outside Excel**. The taskpane is a silent worker that exists only because Excel
is the only place Office.js can run.

```
   desktop window ──POST /api/chat──▶  server ──▶ DeepSeek
   (Electron or a                        │
    browser tab)                         │  dispatches every Excel call
                                         ▼
                        taskpane ──GET /api/executor/next──┘   (long-poll, same origin)
                        (open in Excel)                        executes via Office.js
                                 ──POST /api/executor/result──▶ │
   desktop window ◀──────────── {reply} ───────────────────────┘
```

Why it's built this way:

- **The server mediates.** The window can't touch the workbook and the taskpane can't call the
  LLM, so the server routes between them. Either side can be replaced without touching the other.
- **The UI is just HTTP.** Anything that can POST to `/api/chat` is a client — the Electron
  window, a browser tab, `curl`. No UI code knows about Office.js.
- **Two kinds of tool.** `web_search` runs on the server (it needs the network and a secret);
  everything else runs in the taskpane (it needs Office.js). Server tools are resolved *inside*
  one request, so the taskpane is never asked to run something it can't.
- **Same origin throughout**, so there is no CORS, no mixed content and no port juggling.

## Setup

### 1. The certificate and the add-in — automatic

```
npm run desktop
```

On first run the app does all of this itself, and re-does whichever part is missing on every later
run:

| Step | Where it lands |
|---|---|
| Generates a local CA and a `localhost` certificate | `~/.ai-excel/certs/` (keys `0600`) |
| Adds the CA to the login keychain | one certificate, no trust setting, no admin password |
| Writes the add-in manifest | `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml` |

`npx office-addin-dev-certs install` still works and is still used as a fallback if it is what your
machine already has — the app simply no longer needs you to run it.

**Two steps stay manual**, because Office gives no way to automate them:

1. **Quit and reopen Excel** if it was already running — it reads the add-in folder only at startup.
2. **Insert → Add-ins → My Add-ins → AI Spreadsheet Automation.** Once.

The window shows a *Finish connecting Excel* panel whenever something is still outstanding, and
nothing once Excel is connected. `GET /api/setup` reports the same state; `POST /api/setup` re-runs
the silent half if you want to force it.

### The certificate: generated silently, trusted once with a password

The app generates one **self-signed certificate for `localhost`** into `~/.ai-excel/certs/` and
stores it in your login keychain. That part is silent.

**Trusting it is not, and it cannot be.** macOS raises an authorization prompt for any change to a
user trust setting — measured: `security add-trusted-cert` blocks until someone answers, while
storing and deleting a certificate return in milliseconds. An app that calls it on startup either
hangs behind a password dialog nobody asked for or silently succeeds on a machine that happens not
to ask. So the app does not call it. It tells you to run:

```
npm run trust-cert          # macOS will ask for your password
```

Once. After that the setup panel disappears. Without it Excel says *"The content is blocked because
it isn't signed by a valid security certificate"* — which names neither the certificate nor the
trust setting, and is why the app checks trust explicitly rather than assuming:

```
security verify-cert -p basic -c ~/.ai-excel/certs/localhost.crt    # "verification successful"
```

Use `-p basic`, not `-p ssl`: the ssl policy demands Certificate Transparency timestamps that no
locally-issued certificate has, so it fails on every certificate here and tells you nothing.

To undo the trust: `security remove-trusted-cert ~/.ai-excel/certs/localhost.crt`.

> An earlier version generated a local **CA** and signed a leaf with it. That asked for more than it
> needed — a trusted CA can vouch for *any* hostname, while a trusted leaf can vouch for
> `localhost` only — and the app now deletes those files and removes the CA from your keychain if it
> finds them.

### Closing the pane

**Keep the task pane open.** It is the bridge to your workbook: the assistant in the desktop window
sends every read and write through it, so closing it disconnects Excel until you reopen it
(Insert → Add-ins → My Add-ins). The window's header shows *Excel connected* when the bridge is up.

That is the whole rule, and it is deliberately a simple one. An earlier version declared a
**shared runtime** in the manifest so that closing the pane was harmless and the bridge started
with the workbook. It was removed, for reasons worth recording because the idea is attractive:

- The auto-start it enables is **document-scoped**, and Office applies it the *next* time that
  workbook is opened — so the first open of each new file still needed the add-in inserted by hand.
  It never removed the manual step it was added for.
- It makes the manifest declare a requirement set the client may or may not honour, which is a real
  way to stop an add-in loading at all.

One instruction that always works beats a mechanism that sometimes does.

> **Never `sudo` anything under `~/Library/Containers`.** It is TCC-protected *against root*, so
> `sudo` fails with `Operation not permitted` while the plain command succeeds. The app never uses
> `sudo` for any of this.

> **If Excel stops accepting the certificate**, delete `~/.ai-excel/certs` and restart the app: it
> regenerates the pair and re-adds the CA. If that fails too, `npx office-addin-dev-certs install`
> and restart — the app prefers its own certificate but will fall back to that one.

### 2. API keys (once)

The simplest way is **in the app**: click the **gear** in the header, paste the key, press Save. It
takes effect on the next message — no restart. The dialog shows the masked key, what is set, and
where each value came from.

There is also a CLI, which is what you want before a window exists or from a script:

```
pbpaste | node app/setkey.ts              # DeepSeek  — platform.deepseek.com
pbpaste | node app/setkey.ts --search     # search    — tavily.com or brave.com
```

Copy each key first, then run the command — piping keeps it out of your shell history. Both land
in **`~/.ai-excel/config.json`**, mode `0600`, deliberately outside this project.

```
node app/setkey.ts                        # show what is configured
```

Both routes run the same checks (`checkCredential` in `config.ts`), so a key the CLI accepts
cannot be one the dialog refuses. They catch the mistake that is actually made — pasting
`pbpaste | node app/setkey.ts` *as* the key — and a key for the wrong provider. **Neither route
proves a key works**: only the provider can say that, and it says it on the first call.

Environment variables (`DEEPSEEK_API_KEY`, `SEARCH_API_KEY`, `SEARCH_PROVIDER`) win over the file.
The dialog says so when one is set, because otherwise saving a key there looks like it did nothing.

The DeepSeek key is required. The search key is optional — without it the assistant simply has no
`web_search` tool, rather than failing when it tries to use one.

### 3. Install dependencies

```
npm install
```

### 4. Sideload the manifest

```
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
curl -k -o ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml \
     https://localhost:3443/manifest.xml
```

**Restart Excel** (Cmd-Q — the `wef` folder is read only at startup).

If the server is not running, the same file can be produced from the template — the server only
substitutes the base URL, so the result is byte-identical:

```
sed 's|{{BASE_URL}}|https://localhost:3443|g' app/web/manifest.template.xml \
  > ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml
```

The add-in's identity is its `<Id>`, so replacing the file **renames the existing add-in** rather
than adding a second one. Nothing else in the manifest needs to change for it to keep working:
`<SourceLocation>` already points at `https://localhost:3443/taskpane.html`.

Excel caches what it read under
`~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef`,
keyed `<Id>_<Version>`. Since the version is part of the key, a manifest with a new version is read
fresh rather than shadowed by the cached one. If the add-in list still shows a stale name after a
full quit-and-reopen, clear that cache directory and reopen Excel.

> **Never `sudo` anything under `~/Library/Containers`** — it is TCC-protected *against root*.
> The plain commands above work; `sudo` fails with `Operation not permitted`.

## The two front-ends

They are deliberately different stacks, because they have different jobs.

| | |
|---|---|
| **Desktop window** — `ui/` | React + Tailwind + shadcn/ui, built by Vite to `app/web/ui/`. This is the app: chat, activity, approval cards, undo. It has no Office.js and cannot touch the workbook. |
| **Excel taskpane** — `app/web/taskpane.*` | Plain JS, no build step. It is a worker: long-polls the server, runs Office.js, reports back. Keeping it dependency-free keeps it small and predictable inside Excel's WebView. |

The taskpane is why the build id only hashes `taskpane.html`, `taskpane.ts`, `executor.ts` and
`diagnostics.ts` — rebuilding the UI must not make an open taskpane reload itself.

## Running it

```
npm run desktop
```

Builds `ui/` and opens the window. The server starts inside that process, so quitting the app
frees port 3443.

| | |
|---|---|
| `npm run dev:ui` | Vite dev server with HMR for UI work |
| `npm run preview` | Runs the real server with a **scripted model and fake Excel**, so you can work on the window without Excel, a DeepSeek key, or a search key. It seeds a finished turn and one parked on the review gate. |
| `npm run serve` | Server only, no build. Then open `https://localhost:3443/ui/`. |
| `npm run desktop:fast` | Skip the rebuild and launch what is already built |

Then, **in Excel**: open a workbook and show the add-in —
**Insert → Add-ins → My Add-ins → AI Spreadsheet Automation**.

You now have two things open, and they do different jobs:

| | |
|---|---|
| **Desktop window** | The chat. This is the app. |
| **Taskpane in Excel** | Keep it open and ignore it. It's the bridge; closing it disconnects Excel. |

The window shows a status badge: green means Excel is connected and ready. Every control has a
tooltip explaining what it does.

## Web search

`web_search` lets the assistant fetch facts it doesn't have — stats, prices, specs, docs, current
events — and write them into the sheet. The system prompt explicitly tells it **never to ask you
for data it could look up itself**.

**It is off by default.** Use the **Search** switch in the header to turn it on:

| | |
|---|---|
| Switch off (default) | The tool is not offered to the model at all, and the system prompt says it has no web access. A search the model tries anyway is refused without contacting the provider. |
| Switch on | The assistant may search, and each search appears inline as `🔎 web_search …` |
| Switch greyed out | No search key is configured. Add one and restart the app. |

The choice is remembered between sessions. Having a key configured is deliberately *not* treated as
permission to send your questions to a search provider.

Two providers, chosen with `--provider` (default `tavily`):

| | |
|---|---|
| **Tavily** | Built for LLM agents. Returns a synthesised answer plus source snippets. Free tier ≈1000 searches/month. |
| **Brave** | Independent web index. Free tier ≈2000 queries/month. |

Both auth mechanisms were verified by probing the live endpoints rather than trusting the docs —
Brave's `X-Subscription-Token` header is confirmed; Tavily returns an identical error for every
auth failure, so both the modern and legacy forms are sent.

## Safety: preview and undo

Two protections, both on by default except the review gate.

**Every change is snapshotted first.** Before a mutating tool runs, the range it will actually
touch is read — values, formulas, number format, font, fill, alignment — and recorded against
the task. `Undo last task` replays those snapshots in reverse.

Four details that make it reliable rather than decorative:

- **The snapshot covers the *expanded* range.** `write_range("A1", 3×2)` writes `A1:B3`, not
  `A1`. Snapshotting the literal address would make undo silently incomplete, so the range
  resolution lives beside the code that does the writing.
- **Restore uses `formulas`, never `values`.** Assigning `values` to a cell that held a formula
  writes the computed number and destroys the formula. `formulas` round-trips both.
- **If the range can't be snapshotted, the write doesn't happen.** An unundoable change is worse
  than a refused one.
- **Undo refuses when the cells have moved on.** The journal is kept in
  `~/.ai-excel/journal.json`, so it outlives the app — which means it can also be older than your
  workbook. Every write is fingerprinted just after it happens, and `Undo` re-reads the range and
  compares before restoring. If you have edited those cells since, it says so and changes nothing
  rather than overwriting your newer work.

A task that changed nothing is never journalled, so undo always steps over a real change rather
than a question you asked.

**Tick "Review changes" to approve each change first.** The turn parks, a card appears showing
the range, what it holds now and what it will hold, and nothing is written until you choose
Apply or Skip. `Skip` tells the model it was declined so it can offer an alternative instead of
retrying.

Scope limits worth knowing: the journal is in memory (restarting the server loses it) and keeps
the last 20 tasks. A write larger than ~50k cells makes a correspondingly large snapshot —
measured at roughly 29 bytes per cell.

## AI usage and cost

The header shows a running total (`14.8k tok · $0.00056*`); click it or the ⓘ for the breakdown:

| | |
|---|---|
| Requests | model calls, counting each tool round trip within a task |
| Input / output tokens | with the cache-hit share |
| Web searches | counted separately, because they cost no tokens but do consume the search provider's quota |
| Speed | ≈ output tokens per second, wall clock |
| Estimated cost | model tokens only, priced **per call**, so a task spanning the peak boundary is charged correctly |

Two caveats are shown in the UI rather than hidden:

- **Cost is labelled an estimate when the model isn't in the published line-up.** This app sends
  `deepseek-chat`, which DeepSeek no longer lists — the current models are `deepseek-flash` and
  `deepseek-v4-pro` — so it is priced on the cheapest tier and marked with an asterisk.
- **Speed is a lower bound**: it is wall-clock, so it includes network latency.

DeepSeek prices peak and off-peak (peak is 01:00–04:00 and 06:00–10:00 UTC, Mon–Fri, at double
rate). Prices live only in `app/pricing.ts`; bump `PRICING_AS_OF` when you change them.

## What this deliberately does not do yet

- **One workbook at a time.** Both sides use the session id `default`. With two taskpanes open,
  whichever is polling wins each call — so don't open the add-in in two workbooks at once.
- **Only 6 Excel tools** (read/write/formula/format/selection/overview). Pivot tables, charts,
  sorting and filtering are proven working by the Diagnostics tab but not yet wired up.
- **No packaging.** `npm run desktop` runs from source; installers are a later milestone.

## Tests

```
npm test                                  # 67 checks, no Excel/network/key needed
npm run desktop -- --smoke --no-sandbox   # hidden window, fails on console errors, screenshots
```

`--smoke` loads the UI in a hidden window, fails on any renderer console error, asserts the app
actually rendered, writes a screenshot (`SMOKE_SHOT`, default `/tmp/ai-excel-ui.png`) and exits
0/1. It is how the UI gets verified without a window appearing on screen.

`--no-sandbox` is only needed when running under an outer sandbox that blocks Chromium's own;
you don't need it normally. `SMOKE_THEME=light|dark` chooses the theme.

A hidden window does not repaint, so the smoke test calls `webContents.invalidate()` before
capturing — without it the screenshot shows the frame from *before* the last state update,
which is misleading in exactly the way that matters.

67 checks, none of which need Excel, the network, or an API key. A scripted model stands in for
DeepSeek and an in-memory sheet stands in for Excel — but the wiring between them is real HTTP,
the real long-poll executor channel, and the real server-side tool loop.

| Behaviour | Why it matters |
|---|---|
| A message reaches Excel and the sheet changes | The whole point; verified end to end |
| `web_search` never reaches the taskpane | Server tools must not be dispatched to a client that can't run them |
| `web_search` is not offered without a key | Otherwise the model calls a tool that always errors |
| No taskpane → clear 409, not a hang | The most common setup mistake |
| Invalid arguments never reach Excel | A bad model response must not edit your sheet before erroring |
| Malformed tool-call JSON is caught | Models do emit truncated JSON |
| Unknown tools are refused | A hallucinated tool name must not become an action |
| The search budget is enforced | Costs money and time otherwise |
| Excel failing a tool doesn't break the turn | Errors must round-trip back to the model |
| The activity feed reports what happened | This is what the UI shows you |
| A write is snapshotted **before** it happens | Snapshotting after the write would make undo restore the wrong thing |
| Undo puts the original values back | The whole point of the journal |
| A read-only turn is not journalled | Otherwise undo silently skips past your last real change |
| A declined change is never dispatched and never journalled | Review must block the write, not merely annotate it |
| An unsnapshottable range is refused, not written | An unundoable change is worse than a refused one |
| A taskpane too old to snapshot is refused, never written to | A stale client must not be trusted to have capabilities it never had |
| The journal reloads from disk on a fresh server | Undo has to survive restarting the app, or the safety net isn't one |
| Undo is refused when the cells changed since | An old snapshot must never overwrite newer work |
| A corrupt journal file doesn't stop the app starting | A crash mid-write must not become an app that won't open |
| The build id changes when a client asset changes | This is what lets a long-lived taskpane notice it is stale |
| The certificate exception is loopback-only | A mistake there means trusting any certificate |

## Files

Everything is TypeScript. Sources live in `app/`, `ui/src/` and `scripts/`; **all generated output
goes to `dist/`**, which `npm run build` deletes and rebuilds. There is no `.js` in the source tree,
so there is never a question about which file to edit.

```
ui/src/                   React + Tailwind + shadcn/ui source for the desktop window
  App.tsx                 layout: header, transcript, approvals, composer
  lib/api.ts              all server communication, including the NDJSON stream reader
  components/ui/          shadcn-style primitives (button, tooltip, switch, …)
tsconfig.json             Node + taskpane + scripts -> dist/
tsconfig.ui.json          the React app, typecheck only
vite.config.ts            bundles ui/ -> dist/app/web/ui/ with base "/ui/"
scripts/build.ts          the one build entry point
scripts/ui-preview.ts     scripted server + fake Excel, for UI work without Excel
dist/                     GENERATED — deleted and rebuilt by npm run build
app/
  server.ts               HTTP(S) server, executor channel, tool loop, usage accounting
  ai.ts                   DeepSeek client: streaming, injectable fetch
  tools.ts                tool catalog: JSON Schema = runtime validator; executesOn: taskpane|server
  pricing.ts              DeepSeek prices — the only place to edit them
  search.ts               Tavily + Brave, server-side only
  config.ts               ~/.ai-excel/config.json
  setkey.ts               key setup CLI
  journal.ts              the undo journal on disk (~/.ai-excel/journal.json)
  selftest.ts             the 67 checks
  electron/main.ts        desktop window, smoke test
  electron/cert-policy.ts the one scoped certificate exception
  web/
    taskpane.html         the Excel worker shell (copied verbatim into dist)
    taskpane.ts           long-poll loop + status
    executor.ts           Office.js execution, snapshots
    diagnostics.ts        capability probe
    manifest.template.xml {{BASE_URL}} substituted per request (copied verbatim into dist)
```

## Troubleshooting

| Symptom | Cause |
|---|---|
| "taskpane is running an older version" | The pane was open across an update. It now detects this and reloads itself; close and reopen once to pick up the detector. |
| `the number of rows or columns ... doesn't match` | The model gave a range that disagrees with its values. The error now names both sizes so it can correct itself. |
| `ERR_CERT_AUTHORITY_INVALID` in the window | macOS holds the dev CA in the login keychain with no explicit trust settings. WebKit (Excel) accepts it; Chromium won't. The app accepts the certificate **only** for its own loopback origin and port — see `app/electron/cert-policy.ts`. |
| Status dot red, "Excel not connected" | The taskpane isn't open in Excel, or Excel was restarted |
| "Server unreachable" | The app/server isn't running |
| Excel refuses to load the add-in | Cert not trusted, or Excel wasn't fully restarted |
| Old "M0 spike" name in the add-in list | Cached manifest — quit Excel fully; if it persists, clear the `Wef` cache and reopen (see §4) |
| `Error: HTTP 401` | Bad DeepSeek key — re-run `setkey` |
| Search never happens | No search key; `/api/status` shows `"search": null`. The server reads the key at startup, so restart the app after adding one. |
| "I stopped after 12 steps" | Runaway guard. Usually the task is ambiguous — be more specific |
| Port 3443 already in use | Another server is running; the desktop app will attach to it |
