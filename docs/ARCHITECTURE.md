# Architecture

How the pieces fit, and why. Decisions and their reasoning are in
[DECISIONS.md](DECISIONS.md).

---

## Shape of the system

Three processes, one of which exists only because Excel demands it.

```
┌─────────────────────────┐        ┌──────────────────────────────┐
│  Desktop window          │        │  Excel                       │
│  Electron + React        │        │  ┌────────────────────────┐  │
│                          │        │  │ Add-in taskpane        │  │
│  chat · streamed reply · │        │  │ the only place         │  │
│  approvals · undo        │        │  │ Office.js can run      │  │
└───────────┬─────────────┘        │  └───────────┬────────────┘  │
            │ POST /api/chat                    │               │
            │ GET  /api/status                  │               │
            │ GET  /api/approvals               │               │
            ▼                                   ▼               │
      ┌──────────────────────────────────────────────────┐      │
      │  Local Node server (in the Electron process)     │      │
      │  · model + tool loop      · snapshot journal     │      │
      │  · API keys (never leave) · approval gate        │      │
      │  · NDJSON chat stream     · usage accounting     │      │
      └──────────────────────────────────────────────────┘      │
            │  long-poll: taskpane asks for work                │
            └───────────────────────────────────────────────────┘
```

Everything is served from **one origin** (`https://localhost:3443`). That is deliberate: it
removes CORS, mixed content and Private Network Access from the picture entirely.

---

## Components

| Component | Responsibility | Stack |
|---|---|---|
| `ui/` | The chat window. Renders, streams, sends. No Office.js, no workbook access. | React + TypeScript, Tailwind, shadcn/ui, built by Vite || `app/server.ts` | The hub. Compiles to `dist/app/server.js`. Holds the model, the conversation, the journal, the approval gate. | Node built-ins only |
| `app/journal.ts` | The undo journal on disk, and nothing else. Never reads inside a record. | Node built-ins only |
| `app/tls.ts` | Generates the local CA and `localhost` certificate, and puts the CA in the login keychain. | Node built-ins + `openssl`, `security` |
| `app/setup.ts` | Writes the add-in manifest into Excel's folder and reports what a person still has to do. | Node built-ins |
| `app/web/taskpane.ts` | The Excel worker. Long-polls for work, executes it via Office.js, reports back. | TypeScript, compiled to plain JS with no bundler |
| `app/web/executor.ts` | The actual Office.js operations, plus the undo actions that reverse them. | TypeScript, compiled to plain JS with no bundler |
| `app/electron/main.ts` | Window, cert policy, smoke test. Starts the server in-process. | Electron |

---

## Protocol 1 — a chat turn (window ↔ server)
One request, a **streamed NDJSON response**. The reply is streamed from the model, so the window
prints it as it is written and shows each tool call the moment it runs.

```
window  ──POST /api/chat {sessionId, message, review}──▶  server ──▶ DeepSeek (stream)
        ◀── {"type":"status"}                                  │
        ◀── {"type":"assistant_start"}                         │
        ◀── {"type":"delta","text":"Added "}   ←── printed live
        ◀── {"type":"tool","name":"write_range","ok":true}
        ◀── {"type":"done","reply":"…","stats":{…}}
```

| Event | Meaning |
|---|---|
| `status` | A stage began ("Reading the workbook…") |
| `assistant_start` | Open a new bubble — one per model round that produces prose |
| `delta` | Append text to the open bubble, now |
| `tool` | A tool finished; `ok` reflects what actually happened |
| `search` | A web search is running |
| `approval` | The review gate was answered |
| `done` | Final reply text, plus session usage stats |
| `error` | The turn failed |

Request validation happens **before** the stream opens, so a malformed body is a real HTTP 400
rather than an error event inside a 200.

An earlier progress feed (`/api/activity`) was removed once the stream made it redundant; it
duplicated information the stream already carries, and a second, staler source of truth for "what
happened" is a liability.

---

## Setup on first run (app → this machine)

Nothing here is a server call or a workbook operation: it is the app making itself known to Excel.
It runs at startup, silently, and the window shows only what is left over.

```
startServer
  ├─ ensureCertificate()   openssl → ~/.ai-excel/certs/{ca,localhost}.{crt,key}
  │                        security add-certificates → login keychain
  ├─ installManifest()     render manifest.template.xml → wef/manifest.xml
  └─ resolveTls()          serve with the certificate that now exists
                           (falls back to office-addin-dev-certs, then to plain HTTP + a warning)
```

| Location | Who owns it |
|---|---|
| `~/.ai-excel/certs/` | The app. Regenerated if it is missing, mismatched, or within 30 days of expiry |
| `~/.ai-excel/config.json` | The app, via the settings dialog or `setkey.ts` (D-27) |
| `~/.ai-excel/journal.json` | The app (D-25) |
| `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` | **Excel**, written by the app |
| The login keychain | One CA certificate, added once, no trust setting needed (D-5, D-28) |

Two steps cannot be automated, and the app says so rather than implying success:

- **Excel must be restarted.** It reads the add-in folder only at startup.
- **The add-in must be inserted once** (Insert → Add-ins → My Add-ins). Office exposes no way for
  an outside process to do this. After the first time, the taskpane can ask to start itself.

`GET /api/setup` returns the current state and `POST /api/setup` re-runs the silent half. Both are
absent (`available: false`) on a server that was handed its model, so a smoke test can never
generate a certificate or write into Excel's folder.

---

## Protocol 2 — credentials (window ↔ server)

```
window  ──GET  /api/config────────────────────▶  server ──▶ config.ts ──▶ ~/.ai-excel/config.json
        ◀── {configurable, config:{deepseek:{set,masked,source}, search:{…}}}
        ──POST /api/config {deepseekApiKey:"sk-…"}──▶  validate ──▶ write 0600 ──▶ rebuild clients
        ◀── {ok, config:{…masked…}, warning}
```

The shape of this is decided by one rule: **a key travels one way.** It goes up in a POST and is
never sent back — `GET` returns `{set, masked, source}` and nothing else, so the raw key exists in
the config file and in the model client and nowhere else. That is why the form's fields are always
empty on open and an empty field means "leave it alone": there is nothing to pre-fill it with, and
pre-filling it from the server would mean the key was on this side of the wire.

`source` says which value is actually in charge, because an environment variable beats the file —
without it, saving a key while `DEEPSEEK_API_KEY` is set looks like the form did nothing.

The endpoint is refused when the server does not own its model (`configurable: false`). That is
the case for `--smoke`, which is handed a stub, and for the preview harness. It matters: it is what
stops opening the settings dialog during a test run from building a real, keyed client out of the
user's own config file. A cross-origin POST is refused as well; a request with no `Origin` is
allowed, because the tests and the CLI are not browsers.

Saving rebuilds the model and search clients and swaps them into the running server, so a new key
takes effect on the next message rather than at the next restart.

`app/setkey.ts` still exists and runs the same `checkCredential` rules. It is the only route
before a window exists, and the only one that works from a script.

---

## Protocol 3 — the executor channel (server ↔ taskpane)

A **server-driven** long-poll. The taskpane asks for work; the server holds the request until
there is something to do or the poll times out.

```
taskpane ──GET /api/executor/next?sessionId&capabilities&timeout──▶ server
         ◀──── 200 {call:{id,name,args}}  or  204 (timeout) ───────
taskpane  executes via Office.js
         ──POST /api/executor/result {sessionId, id, content|error}─▶
```

Three internal call names never reach the model and are never advertised as tools:

| Call | Purpose |
|---|---|
| `__workbook_context` | Sheet list, used ranges, current selection — fetched from Excel rather than supplied by the UI, so any UI gets it for free |
| `__snapshot` | "What does this tool affect, how is it put back, and what is in it right now?" — resolved in the taskpane, which is the only place that knows how a tool expands its range |
| `__fingerprint` | "What did this change leave behind?" Asked after a write, so a much later undo can tell whether anything has touched those cells since |
| `__restore` | Replay one recorded undo action |

Capabilities are advertised on every poll, which is how the server knows whether an open taskpane
can snapshot before it dispatches a write.

---

## Protocol 4 — the approval gate

Blocking, and it spans all three processes.

```
server: about to write, review enabled
  1. __snapshot to Excel                      -> {action, label, preview, seal}
  2. push an approval, park the turn          -> UI polls /api/approvals
  3. UI shows the card, user clicks Apply     -> POST /api/approvals/decision
  4. turn resumes; the write is dispatched    -> Excel
  5. __fingerprint (only when seal)           -> {expect}
  6. {action, label, expect} is appended to the journal, and written to disk
```

If the user does nothing, the approval times out after 120 s and the change is not applied. If the
user declines, the model is told so and does not retry.

---

## Undo actions

`__snapshot` answers with an **undo action**: the instruction `__restore` needs to reverse the
change, worked out *before* the change is made. It is a union, because not every change is a range
of cells.

```ts
type UndoAction =
  | { kind: 'range';         sheet; address; state }   // formulas, number format, font, fill, alignment
  | { kind: 'created-table'; sheet; name }             // undone with convertToRange(), not delete()
```

Adding a kind is a taskpane-only change: the server stores a record, hands it back, and never
reads inside it, so the union lives in one place (`app/web/executor.ts`) instead of drifting
between two copies on either side of the wire. What the server does need is a `label` — the words
the taskpane used to describe the target ("Sheet1!A1:C3", "SalesTable — a table over Sheet1!A1:D20")
— for the approval card and for undo messages.

The constraint this puts on a new tool is worth stating plainly: **if the change cannot be
described before it is made, it cannot be made.** A created table is named by the model rather
than by Excel, precisely so the plan can be written down first. Return to this when adding
`create_chart` / `create_pivot_table` / `manage_sheet`.

---

## The journal on disk

The journal lives at `~/.ai-excel/journal.json`, mode 0600, keyed by session id — beside the config,
outside any checkout. It is written with write-then-rename, so a crash mid-write leaves the previous
journal rather than a half-file that reads back as "nothing to undo". A missing or corrupt file is
logged and treated as empty; losing the ability to undo is bad, refusing to open the app is worse.

Because a file outlives the process, a record can be far older than the workbook it describes. So
each entry carries an `expect` fingerprint of what the change left behind, and `__restore` re-reads
the range and compares before writing anything:

```
   write happens ──▶ __fingerprint ──▶ {action, label, expect} ──▶ journal.json
                                                       │
   later undo ──▶ __restore ──▶ read the range ──▶ matches `expect`?
                                                   ├─ yes ──▶ restore
                                                   └─ no  ──▶ refuse, say so
```

Only the kinds that rewrite cell state ask for a fingerprint (`seal` on the snapshot result), and
only those refuse when it is missing. Removing an object you created cannot overwrite data, so a
stale-check there would protect nothing.

The fingerprint is 64 bits of FNV-1a over the JSON of everything `__restore` would put back — not a
cryptographic hash, because `node:crypto` does not exist inside Excel's WebView, and this is a
staleness check rather than a security boundary.

---

## Safety model

| Mechanism | Guarantee |
|---|---|
| **Snapshot before write** | Every mutating operation captures the target range's formulas, number format, font, fill and alignment first; a tool that creates an object records how to remove it instead |
| **Refuse on snapshot failure** | If the change cannot be captured — or described in a way `__restore` can replay — the write does not happen. An unundoable change is worse than a refused one |
| **Journal** | Undo actions are grouped per user turn. A turn that changed nothing is not journalled, so undo steps over real changes only |
| **Undo** | Replays the last turn's actions in reverse order. Newest first, because later writes can overlap earlier ranges |
| **Refuse a stale undo** | Before restoring, the range is re-read and compared with the fingerprint taken when the change was made. If anything has touched those cells since, the undo refuses rather than overwrite newer work |
| **Review gate** | Optional. Nothing is written until the user approves |

The journal is capped at 20 turns per session and **persisted** to `~/.ai-excel/journal.json`, so
undo survives a restart. It describes the workbook, not the conversation, which is why starting a
new chat does not clear it and why a restart does not lose it.

---

## State (per session, in the server)

```js
{
  messages,          // model conversation, OpenAI format
  queue, waiter,     // executor channel
  pending,           // in-flight dispatched calls, keyed by id
  capabilities,      // what the connected taskpane can execute
  lastPollAt,        // staleness -> "Excel is not connected"
  journal,           // undo: [{at, entries:[{action, label, expect}]}], loaded from disk
  approvals,         // pending decisions
  usage,             // tokens, cost, timing
}
```

---

## Security boundaries

| Boundary | Treatment |
|---|---|
| **API keys** | Server-side only, in `~/.ai-excel/config.json` (mode 0600). They never enter a WebView, and `GET /api/config` returns only a masked form (D-27) |
| **Reachability** | The server binds `127.0.0.1` and `::1` and nothing else, so it is not on the network at all. A port already in use is fatal rather than half-bound (invariant 18) |
| **Certificates** | Verified normally, *except* our own loopback origin and port — see `cert-policy.ts` and invariant 5 in AGENTS.md |
| **Request bodies** | Size-capped (1 MB) and JSON-parsed with an explicit error |
| **Static files** | Path-resolved and rejected if they escape the web root |
| **Tool arguments** | Validated against the tool's own JSON Schema before dispatch; invalid calls are returned to the model as errors and never executed |
| **Unknown tools** | Refused before dispatch, so a hallucinated tool name cannot become an action |
| **Electron renderer** | `contextIsolation: true`, `nodeIntegration: false`, `sandbox: true`; external links open in the real browser |
| **Model output** | Treated as untrusted: every argument is validated, and every write is snapshotted |

---

## Failure modes

| Condition | Behaviour |
|---|---|
| Excel not connected | `/api/chat` returns 409 with instructions to open the taskpane |
| Taskpane running old code | Server refuses mutating writes with an actionable message; the taskpane also reloads itself when the build id changes |
| Excel fails an operation | The error is returned to the model as a tool result so it can correct itself |
| Model loops | Capped at 12 Excel round trips and 3 searches per turn |
| Model emits malformed JSON args | Caught, reported back, never executed |
| Server restarts | Conversation and approvals are lost; **the journal reloads from disk**, and Excel reconnects on its next poll |
