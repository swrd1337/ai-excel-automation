# Decision log

Why things are the way they are. Append a dated entry when you make a non-obvious choice or
discover a trap; do not rewrite history — supersede it.

Format: **D-n — title**, then Context / Decision / Consequence.

---

## D-1 — Drive Excel through an Office.js add-in, not AppleScript or COM
*2026-09-23*

**Context.** The obvious way to script Excel from a desktop app is AppleScript on macOS and COM on
Windows — two implementations, two sets of quirks, and a macOS path Microsoft does not actively
maintain (Outlook for Mac already dropped AppleScript). The user asked for Office.js explicitly
for future-proofing.

**Decision.** An Office.js add-in taskpane, loaded inside Excel.

**Consequence.** One backend covers **macOS, Windows and Excel on the web** — the Windows COM
adapter and the macOS AppleScript adapter both disappeared from the plan. In exchange, the add-in
must be sideloaded, needs a trusted HTTPS origin, and Excel must be open with the taskpane visible.

---

## D-2 — The taskpane drives the tool loop, not the server
*2026-09-23*

**Context.** The first design had the server push work to the taskpane over SSE or a WebSocket, so
the server could orchestrate everything.

**Decision.** Invert it: the taskpane POSTs a message and receives tool calls back, then POSTs the
results. Plain request/response, driven by whoever touches Excel.

**Consequence.** No sockets, no streaming, no reconnection logic, no long-poll for the *chat* path.
The server holds state only keyed by `sessionId`. When the chat window later moved outside Excel, a
second, genuinely server-driven channel (the executor channel, D-8) was needed — but it reuses the
same idea in the one direction that actually required holding a request open.

---

## D-3 — One origin for everything
*2026-09-23*

**Context.** The window needs the AI endpoints; Excel needs the taskpane. Making them separate
local servers would mean cross-origin calls to `http://127.0.0.1`, which browsers variously block
as mixed content or Private Network Access.

**Decision.** The server serves the taskpane, the built UI **and** the API from the same origin.

**Consequence.** No CORS, no mixed content, no port juggling. A whole class of failure simply does
not exist. The cost: the UI is served from `/ui/`, so every API path in it must be absolute.

---

## D-4 — An early "bridge probe" produced a false finding
*2026-09-23*

**Context.** The M0 spike probed whether a taskpane could reach `http://127.0.0.1:3443` to test
whether a side-car process could drive it. It reported BLOCKED.

**Decision.** Discard the result. The server was serving **HTTPS only** on that port, so plain HTTP
could never have answered — the probe measured nothing about WebView policy. The probe was also
asking the wrong question, since D-3 removed the cross-origin hop entirely.

**Consequence.** A lesson that recurs in this project: a test that cannot fail for the reason you
think is not evidence. This is why the smoke test now asserts the page actually rendered, and why
`checkShape` failures name both sizes.

---

## D-5 — Snapshot before every write; refuse if the snapshot fails
*2026-09-23*

**Context.** An LLM writing into live business files without a way back is not shippable.

**Decision.** Before any mutating tool runs, capture the target range. If it cannot be captured, do
not perform the write.

**Consequence.** Undo is possible, and "cannot undo" can never happen silently. Cost is ~29 bytes
per cell, measured on real data.

---

## D-6 — Restore from `formulas`, never `values`
*2026-09-23*

**Context.** A snapshot naturally keeps `values`. But assigning `values` to a cell that held a
formula writes the *computed number* and destroys the formula.

**Decision.** Snapshot both; **restore from `formulas`**, which returns the formula string for
formula cells and the plain value everywhere else, so it round-trips both.

**Consequence.** Undo does not silently convert a user's formulas into dead numbers. This is
invariant 1 in AGENTS.md, and it is the single most damaging bug this project has avoided.

---

## D-7 — Resolve a tool's target range in the taskpane
*2026-09-23*

**Context.** `write_range("A1", 3×2)` writes `A1:B3`, not `A1`. The server could re-derive that,
but then two implementations of "what does this tool affect" would exist — and the snapshot could
silently cover the wrong range, making undo look like it worked while leaving part of the change
behind.

**Decision.** The taskpane answers "what range does this affect?" via `__snapshot`, next to the
code that does the writing.

**Consequence.** One source of truth. Also means the server cannot snapshot without a taskpane that
supports it — which led to D-9.

---

## D-8 — A server-driven long-poll for the executor channel
*2026-09-23*

**Context.** Once the chat window moved outside Excel, the window could no longer execute tool
calls. The server had to dispatch them, and the taskpane had to receive them without the window
being involved.

**Decision.** The taskpane long-polls `/api/executor/next`; the server holds the request until work
arrives. Results come back on a separate POST.

**Consequence.** Work reaches Excel in milliseconds, the UI protocol stayed unchanged, and no
socket library was needed. The window still uses plain request/response for chat.

---

## D-9 — Capabilities are advertised, and internal calls are included
*2026-09-23*

**Context.** A taskpane left open across an update had no `__snapshot`. The server assumed it did,
dispatched it, and the user saw `this taskpane cannot execute "__snapshot"` — after the change had
already been refused. Correct behaviour, useless explanation.

**Decision.** The taskpane advertises **every** executor it has, including the `__`-prefixed
internals. The server refuses a mutating write when `__snapshot` is not advertised, with a message
naming the fix.

**Consequence.** A stale client cannot be sent a tool it never had. Internal calls stay out of the
model-facing catalog by naming convention, not by omission from the handshake.

---

## D-10 — Build id from taskpane files only
*2026-09-23*

**Context.** The taskpane reloads itself when the server's build id changes. Hashing all of `web/`
meant every UI rebuild reloaded the taskpane for no reason.

**Decision.** Hash only `taskpane.html`, `taskpane.ts`, `executor.ts`, `diagnostics.ts`.

**Consequence.** The UI can be rebuilt freely. Invariant 6.

---

## D-11 — A scoped certificate exception
*2026-09-23*

**Context.** macOS has the Office dev CA in the **login keychain with no explicit trust settings**.
WebKit (Excel) accepts it; Chromium (Electron) rejects it with
`ERR_CERT_AUTHORITY_INVALID`, so the app could not load its own UI.

**Decision.** Accept the certificate **only** for our own loopback origin and port, and fail closed
for everything else. Extracted into `cert-policy.ts` and tested, because the failure mode is
silently trusting any certificate.

**Consequence.** No system trust-store changes are needed. The hostname check is exact-match, so
`localhost.evil.com` is rejected; `0.0.0.0` is deliberately not treated as loopback.

---

## D-12 — React/shadcn for the window, vanilla JS for the taskpane
*2026-09-23*

**Context.** The user asked for shadcn/ui. shadcn is React, which means a bundler, Tailwind and a
build step in what had been a zero-dependency static app.

**Decision.** Split the front-ends. The desktop window becomes a Vite-built React app; the taskpane
stays dependency-free vanilla JS.

**Consequence.** The window gets proper components and tooltips; the taskpane stays small and
predictable inside Excel's WebView. Two stacks is a real cost, accepted because the taskpane's UI
is a status line. `--smoke` now screenshots the window so UI changes can be reviewed without a
human looking at a screen.

---

## D-13 — A screenshot is not evidence unless the window repainted
*2026-09-23*

**Context.** The smoke test's first two screenshots showed the empty state while the DOM probe
confirmed the approval card existed. A hidden Electron window does not repaint, so `capturePage()`
returned the frame from *before* the state update.

**Decision.** Call `webContents.invalidate()` and wait before capturing.

**Consequence.** Two false "verified" results were caught. Any future screenshot-based verification
must do the same, or it will confirm the wrong thing.

---

## D-14 — Usage is an estimate and says so
*2026-09-23*

**Context.** The app sends `deepseek-chat`, which DeepSeek **no longer publishes** — the current
line-up is `deepseek-flash` and `deepseek-v4-pro`. It still resolves, but which model actually
served a request cannot be known from here. Pricing is also peak/off-peak, and peak excludes
Chinese public holidays, which cannot be determined offline.

**Decision.** Track tokens, wall-clock speed and cost; price the unknown model on the cheapest tier
and **mark the figure as estimated** in the UI. Missing cache reporting is treated as cache-miss so
the number over-estimates rather than flattering itself.

**Consequence.** The user gets a useful figure with an honest label instead of false precision.
`app/pricing.ts` is the only place to update prices.

---

## D-15 — Approval polling is unconditional
*2026-09-23*

**Context.** The window only polled for approvals while it was busy. An approval created while the
window was idle would sit pending with nothing on screen to answer it, blocking the turn.

**Decision.** Poll approvals on their own interval, always.

**Consequence.** A pending decision always surfaces. Invariant 12.

---

## D-16 — TypeScript everywhere, compiled in place
*2026-09-23*

**Context.** The project began as plain JS to keep it dependency-free and fast to iterate.

**Decision.** Convert everything — server, taskpane, Electron main, scripts, and the React app — to
TypeScript. `tsc` compiles `app/**/*.ts` and `scripts/**/*.ts` **in place**, so `app/server.js` sits
beside `app/server.ts` and every existing path still resolves. The UI has its own `noEmit` config
for typechecking, and Vite compiles it.

**Consequence.** `@types/office-js` checks the Excel calls at compile time — it is what would have
caught the `conditionalFormats.count` and `autoFilter.range` mistakes. Migration surfaced three
real errors (`res.end()` returning `ServerResponse`, `diagnostics.build` missing from the published
type, and `numberFormat` being typed narrower than the API). Cost: a build step before anything
runs, and generated `.js` files in the tree. Two tsconfigs are deliberate — the UI must not see
Node or Office globals. TypeScript 7 removed `baseUrl`, so `paths` are now relative.

## D-17 — The chat reply is a stream
*2026-09-23*

**Context.** The user asked for realtime text printing, as in other AI chats.

**Decision.** `POST /api/chat` returns NDJSON, one event per line, and the model is called with
`stream: true`. Text deltas are forwarded as they arrive; tool events go out as they finish.

**Consequence.** The reply prints as it is written and tool calls appear live. Two consequences
worth remembering: the server must send `stream_options: {include_usage: true}` or a streamed
response carries no token counts and the usage panel silently reads zero; and tool-call arguments
arrive in fragments keyed by index, so `foldDeltas` reassembles them — treating one chunk as a
whole call yields truncated JSON. `foldDeltas` is a pure function with tests for exactly that.

## D-18 — A tool's outcome is reported after it is known
*2026-09-23*

**Context.** Both the activity feed and, initially, the new stream emitted the tool event *before*
dispatching it, with `ok: true`. A call that Excel then rejected was shown to the user as a
success.

**Decision.** Emit the tool event after the outcome is known, with `ok` reflecting reality and the
error as the detail.

**Consequence.** The pre-existing bug only surfaced because the TypeScript rewrite added an
assertion on it. Lesson worth keeping: when adding a feature that reports on existing behaviour,
assert on the *content*, not just the presence, of the report.

---

## D-19 — A test server must never share the live port
*2026-09-23*

**Context.** While capturing a screenshot of a UI change, the smoke test reported a real model name
instead of the scripted one. The chain:

1. A preview server was still holding port 3443 from an earlier run, so the new one died with
   `EADDRINUSE`.
2. The shell loop waiting for the preview's "seeded" line **timed out silently**, and the next
   command ran anyway.
3. The smoke test therefore started its own server — on 3443, with the **user's real API key**.
4. The user's Excel taskpane polls 3443, because that is baked into the sideloaded manifest. It
   connected to the throwaway test server and a turn was dispatched against a live workbook.

**Decision.** Two structural guards, because "be more careful" is not a fix:

- Test servers bind **port 3499**, never 3443. The taskpane only ever polls 3443, so a test server
  there is reachable by a real workbook and one anywhere else cannot be.
- `--smoke` **never uses a real API key**. If it must start a server it installs a stub model that
  refuses to answer, so a stray test can neither spend money nor write to a workbook.

Plus a procedural one: gate on evidence. Confirm the preview actually seeded and abort otherwise.
A wait-loop that times out silently is precisely how this happened.

**Consequence.** A test can no longer reach live infrastructure even when the operator makes a
mistake — which is the only kind of guard worth having. The preview also now fails loudly, with the
remedy (`pkill -f ui-preview`), instead of dying with a stack trace.

---

## D-20 — DeepSeek's native web search is not available; keep our own provider
*2026-09-24*

**Context.** Two independent issues from 2026-08-06
([hermes-agent #79820](https://github.com/NousResearch/hermes-agent/issues/79820),
[oh-my-pi #7794](https://github.com/can1357/oh-my-pi/issues/7794)) state that DeepSeek's Responses
API supports **server-side `web_search`** on `deepseek-v4-flash`, streaming
`response.web_search_call.*` events and returning `web_search_call` output items with citations. If
true, this app could drop its Tavily/Brave dependency and its third-party key.

The current [official compatibility table](https://api-docs.deepseek.com/guides/responses_api) says
the opposite — `web_search` under `tools` is **"Ignored"** — while still documenting that
`web_search_call` items are accepted in `input`. That footnote is the tell: *"search results
produced by an earlier request with an **older model**"*.

**Decision.** Measured it (`scripts/probe-search.ts`, both `/responses` and `/v1/responses`, model
`deepseek-flash`, ~$0.0005):

```
status:     200
item types: reasoning, message
searched:   no
reply:      "We have no browser tools listed. We need tell if cannot search."
```

No `web_search_call` item, and the model's own reasoning states it has no search tool. The tool
entry was silently ignored. **Keep the server-side Tavily/Brave provider**, which is implemented,
tested and one key away from working.

**Consequence.** The August reports describe behaviour that no longer applies to the current model;
it was withdrawn or gated, and the API retains only input-side compatibility. The probe is kept in
`scripts/probe-search.js` so this can be re-tested cheaply when DeepSeek ships it again.

The general lesson, and this is the second time it has bitten: **do not build on undocumented
provider behaviour.** The published compatibility table is the contract. A capability that appears
only in third-party issue trackers can vanish between releases, and this one did.

**Incidental findings worth knowing.** The Responses API itself works with `deepseek-flash` on both
endpoint paths. And `deepseek-flash` runs in **thinking mode by default** — 285 of 359 output tokens
in the probe were `reasoning_tokens`, which are billed as output. Any future move to that transport
should expect roughly triple the output cost of a non-thinking call.

---

## D-21 — Sources are pure TypeScript; all output goes to `dist/`
*2026-09-24*

**Context.** TypeScript was compiled **in place**, so `app/server.js` sat beside `app/server.ts` and
`app/web/ui/` (the Vite bundle) sat inside the source directory. Fifteen generated `.js` files were
interleaved with the sources, and "which file do I edit?" had to be answered by convention rather
than by looking.

**Decision.** `app/`, `ui/src/` and `scripts/` contain sources only. `scripts/build.ts` cleans
`dist/`, compiles with `tsc`, copies the static assets and bundles the UI, so `dist/` mirrors the
source tree and nothing generated is ever ambiguous. Node 24 runs TypeScript directly, so the build
script itself needs no compile step. Added a `.gitignore` so the split is visible to anyone opening
the repo.

**Consequence.** The source tree has no `.js` at all. One trap came out of it, worth keeping: the
first version of `--no-ui` deleted all of `dist/` before compiling — including `dist/app/web/ui`,
which it was deliberately not rebuilding. The result was a window served with no bundle and a smoke
test reporting "the React app did not render". `build:server` now cleans around that subtree rather
than over it. Generated output and the source that produces it must not share a directory.

---

## D-22 — Per-turn state does not belong in conversation history
*2026-09-24*

**Context.** The workbook context packet was concatenated onto the user's message:

```ts
s.messages.push({ role: 'user', content: `${message}\n\n--- current workbook state ---\n${ctxText}` });
```

So every turn left a full snapshot in the conversation. Nothing ever removed them, and `MAX_HISTORY`
only trimmed the oldest messages once the list grew long. Ten turns in, the model was reading ten
descriptions of the workbook — several of them describing a sheet shape that had since changed.

**Decision.** History stores the user's own words, and nothing else. The context travels in the
system prompt, rebuilt from Excel each turn and attached only to the first round, because later
rounds have the model's own fresher tool results.

**Consequence.** Two problems fixed at once: fewer tokens, and no chance of the model acting on a
stale view of the workbook — which was the more serious of the two, since it would look like the
model misunderstanding rather than a data bug. The test asserts the context block appears exactly
once in the request, so a regression fails loudly.

The general rule: **anything that describes the world *now* is per-turn state, not history.**
Conversation history is for what was said, not for what was true at the time.

---

## D-23 — Render markdown with react-markdown, not a hand-rolled formatter
*2026-09-24*

**Context.** Model replies are markdown, but the assistant bubble rendered `{entry.text}` as plain
text, so `**bold**` and inline code appeared literally to the user. Replies that combine a search
result with a summary are structured — bold labels, bullet lists, and tables — because this is a
spreadsheet assistant and tabular answers are a normal case.

**Decision.** `react-markdown` with `remark-gfm` (tables, strikethrough) and `remark-breaks`
(single newline = line break). Styling is mapped element by element with Tailwind classes rather
than adding `@tailwindcss/typography`, so the output matches the bubble it sits in and there is no
second theme to keep in sync.

**Why not hand-roll it.** ~80 lines of regex would cover bold, italics and lists, and would be
wrong on the cases that actually appear: GFM tables, nested lists, and emphasis inside inline code.
Markdown is a specification; a partial implementation of it fails quietly on real input, which is
the same failure mode this project keeps running into elsewhere.

**Consequence.** The UI bundle went from 488 kB to 868 kB (gzip 107 kB → 184 kB). Served from
loopback that costs nothing measurable in load time, and the alternative is user-facing text that
renders wrong. `react-markdown` does not render raw HTML unless explicitly enabled, so model output
cannot inject markup into the window. `remark-breaks` is not cosmetic: without it, two adjacent
lines of model output silently join into one paragraph — the same class of bug as leaving the
markers visible.


---

## D-24 — Undo is a union of typed actions, and the server never reads it
*2026-09-24*

**Context.** The journal stored a range snapshot: `{sheet, address, state}` plus a preview. That
works for anything that changes cells, and it is why `sort_range` was safe to add. It does not work
for a tool that *creates* something. `create_table` makes a ListObject over cells that already held
data; replaying a snapshot of those cells restores values that were never touched and leaves the
table sitting there, so "Undo last task" would report success and change nothing. Nine of the ten
remaining tools are blocked on the same question — what does undo actually mean for this tool? —
and wiring them up before answering it would make the undo button a lie.

**Decision.** `__snapshot` returns an **undo action** rather than raw cell state:

```ts
type UndoAction =
  | { kind: 'range';         sheet; address; state }
  | { kind: 'created-table'; sheet; name }
```

`create_table` is undone with `convertToRange()`, not `delete()` — the cells held that data before
the table existed, so deleting the table would delete the user's data along with it.

Two supporting choices:

- **The server never looks inside `action`.** It stores `{action, label}`, hands `action` back to
  `__restore`, and shows `label` to the user. Mirroring the union in `server.ts` would put the same
  knowledge on both sides of the wire, where the copies drift — and the server's need is genuinely
  only "something to store, something to hand back, a name for a person". `label` replaced the
  approval card's `address` field, because a created object does not have an address.
- **The tool asks the model to name the table.** Excel would otherwise invent `Table1` at write
  time, *after* the undo plan had to be written. The plan is recorded before the change (invariant
  2), so a change whose undo cannot be described in advance is a change we refuse — hence
  `tableName` is a required argument, checked with `checkTableName()` for the cases Office.js
  reports as "The argument is invalid or missing or has an incorrect format".

A created object's plan is also checked for collision before the write, because a `create_table`
that succeeds and *then* fails to take the requested name leaves a table this journal cannot find.

**Consequence.** `create_table` is the first tool outside the range model, and it is the template
for the rest: `create_chart` and `create_pivot_table` are new `kind`s with the same shape (delete by
name), `manage_sheet` needs one for sheet lifecycle, and `insert_delete_rows_cols` / `text_to_columns`
/ `autofilter` still need the *range* action widened or made span-aware. The rule the next tool has
to satisfy is now explicit: **if the change cannot be described before it is made, it cannot be
made.**

`checkTableName` and the plan-before-write ordering are covered by tests; the Office.js calls
themselves are not, and need the Diagnostics tab in real Excel.

---

## D-25 — The journal outlives the process, so an undo must prove it is still true
*2026-09-24*

**Context.** The journal was in memory, capped at 20 tasks, and died with the process. That made
the one safety net this app has evaporate on every restart — and by the time the app could make
real changes to a spreadsheet, it was the most likely way a user could actually lose work.

The obvious fix, writing it to a file, creates a worse bug than the one it solves. A file outlives
the workbook it describes. Close the app, edit those cells in Excel, reopen, press Undo: the
snapshot is now a description of a workbook that no longer exists, and restoring it destroys the
newer work while reporting success. In memory the window was seconds; on disk it is days.

**Decision.** Persist to `~/.ai-excel/journal.json`, and make every entry carry proof that it is
still applicable:

```ts
{ action, label, expect }        // expect = fingerprint of what the write left behind
```

`__fingerprint` is asked for immediately after a successful write, and `__restore` re-reads the
range and compares before writing anything. A mismatch refuses and says why. A *missing*
fingerprint refuses too — "cannot verify" is not the same as "verified", and the whole point of
invariant 2 is that an unundoable change is worse than a refused one.

Supporting choices:

- **Only the kinds that rewrite cell state ask for it** (`seal` on the snapshot result). Removing a
  table you created cannot overwrite data, so a stale-check there would protect nothing and would
  cost a round trip.
- **The fingerprint is 64 bits of FNV-1a, not a cryptographic hash.** `node:crypto` does not exist
  inside Excel's WebView, and this is a staleness check, not a security boundary. It is computed
  over the JSON of everything `__restore` would put back, so it covers formulas, number formats,
  font, fill and alignment — the same set the snapshot restores.
- **The file is written with write-then-rename.** A crash mid-write must leave the previous journal,
  not a half-file that reads back as "nothing to undo" — which would quietly turn a recoverable
  situation into a lost one.
- **A corrupt or missing file is logged and treated as empty.** Losing the ability to undo is bad;
  refusing to open the app is worse.
- **The store is generic over the record and never reads inside it**, and the server passes
  `{action, expect}` across the wire without the taskpane's `label`, which is for the user. Same
  rule as D-24: the knowledge of what an undo means lives in one place.
- **The journal describes the workbook, not the conversation**, so `reset` does not clear it and a
  restart does not lose it.
- **Test servers get their own journal file.** `startApp` creates a temp file per test, because the
  default is the user's real journal and a test writing there would destroy the very thing this
  change exists to protect. Same class of mistake as D-19's port.

**Consequence.** Undo now survives a restart, and — more valuable — it refuses when it would
otherwise overwrite newer edits, which was already possible in-session before this change. The cost
is one extra `__fingerprint` round trip per mutating call, and a new invariant (16) with three
tests. Journal entries written by an older build have no `expect` and will refuse rather than
restore; that is the safe direction, and the file did not exist before this change.

---

## D-26 — Excel only: LibreOffice and Google Sheets are dropped; Claude is last
*2026-09-24*

**Context.** The original intent was Excel "and its alternatives" — LibreOffice through a UNO
Python sidecar, Google Sheets through REST v4 + OAuth. Both were still on the board as §3 Reach,
and both were sized as "later".

**Decision.** Drop them. LibreOffice and Google Sheets are no longer targets and are not to be
re-added as small additions to a milestone. Excel for Windows stays, because it is the same
codebase and the same add-in rather than a second implementation. The Claude provider stays
planned but moves to the **end** of the work; its API key is added when we get there.

**Why dropping them is the cheap choice, not the expensive one.** The reasoning is not "less work
later" but "the work would not be shared". Everything that makes this app safe is tied to
Office.js semantics:

- **Undo.** `UndoAction` is a range snapshot plus object kinds, fingerprinted against cells. UNO
  and the Sheets API both need that model re-derived from scratch, and then two undo models exist
  to keep honest.
- **The safety layer.** Snapshot-before-write, the review gate, the staleness check and the
  fail-closed capability handshake are all properties of the executor channel. A second host is a
  second channel with its own failure modes.
- **Diagnostics.** The capability probe is Office.js-specific, and it is the only thing standing
  between a claim and a verified one. A second host needs its own, or the claims get weaker.
- **The tool catalog.** Sheets cannot create a pivot table at all, so the catalogs could never be
  the same shape — the model would need host-conditional instructions, and `tools.ts` would stop
  being the single source of truth (invariant 4).

So the cost of a second host is not "another API client". It is a second implementation of the
part that is already correct, for a smaller audience, with the two drifting apart.

**Consequence.** §3 Reach is now Excel for Windows and nothing else. The **code escape hatch loses
its runtime**: it was going to borrow the LibreOffice Python sidecar, so it now needs its own
subprocess decision before it can start — recorded in `TODO.md` §6 rather than left implicit. The
Claude provider moves to its own final section, and when it lands it will need two things this
codebase does not have yet: a place to say which provider to use (header control or config file),
and cost accounting that is not DeepSeek-only (`app/pricing.ts`).

---

## D-27 — API keys are set in the window, and a key only ever travels one way
*2026-09-24*

**Context.** The only way to set a key was `pbpaste | node app/setkey.ts`. That is fine for whoever
wrote the app and wrong for anyone else — including whoever wrote it, six months later, who has
forgotten the command. The app already had a window, a config file and a form of validation; what
it did not have was a way in.

**Decision.** `GET`/`POST /api/config`, and a settings dialog behind a gear in the header.

The difficult part is not the endpoint, it is what the endpoint is allowed to say. **A key travels
one way: up in a POST, never back.** `GET` returns `{set, masked, source}` and nothing else. If the
raw key ever came back it would be in the renderer's memory, in devtools, and in every screenshot
this project takes of its own UI, and the config file's `0600` would be protecting nothing. Two
consequences fall out of that rule and are easy to mistake for bugs:

- **The fields are always empty when the dialog opens**, and an empty field means "leave it alone".
  There is nothing to pre-fill them with — that is the point. Clearing a key is therefore its own
  button, not an empty submit.
- **`source` exists.** An environment variable beats the file, so saving a key while
  `DEEPSEEK_API_KEY` is set stores it and changes nothing. Without saying which value is in charge,
  the form would look broken at exactly the moment it is working correctly.

Supporting choices:

- **Validation is shared with the CLI** (`checkCredential`, `checkCredentialPatch` in `config.ts`).
  Two rule sets would drift, and the failure would be a key the CLI accepts and the dialog refuses.
  It catches the mistake that is actually made — pasting the command *as* the key — and the wrong
  provider, and nothing more. It does not prove a key works.
- **Saving swaps the live clients.** `ai` and `search` are reassignable in `createApp`'s closure,
  so the next message uses the new key. Requiring a restart would make the form feel broken.
- **The server only manages keys when it owns its model.** `startServer` disables the endpoint when
  a client was injected. `--smoke` passes a stub, so a test run cannot build a real keyed client
  from the user's config file — the same class of guard as the test port and the test journal.
- **A cross-origin POST is refused.** Loopback is reachable from any page in a browser. A request
  with no `Origin` is allowed on purpose: the tests and the CLI are not browsers.

**Consequence.** "Never send a key to the window" is invariant 17, with a test that saves a key and
asserts the raw value and a recognisable tail of it appear nowhere in the responses. What is *not*
done, and should be: nothing verifies a key against its provider, so a typo that passes the format
checks is only discovered on the first message. That is recorded in `TODO.md` rather than half-built
— the honest position is that "saved" means saved, not working, and the dialog says so.

---

## D-28 — The app sets up Excel itself, and binds only to loopback
*2026-09-24*

**Context.** First run was four manual steps: install a certificate with `npx office-addin-dev-certs`,
find `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef` (a path nobody would guess), copy
a manifest into it, restart Excel, then insert the add-in. Three of those are the app's job. While
checking what could be automated, something else turned up: **the server was bound to every
interface.** `server.listen(port)` with no host means `0.0.0.0`, verified with `lsof` — so the app
was reachable from the local network, while its certificate exception, its same-origin check on the
credentials endpoint and its "nothing leaves this machine" promise all assumed otherwise.

**Decision.**

1. **Bind loopback only — `127.0.0.1` and `::1`, and treat a busy port as fatal.** Both families
   because macOS resolves `localhost` to `::1` first, and a server on IPv4 loopback alone makes
   every client pay a failed connection first.
2. **The app generates its own certificate.** One local CA plus a `localhost` leaf, written to
   `~/.ai-excel/certs`, with the CA added to the login keychain. `openssl` does the X.509 because
   Node cannot create a certificate, only parse one — an external binary, not an npm dependency,
   and its absence is reported rather than guessed at.
3. **Setup runs silently at startup**, and the window shows only what is left.
4. **The manifest is written and compared, not blindly copied** — byte-for-byte, so an older build's
   manifest is replaced and an identical one is left alone.

Supporting choices, each of which cost something to learn:

- **A half-bound server is worse than no server.** Tolerating `EADDRINUSE` on one family let the
  smoke test bind `::1:3499` while a preview held `127.0.0.1:3499`, and because `localhost` resolves
  to `::1` first the desktop window was served a *stub model* by a server nobody had started on
  purpose. Only genuinely unavailable families (`EAFNOSUPPORT`, `EADDRNOTAVAIL`) are skipped now.
- **The recipe writes a `[req]` section into its openssl config.** Apple's LibreSSL refuses to run
  without `distinguished_name`; OpenSSL 3 does not care. Both are exercised.
- **Only DNS names can be validated in Node.** `X509Certificate.checkHost` implements the host case
  and not the IP one, so asking it about `127.0.0.1` returns undefined even when the SAN has it.
  `localhost` is the check, and the comment says so, because the obvious "fix" is to assert on IPs.
- **Two things are still manual, and are reported instead of faked.** Office gives no way for an
  outside process to insert an add-in, and Excel reads the `wef` folder only at startup. The app
  does everything else and says the two steps plainly.

**Consequence.** Two new invariants: loopback-only listening, and "setup touches the user's machine,
so never run it from a stub server" — `startServer` skips it when a client was injected, which is
what `--smoke` does, and `ensureCertificate` takes `keychain: false` for tests. That last guard
matters more than the others: the test suite runs on the machine whose keychain it would otherwise
be writing a root certificate into. The taskpane still has to be inserted once — after which
`Office.addin.setStartupBehavior(load)` can start it by itself, which is the next piece of work and
needs a shared runtime in the manifest.

---

## D-29 — The taskpane shares a runtime, and the manifest is validated by Microsoft's own tool
*2026-09-24*

**Context.** The last piece of "open the app and Excel, describe the task" was the bridge: the
taskpane had to be opened by hand and stayed only as long as its pane was open. Office offers three
ways to avoid that, and the docs are blunt about all three: they are **document-scoped**, and the
only application-scoped version ("Activate add-ins with events", `OnDocumentOpened`) requires a
Microsoft 365 admin to deploy the add-in from the tenant admin centre. That is a hosted,
administered setup, and this app is meant to be neither.

Running the manifest through Microsoft's own validator then turned up something better than a
design problem: **the manifest we had been shipping was invalid.** `<Requirements>` sat after
`<Permissions>`, where the schema does not allow it, and `<Version>` was `0.2.0.0`, below the schema's
1.0 floor. Excel had been loading it anyway — leniently, silently — so the file had never actually
been correct, and nothing said so.

**Decision.**

1. **Declare `SharedRuntime 1.1` and a long-lifetime `<Runtime resid="Taskpane.Url">`**, and call
   `Office.addin.setStartupBehavior(Office.StartupBehavior.load)` from the pane once Excel is ready.
2. **Fix the two schema errors**: `<Requirements>` moved between `<Hosts>` and `<DefaultSettings>`,
   and `<Version>` raised to `1.1.0.0` — also above the M0 spike's `1.0.0.0`, so the sideload is no
   longer a downgrade.
3. **Add `npm run validate:manifest`**, which renders the template the way the app will and hands it
   to `office-addin-manifest validate`.

What the shared runtime actually buys, stated precisely, because it is less than the phrase
"auto-opens" suggests:

- **Closing the task pane stops breaking the bridge.** Without a shared runtime the runtime dies with
  the pane; with one it lives until the document closes. This is the part that works immediately and
  needs no per-document setup.
- **Reopening a workbook where the add-in has been used starts the bridge by itself**, with no pane.
- **It does not remove the first manual step.** The add-in has to be inserted once in each workbook,
  because Office gives no way for an outside process to do it — and `setStartupBehavior` explicitly
  does not apply on the first open of a document. Claiming otherwise would be the sort of promise
  this project keeps having to walk back.

**Why the validator is a script and not a test.** `office-addin-manifest` pulls its own tool copy on
first use and its acceptance step talks to Microsoft, and the test suite is offline by design. So the
suite keeps three cheap offline checks for the failures that would actually ship — element order, the
version floor, and a `<Runtime resid>` that no `<bt:Url>` defines — and `npm run validate:manifest`
does the real schema validation on demand.

**Consequence.** The taskpane now reports its auto-start state on the Connection tab, worded around
"from the next time it opens", because that is when Office applies it. One trap is recorded in the
test itself: the ordering check has to strip XML comments first, because the manifest's own header
mentions the element names in prose — the first version of that test passed for the wrong reason
after reading a comment. `TODO.md` §3 keeps the Windows sideload path, which is a different
mechanism again.

---

## D-30 — A certificate in the keychain is not a trusted certificate
*2026-09-26*

**Context.** The first real run of the automatic setup broke Excel. Two symptoms, and the second is
the one that matters:

1. The setup panel said the add-in could not be installed: `EPERM: operation not permitted, open
   '…/wef/manifest.xml'`.
2. Opening the add-in in Excel said **"The content is blocked because it isn't signed by a valid
   security certificate."**

**Diagnosis.**

The EPERM was **not** macOS and **not** the app. A probe running the same read and write inside
Electron succeeded, and the installed manifest's mtime was still from the earlier manual copy — so
something environmental had blocked both. The app had been launched inside a filesystem sandbox that
allowed the workspace but not Excel's container. Nothing to fix in the app beyond saying so
clearly.

The certificate was a real bug, and the docs caused it. D-5 and D-28 both assert that Office's
WebView accepts a CA merely *present* in the login keychain — "no explicit trust settings needed".
That was inferred from `security dump-trust-settings` reporting no user trust settings, which is a
statement about a *domain*, not about a certificate. The check that actually answers the question is:

```
security verify-cert -p basic -c <cert>
```

The Office dev CA: `certificate verification successful`.
Our generated CA: `CSSMERR_TP_NOT_TRUSTED`.

So the dev CA **is** trusted, it merely lives somewhere `dump-trust-settings` does not enumerate, and
Office accepts it for the ordinary reason. Our CA was never trusted, and Office refused it for the
ordinary reason too. `-p ssl` cannot see this: it demands Certificate Transparency timestamps that
no locally-issued certificate will ever have, so it fails on both and distinguishes nothing.

**Decision.**

1. **Trust the CA, don't just install it.** `security add-trusted-cert -k ~/Library/Keychains/login.keychain-db`
   needs no admin password and no prompt (verified), so this stays a silent step.
2. **Prefer a certificate macOS actually trusts.** `resolveTls` checks ours, then the Office dev
   certs, and only falls back to an untrusted one so the app still starts.
3. **Serve the chain, not just the leaf**, and give the leaf `subjectKeyIdentifier` /
   `authorityKeyIdentifier`.
4. **Replace rather than re-add on regeneration.** A new CA has the same common name as the one it
   replaces, so "already present" would leave the *old* certificate trusted while serving the new
   one — a mismatch that presents as exactly the same error.
5. **Say so when it is not trusted.** The setup panel and `/api/setup` now report the trust state,
   because the symptom otherwise is "Excel not connected", which sends you looking at the add-in.
6. **Distinguish "cannot read the folder" from "no add-in installed."** The old message blamed the
   add-in for what was a permission error.

**Consequence.** Trusting a CA is a bigger act than storing one, so it is written down: it is our
own CA, its private key is `0600` in a directory the app owns, and `security remove-trusted-cert`
undoes it. Correcting D-5's explanation does not change the certificate exception in
`cert-policy.ts` — Chromium still refuses a certificate WebKit accepts — but the reason recorded for
it was wrong, and a wrong reason in a decision log is how this bug happened. Invariant 20 now holds
the check, and three tests cover the distinction.

---

## D-31 — Reworked: one self-signed certificate, setup outside the server, no shared runtime
*2026-09-27*

**Context.** D-28 to D-30 built the automation, and it worked — but it broke the user's Excel twice,
and the second time the cause was a claim in these very documents. The user asked for it to be redone
properly, for the shortest path to Excel just working.

**What was wrong, in order of how much it cost.**

1. **The CA asked for more trust than it needed.** Trusting a *CA* grants it the power to vouch for
   **any** hostname; trusting a *single leaf* grants it power over `localhost` only. Measured first:
   a self-signed leaf verifies on macOS after `add-trusted-cert`, so the CA bought nothing.
2. **Setup lived inside `startServer`.** Machine-mutating work in a factory that tests, the CLI and
   the preview harness all call is the wrong boundary — it needed invariant 19 purely to guard
   against itself.
3. **The shared runtime cost more than it gave.** Its auto-open is *document-scoped* and does not
   apply to a document's first open, so it never removed the manual step it existed for. It also
   means a manifest declaring a requirement set the client may not honour, which is a real way to
   stop an add-in loading. Dropped.
4. **`add-trusted-cert` prompts for a password**, measured: it blocks indefinitely, twice, at 12s
   each, while `add-certificates` and `delete-certificate` return in milliseconds. macOS offers no
   non-interactive way to change a user trust setting. The old design called it during startup, so
   the app could hang behind a dialog nobody asked for.

**Decision.**

- **One self-signed `localhost` leaf**, trusted directly. `ca.crt`/`ca.key` are deleted if found.
- **Setup moved to the app's entry point**, before the server exists. The server now only *reports*
  (`GET /api/setup`; anything else is 405) and physically cannot change the machine — which is why
  invariant 19 is no longer needed as a guard.
- **The shared runtime is gone**; the manifest is a plain taskpane again, keeping only the two
  schema fixes from D-29. The pane stays open, which is one instruction and no risk.
- **Trust is never granted automatically.** The app generates and *stores* the certificate
  silently, then reports whether macOS trusts it and prints the one command
  (`npm run trust-cert`). One password prompt, once, with the reason on screen — and no way for
  startup to block.
- **A certificate that is not self-signed is rejected**, not served: that is how the retired
  CA-signed leaf was caught, and `add-trusted-cert`'s failure for it names neither the cause nor the
  certificate.

**Consequence.** `tls.ts` is a third smaller and trusts strictly less. The server has no ability to
mutate the machine, so the whole class of "a test server touched the user's setup" is designed out
rather than guarded against. The honest cost is that first-run setup is now **two** manual actions —
run `npm run trust-cert`, and insert the add-in once per workbook — both of which macOS or Office
require, and both of which the panel states with the exact command.

Not verified: that Excel loads the add-in. The certificate is trusted-checkable and the manifest is
schema-valid, but only Excel can answer the last question.

---

## D-32 — Undo an object by deleting it only when the object owns no cells
*2026-09-27*

**Context.** `create_chart` and `create_pivot_table` were added to the catalog. The undo layer was
already a union of typed actions ([D-24](docs/DECISIONS.md)), so each needed one new `kind` — but
choosing what each `kind` *does* turned on a distinction that is easy to get backwards.

**Decision.** The three created-object kinds do not share an undo:

| Kind | Undo | Why |
|---|---|---|
| `created-table` | `convertToRange()` | The cells held that data *before* the table existed. `delete()` would take the data with it. |
| `created-chart` | `delete()` | A chart floats over the cells and owns none of them, so nothing is lost. |
| `created-pivot` | `delete()` | The destination cells are the pivot's own computed output, recomputed from a source that is left alone. |

The rule underneath: **an object may be deleted only if it owns the cells it appears to occupy.** A
table is a view over cells the user put there; a chart is a drawing; a pivot's cells are derived. Get
this wrong in either direction and the failure is silent — `delete()` a table and undo has destroyed
data, while `convertToRange()` a chart and undo has changed nothing at all while reporting success.

**Also.** Both new planners look their object up with `getItemOrNullObject` before acting, so an
object the user deleted themselves reads as "already gone" instead of throwing. `undoCreatedTable`
was rewritten to match: it had loaded every table and searched by hand, because a `grep … | head -5`
made it look as though the method did not exist on `TableCollection`. It does. The three paths are
now identical, and the Diagnostics check for the table was updated to verify the same call the code
makes — a check that exercises a different mechanism than the one that ships is worse than none.

**Not verified.** Office.js behaviour needs Excel. The chart and pivot *creation* calls were already
covered by existing Diagnostics checks, but the two new undo paths have new checks
(*Undo a created chart*, *Undo a created pivot*) which have not been run.

---

## D-33 — `remove_duplicates` is undone by the range snapshot, and that rests on one measurement
*2026-10-06*

**Context.** `remove_duplicates` was the one tool held back for a *measurement* rather than for
missing code. Its undo is the ordinary range snapshot, which is only honest if the rows that move
when duplicates are deleted come from **inside** the range being snapshotted. If Excel pulled a row
up from below it, the snapshot would restore the wrong cells and the journal would claim a restore
it did not perform.

**Decision.** The Diagnostics check *Remove-duplicates shift scope* reported **CONFINED**: "a row
below the range was untouched, so a range snapshot can undo it." The tool was added on that basis —
it needed no new undo kind, no `targetAddressFor` case (that function already falls back for any
tool with an `address`), and no change to the journal.

**The part worth remembering.** Nothing in the offline suite can prove this, and nothing refuses if
it stops being true. Every other undo path asserts its own precondition: a range restore compares a
fingerprint and refuses on a mismatch, and the created-object kinds are lookups that read as
"already gone". This one trusts a measurement taken once, on one Excel build. If a future version
shifts rows from outside the range, the undo becomes a silent lie — it would report success and
write back cells that were never in the range.

That is recorded as an open item in §0 of [TODO.md](../TODO.md) rather than left implicit: either
the check becomes authoritative for advertising the tool, or the tool goes. It is acceptable while
this runs on one machine and is not acceptable once it does not.

**Also measured the same day.** Excel 16.113.920.2, 21 checks, nothing failed — including the two
new undo paths from [D-32](DECISIONS.md) (`Undo a created chart`, `Undo a created pivot`), which are
the only evidence that exists for them.

---

## D-34 — An expired approval is a silence, not a refusal
*2026-10-06*

**Context.** The approval card disappears after two minutes and the turn carries on without the
change. The complaint was cosmetic — the card "vanishes with no explanation" — and it was written up
in TODO as a wording problem.

**It was not a wording problem.** On expiry the server pushed this into the conversation:

> The user declined this change. Do not retry it; ask what they would prefer instead.

The model then did exactly what that says: stopped, and deferred to a user who had never seen the
card, usually because they had walked away from the machine. A user who *declines* wants the model
to stop; a user who is *absent* wants it to wait or ask again. Both events arrived as one word, so
the model could not tell them apart.

**Decision.** The two outcomes are distinguished everywhere they surface:

- **To the model**, which is the one that changes behaviour: an expiry says the request expired with
  no answer and nobody declined it, and asks it to offer to retry; only an explicit decline says
  "the user declined".
- **To the window**: "expired with no answer, so nothing was written (…)" rather than the internal
  word "timeout", which describes the mechanism rather than the event.

**And the reason it went unnoticed:** the timeout was a module-level constant of 120 seconds, so no
test could reach the path — a second test would have cost two minutes. It is now read per call from
`AI_EXCEL_APPROVAL_TIMEOUT_MS`, which makes expiry testable in 40 milliseconds and is separately
useful for a person who steps away from a long turn. The rule is invariant 22.

The general shape, worth recognising again: **a path that cannot be exercised is a path that is
not known to work.** This one had been wrong since it was written, and its bug was not in the
feature — it was in the sentence the feature sent to the model.
