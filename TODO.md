# TODO

Everything still outstanding, as checkboxes. Ordered by what I would do first.

- `- [x]` done and verified
- `- [ ]` not done
- `- [x] ~~…~~` considered and closed — either done and later replaced, or dropped (§7)
- **(needs you)** means it cannot be done without Excel, your eyes, or a decision

Keep this file current: tick items as they land and add new ones as they appear. It is referenced
from [AGENTS.md](AGENTS.md).

---

## 0. Verify before building anything else

- [ ] **(needs you)** **Confirm your workbook is unchanged.** A stray test of mine started a server
      on port 3443 with your real API key, and your Excel taskpane connected to it and ran a turn.
      The request was read-only (`what's in this workbook?`) and **Review** was on, so nothing should
      have been written — but only you can check. See
      [D-19](docs/DECISIONS.md) and the "Verifying your work" section of AGENTS.md.
- [x] **Run Diagnostics once more, now that `executor.ts` has new Office.js code.** Done 2026-10-06
      on Excel 16.113.920.2 (ExcelApi through 1.21) — **21 checks, everything PASS, nothing FAILED**.
      Both questions this was gating are answered:
      - *Remove-duplicates shift scope* → **CONFINED**: "a row below the range was untouched, so a
        range snapshot can undo it". `remove_duplicates` was added on the strength of it (§2).
      - *Undo a created table* → **PASS**, and it now exercises `getItemOrNullObject`, the call the
        shipped code actually makes.
      - The two new undo paths also pass: *Undo a created chart* and *Undo a created pivot*. That is
        the only proof that exists for those — nothing here can run Office.js.
- [ ] **The shift-scope result is reported, not enforced.** `remove_duplicates`'s undo is honest
      *because* that check says CONFINED. If a future Excel version changes its behaviour, the
      journal for that tool silently becomes a lie and nothing refuses. Either make the check
      authoritative when the tool is advertised, or drop the tool. Worth doing before this leaves
      one machine.
- [x] **Re-sideload the manifest.** Confirmed 2026-10-06: Excel loaded the current taskpane, since
      the report above contains the checks added in the latest build. The old "AI Spreadsheet Pilot"
      name is gone. The app installs this itself now, so there is no manual step left.
- [x] **The manifest `<Version>` was below the schema's floor.** Fixed: it is now `1.1.0.0`, above
      both the minimum Office requires (1.0) and the M0 spike's `1.0.0.0`, so the sideload is no
      longer a downgrade. Microsoft's validator caught it — see [D-29](docs/DECISIONS.md).
- [ ] **It still drifts from `package.json`** (1.1.0.0 vs 0.4.0). Two version numbers that move
      independently will disagree; decide whether the app version should follow the add-in or the
      other way round.

---

## 1. Safety and correctness

The undo/review layer is built; these are the holes left in it.

- [x] **Persist the undo journal to disk.** `~/.ai-excel/journal.json`, write-then-rename, 0600,
      reloaded on startup. Because a file outlives the workbook it describes, every entry now
      carries a fingerprint of what its change left behind and undo refuses when the cells have
      moved on — see [D-25](docs/DECISIONS.md) and invariant 16.
- [ ] **An audit log of applied changes** — what was written, when, by which task. The journal is
      disposable by design; an audit trail is not the same thing and should survive.
- [x] **Tell the user when an approval times out.** Fixed 2026-10-06, and the fix was bigger than
      the wording. The window now says "expired with no answer, so nothing was written (…)" instead
      of "not applied (timeout)" — but the real defect was in what the *model* was told: an expired
      approval sent it "The user declined this change. Do not retry it; ask what they would prefer
      instead." So the model stopped and deferred to a user who had never seen the card. Silence and
      refusal are different events ([D-34](docs/DECISIONS.md), invariant 22), and
      `AI_EXCEL_APPROVAL_TIMEOUT_MS` now makes the expiry path testable at all.
- [ ] **Guard against the wrong workbook.** Session `default` is shared, so with two workbooks open
      whichever taskpane is polling wins the call. Either key sessions to the workbook or refuse to
      connect a second taskpane. (A *stale* undo is now caught by the fingerprint, but two live
      taskpanes racing is still not — and this is the one hole left that can put a change in a file
      the user did not ask for.)
      **Waiting on a measurement.** A *Workbook identity* Diagnostics check now reports the two
      candidate identifiers: `document.url` (empty for an unsaved workbook) and a round-trip through
      `document.settings`, which could hold a generated id that survives a reload — if Excel
      persists it. Which one is usable decides the design, so run that check before building this.
- [ ] **Cap snapshot size.** A very large write produces a correspondingly large snapshot (~29
      bytes/cell). Decide what happens past some threshold instead of finding out.
- [ ] **A confirm step for destructive ranges.** `write_range` over an occupied range is currently
      treated exactly like writing into empty space.
- [ ] **Prune the journal file.** Sessions are keyed by id and never removed, so a long-lived
      install accumulates one entry list per session. Bounded by 20 turns each and by how many
      session ids are ever used, but nothing ever deletes them.

---

## 2. Capability — the remaining Excel tools

All are *proven working* by the Diagnostics panel; none are exposed to the model yet. **What
blocks each one is undo coverage, not wiring.** A tool whose change the journal cannot reverse
would make "Undo last task" a lie, and invariant 2 says an unundoable change is worse than a
refused one.

**Undoable with today's range snapshot** — just add the tool, a `targetAddressFor` case and a test:

- [x] `sort_range` — multi-key sort, values permuted inside the range
- [x] `remove_duplicates` — **measured CONFINED** by Diagnostics on 2026-10-06, so the range
      snapshot genuinely covers the row shift. Added the same day. Column indices are validated
      against the range before calling Excel, whose own error for an out-of-range index names
      nothing. The tool is only as honest as that measurement — see the note in §0.

**Needs the range action widened** — the change escapes the address it was given:

- [ ] `text_to_columns` — writes into columns *beyond* the given address, so the capture must cover
      the destination span too
- [ ] `autofilter` — filtering hides rows; hidden state is not part of a range snapshot
- [ ] `insert_delete_rows_cols` — shifts everything below or right, which a range snapshot cannot
      put back
- [ ] `set_conditional_format` — the rules live on the range but are not captured today, so the
      capture has to take them and the restore has to put them back

**Needs object-level undo** — these create things, not cell values:

- [x] `create_table` — undo is `convertToRange()`, not `delete()`, so the cells under the table
      stay. The name is a **required argument** so the undo plan can be written down before the
      table exists (D-24).
- [x] `create_chart` — undo deletes the chart by name (`getItemOrNullObject`, then `delete()`).
      Same shape as `create_table`: the model supplies the name. **Added 2026-09-27**; the Office.js
      calls it uses were already proven by Diagnostics, but the *undo* path is new and has its own
      check now — run it.
- [x] `create_pivot_table` — undo deletes the pivot by name. Fields are addressed by source header
      name, so a hallucinated field fails loudly instead of producing an empty pivot. Added
      2026-09-27, same caveat: the undo path needs the new Diagnostics check.
- [ ] `manage_sheet` — add/rename are simple; **deleting a sheet is not range-undoable at all**
      and probably should require explicit confirmation rather than pretend

The journal is now a union of typed undo actions, and `create_table` is the worked example:

```ts
type UndoAction =
  | { kind: 'range';         sheet; address; state }
  | { kind: 'created-table'; sheet; name }
  | { kind: 'created-chart'; sheet; name }   // added
  | { kind: 'created-pivot'; sheet; name }   // added
  // sheet lifecycle still slots in here
```

`__restore` dispatches on `kind`, and the server stores the action without ever reading inside it,
so the union lives in one place (`app/web/executor.ts`). That is what makes the chart and pivot
tools small: a new kind is a taskpane-only change.

- [x] ~~**Re-check the system prompt for tool-choice guidance.**~~ The prompt deliberately does not
      enumerate tools — the schemas are the only description the model gets, so there is nothing to
      drift out of step. It gained no chart-specific rules on purpose: "read before you write" and
      "prefer one call" already cover when to reach for them.

---

## 3. Reach

**This is an Excel product.** The alternatives were dropped deliberately — see §7.

- [x] **Automate the Excel setup** rather than documenting it. The app generates its own local
      certificate, adds the CA to the login keychain, writes the add-in manifest into Excel's folder,
      and shows a panel for the two steps Office will not let it do — restart Excel, and insert the
      add-in once. See [D-28](docs/DECISIONS.md).
- [x] **Local-only was not actually true.** `server.listen(port)` bound every interface, so the app
      was reachable from the local network — while the certificate exception, the origin check on
      the credentials endpoint and the "nothing leaves this machine" promise all assumed otherwise.
      It now binds `127.0.0.1` + `::1` and nothing else, and a busy port is fatal (invariant 18).
- [x] **Make the taskpane start itself.** Shared runtime declared (`SharedRuntime 1.1` plus a
      long-lifetime `<Runtime>`), and the pane calls `Office.addin.setStartupBehavior(load)`. Two
      limits are Office's rather than ours and are stated instead of glossed: it is **document
      scope**, and it does not apply on the first open of a document. What it does give: closing the
      pane no longer disconnects the bridge, and reopening a workbook where it has been used starts
      it by itself. See [D-29](docs/DECISIONS.md).
- [x] **The manifest was invalid and nobody knew.** Microsoft's own validator found
      `<Requirements>` in a position the schema rejects and a `<Version>` below the 1.0 minimum —
      Excel had been loading it leniently the whole time. Both fixed, `<Version>` is now `1.1.0.0`,
      and `npm run validate:manifest` runs the validator on the rendered template.
- [ ] **Excel for Windows.** The same add-in should work; it needs the sideload path
      (`%APPDATA%\Microsoft\AddIns`, or the `WEF\Developer` registry key, which needs no admin) and
      real testing. `setup.ts` is macOS-only today and says so rather than guessing.

---

## 4. Product

- [x] **Set API keys from the window.** A settings dialog behind a gear in the header, with
      `GET`/`POST /api/config`, validation shared with the CLI, and a live client swap so a new key
      works on the next message. A key travels one way only — see [D-27](docs/DECISIONS.md) and
      invariant 17. The CLI stays for scripts and for the first run before a window exists.
- [ ] **Verify a key against its provider before calling it working.** Today "Saved" means saved,
      not working: the format checks catch a pasted command and a key for the wrong provider, and
      only the first message catches a typo. DeepSeek has a free balance endpoint that would settle
      it; Tavily and Brave would cost a request each. Deliberately not half-built — see D-27.
- [ ] **Package as a real app** (.dmg / .msi). Today it is `npm run desktop` from a source checkout,
      which is fine for you and wrong for anyone else.
- [ ] **Restore the conversation on window reload.** The server keeps history; a reloaded window
      shows a blank transcript while the model still remembers.
- [x] **Re-enable web search.** A Tavily key was added 2026-09-24. The usage panel now counts
      searches, so a search-heavy session does not look free.
- [x] **Web search is off by default,** behind a header toggle. A configured key is not consent to
      use it, and a search the model attempts while the toggle is off is refused without contacting
      the provider.
- [x] ~~Use DeepSeek's own native web search instead~~ — **tested 2026-09-24 and unavailable.**
      `tools:[{type:'web_search'}]` is ignored on the current model, despite two August 2026 issues
      saying it worked on `deepseek-v4-flash`. See [D-20](docs/DECISIONS.md).
- [ ] **Re-run `node scripts/probe-search.js`** if DeepSeek announces native search. It costs about
      $0.0005 and settles whether we can drop the third-party key.

---

## 5. Cleanup and debt

- [x] **Remove or use `ChatEntry.streaming`.** Removed — nothing rendered it.
- [x] **Decide the fate of `/api/activity`.** Removed. The stream supersedes it, and a second,
      staler source of truth for "what happened" is a liability. Its coverage moved to stream
      assertions.
- [x] **Get generated `.js` out of the source tree.** All output now goes to `dist/`.
- [ ] **`deepseek-chat` is an unpublished model name.** Switching the default to `deepseek-flash`
      would make the cost figure exact rather than an estimate. It changes which model answers, so
      it is your call — one line in `app/ai.ts` plus a re-test.
- [ ] **Make the pricing table updatable without a code change** if it drifts again
      (`app/pricing.ts` is the single place today).

---

## 6. Deferred, deliberately

- [ ] **Code escape hatch** — AI-generated code for tasks the tool catalog cannot express, shown
      for approval and run in a subprocess. Deferred until the fixed tools demonstrably fall short.
      **Its runtime went with LibreOffice** (§7): there is no longer a bundled Python sidecar to
      borrow, so this now needs its own subprocess decision before it can start.
- [ ] **Multi-workbook sessions** — see the safety item above; the fix is the same.
- [ ] **Local models** — a provider that never sends sheet contents off the machine.

---

## 7. Dropped, deliberately

Decided 2026-09-24. Recorded here so they are not silently re-proposed, and not "just one more
thing" at the end of a milestone. See [D-26](docs/DECISIONS.md) for the reasoning.

- [x] ~~**LibreOffice** — UNO via a bundled Python sidecar.~~ **Dropped.** The add-in path already
      covers macOS and Windows Excel with one codebase; a UNO sidecar would be a second, parallel
      implementation of every tool, every undo action and every diagnostic for a smaller audience.
- [x] ~~**Google Sheets** — REST v4 with OAuth.~~ **Dropped.** Same reasoning, plus the API cannot
      create a pivot table at all, so the tool catalog would be permanently unlike the Excel one.

Dropping them also removes a class of bug: the undo model (typed actions, range snapshots,
fingerprints) is built around Office.js semantics, and a second host would need it re-derived
rather than reused.

---

## 8. Last — the Claude provider

- [ ] **Claude provider** behind the same `AiClient` interface. `ai.ts` is already shaped for it:
      one interface, one implementation per vendor. **The API key is added at the end**, so this is
      the final piece of work, not something to start early. Provider selection will then need a
      decision about where it lives (header control vs config file) and how cost accounting copes
      with two vendors' price tables (`app/pricing.ts` is DeepSeek-only today).

---

## Done

For orientation. Not exhaustive — the reasoning for each is in [docs/DECISIONS.md](docs/DECISIONS.md).

- [x] Office.js add-in driving installed Excel, verified against ExcelApi 1.21
- [x] Electron + React + shadcn/ui window, outside Excel
- [x] Server-mediated executor channel; the taskpane is a silent worker
- [x] DeepSeek tool loop with server-side web search
- [x] Snapshot before every write, one-click undo, review gate
- [x] Token, speed and cost accounting with honest estimate labelling
- [x] Realtime streamed replies with inline tool activity
- [x] TypeScript throughout, compiled in place
- [x] `sort_range` tool, snapshotted and undoable
- [x] `create_chart` and `create_pivot_table`, undoable by deleting the object by name (D-32)
- [x] `remove_duplicates`, undoable by the range snapshot (verified CONFINED in Excel)
- [x] `create_table` tool, with undo as a typed action rather than a cell snapshot (D-24)
- [x] Undo journal persisted to disk, with a staleness check before every restore (D-25)
- [x] API keys set from the window, with the raw key never crossing back (D-27)
- [x] 67 self-tests, a preview harness, and a screenshot-verifying smoke test
- [x] Agent documentation set (AGENTS.md, ARCHITECTURE.md, DECISIONS.md)
