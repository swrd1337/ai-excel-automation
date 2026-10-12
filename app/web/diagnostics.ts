/*
 * Diagnostics — the M0 capability probe, kept as a permanent panel.
 *
 * It earned its place twice already (it caught two real bugs in its own checks,
 * and established the ExcelApi ceiling and the bridge behaviour). Anything the
 * app depends on should be provable on the user's own machine, on demand.
 *
 * Writes ONLY to a scratch sheet named _AI_SPIKE.
 */

const SCRATCH = '_AI_SPIKE';

const CORE = [
  ['region', 'units', 'price'],
  ['North', 10, 2.5],
  ['South', 20, 3.75],
];

interface Check {
  name: string;
  /** ExcelApi minor version this needs, e.g. '1.6'. Omit for checks that do not use one. */
  minSet?: string;
  info?: boolean;
  run: () => Promise<string>;
}

const j = (v: unknown): string => JSON.stringify(v);

function colName(n: number): string {
  let s = '';
  for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function fmtError(err: unknown): string {
  if (!err) return 'unknown error';
  const e = err as { code?: string; message?: string; debugInfo?: unknown };
  const code = e.code ? `${e.code}: ` : '';
  const msg = e.message || String(err);
  const dbg = e.debugInfo
    ? ` [${typeof e.debugInfo === 'string' ? e.debugInfo : JSON.stringify(e.debugInfo)}]`
    : '';
  return `${code}${msg}${dbg}`;
}

function assert(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}

async function scratch<T>(
  fn: (context: Excel.RequestContext, sheet: Excel.Worksheet) => Promise<T> | T
): Promise<T> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(SCRATCH);
    return fn(context, sheet);
  });
}

function supportedSets(): string[] {
  const out: string[] = [];
  for (let minor = 1; minor <= 21; minor++) {
    if (Office.context.requirements.isSetSupported('ExcelApi', `1.${minor}`)) out.push(`1.${minor}`);
  }
  if (Office.context.requirements.isSetSupported('ExcelApiOnline', '1.1')) out.push('Online 1.1');
  return out;
}

const CHECKS: Check[] = [
  {
    name: 'Host, platform and requirement sets',
    info: true,
    run: async () => {
      // `build` is present at runtime but absent from the published type.
      const d = Office.context.diagnostics as { version?: string; build?: string } | undefined;
      return [
        `host=${Office.context.host}`,
        `platform=${Office.context.platform}`,
        d?.version ? `version=${d.version}` : null,
        d?.build ? `build=${d.build}` : null,
        `ExcelApi=${supportedSets().join(' ') || 'NONE'}`,
      ]
        .filter(Boolean)
        .join('  ');
    },
  },
  {
    name: 'Workbook identity (telling two open workbooks apart)',
    info: true,
    minSet: '1.1',
    run: async () => {
      /*
       * This decides whether the app can refuse to run a task against the wrong
       * workbook. Session 'default' is shared, so with two workbooks open
       * whichever taskpane polls wins the call — and a task the user aimed at one
       * file can land in the other. A stale undo is caught by the journal
       * fingerprint; two live taskpanes racing is not caught by anything.
       *
       * Reported rather than assumed, because both candidates are unverified on
       * this platform: `document.url` is empty for a workbook that has never been
       * saved, and `document.settings` is the per-document store that could hold a
       * generated id that survives a reload — if Excel persists it.
       */
      const url = Office.context.document.url || '';
      const notes = [url ? `url=${url}` : 'url=EMPTY (unsaved workbook, so url cannot identify it)'];

      let settings = 'settings=UNAVAILABLE';
      try {
        const store = Office.context.document.settings;
        const key = 'aiExcelIdentityProbe';
        const before = store.get(key) as string | undefined;
        store.set(key, 'probe');
        await new Promise<void>((resolve, reject) => {
          store.saveAsync((result) => (result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(String(result.error?.message || result.status)))));
        });
        const after = store.get(key) as string | undefined;
        store.remove(key);
        settings = `settings=${after === 'probe' ? 'WRITABLE (round-trip OK)' : 'read back wrong'}${before ? ' [had a prior value]' : ''}`;
      } catch (err) {
        settings = `settings=FAILED (${err instanceof Error ? err.message : String(err)})`;
      }

      return `${notes[0]}  ${settings}`;
    },
  },
  {
    name: 'Write a 2D range and read it back intact',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        const r = sheet.getRange('A1:C3');
        r.values = CORE;
        r.load('values');
        await context.sync();
        assert(j(r.values) === j(CORE), `round-trip mismatch: got ${j(r.values)}`);
        return '3x3 matrix round-tripped byte-identical';
      }),
  },
  {
    name: 'Set a formula and read its computed value',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        sheet.getRange('E1').formulas = [['=SUM(B2:B3)']];
        await context.sync();
        const r = sheet.getRange('E1');
        r.load(['values', 'formulas']);
        await context.sync();
        assert(Number(r.values[0]?.[0]) === 30, `expected 30, got ${j(r.values[0]?.[0])}`);
        return `=SUM(B2:B3) -> ${r.values[0]?.[0]}`;
      }),
  },
  {
    name: 'Number format',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        sheet.getRange('C2:C3').numberFormat = [['0.00%'], ['0.00%']];
        await context.sync();
        const r = sheet.getRange('C2:C3');
        r.load('numberFormat');
        await context.sync();
        assert(r.numberFormat[0]?.[0] === '0.00%', `expected 0.00%, got ${j(r.numberFormat)}`);
        return `numberFormat = ${r.numberFormat[0]?.[0]}`;
      }),
  },
  {
    name: 'Formatting: font, fill colour, border weight',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        const r = sheet.getRange('A1:C1');
        r.format.font.bold = true;
        r.format.font.color = '#FFFFFF';
        r.format.fill.color = '#1F6FEB';
        const bottom = r.format.borders.getItem('EdgeBottom');
        bottom.style = Excel.BorderLineStyle.continuous;
        bottom.weight = Excel.BorderWeight.medium;
        await context.sync();
        const check = sheet.getRange('A1');
        check.load(['format/font/bold', 'format/fill/color']);
        await context.sync();
        assert(check.format.font.bold === true, 'bold not applied');
        assert(
          String(check.format.fill.color).toUpperCase().endsWith('1F6FEB'),
          `fill not applied: ${check.format.fill.color}`
        );
        return `bold=${check.format.font.bold} fill=${check.format.fill.color}`;
      }),
  },
  {
    name: 'Column width and autofit',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        const cols = sheet.getRange('A:C');
        cols.format.columnWidth = 120;
        await context.sync();
        cols.load('format/columnWidth');
        await context.sync();
        const set = cols.format.columnWidth;
        assert(set === 120, `expected 120, got ${set}`);
        cols.format.autofitColumns();
        await context.sync();
        // Read a single column: columnWidth on a multi-column range is meaningless after autofit.
        const colA = sheet.getRange('A:A');
        colA.load('format/columnWidth');
        await context.sync();
        return `set 120pt (read back ${set}pt), autofitColumns() -> A = ${Math.round(colA.format.columnWidth)}pt`;
      }),
  },
  {
    name: 'Conditional formatting (cell-value rule)',
    minSet: '1.6',
    run: () =>
      scratch(async (context, sheet) => {
        const r = sheet.getRange('B2:B3');
        const cf = r.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        cf.cellValue.format.fill.color = '#FFC7CE';
        cf.cellValue.rule = { formula1: '=15', operator: Excel.ConditionalCellValueOperator.greaterThan };
        await context.sync();
        const again = sheet.getRange('B2:B3').conditionalFormats;
        const count = again.getCount(); // not a `count` property — this collection only has getCount()
        await context.sync();
        assert(count.value === 1, `expected 1 rule, got ${count.value}`);
        return 'cellValue > 15 rule added and persisted';
      }),
  },
  {
    name: 'Sort a range by one column',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        sheet.getRange('A6:B9').values = [
          ['item', 'qty'],
          ['a', 5],
          ['b', 20],
          ['c', 12],
        ];
        await context.sync();
        sheet.getRange('A7:B9').sort.apply([{ key: 1, ascending: false }], false, false);
        await context.sync();
        const r = sheet.getRange('A7:A9');
        r.load('values');
        await context.sync();
        const got = r.values.flat();
        assert(j(got) === j(['b', 'c', 'a']), `expected b,c,a got ${j(got)}`);
        return `sorted desc by qty -> ${got.join(', ')}`;
      }),
  },
  {
    name: 'AutoFilter apply and read back',
    minSet: '1.9',
    run: () =>
      scratch(async (context, sheet) => {
        sheet.autoFilter.apply(sheet.getRange('A6:B9'));
        const af = sheet.autoFilter;
        const applied = af.getRange(); // not a `range` property — AutoFilter exposes getRange()
        applied.load('address');
        af.load('isDataFiltered');
        await context.sync();
        return `applied on ${applied.address} (isDataFiltered=${af.isDataFiltered}; apply(range) attaches no criteria)`;
      }),
  },
  {
    name: 'Create a Table (ListObject) with headers',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        sheet.getRange('A12:B15').values = [
          ['sku', 'stock'],
          ['x1', 4],
          ['x2', 9],
          ['x3', 0],
        ];
        await context.sync();
        const t = sheet.tables.add('A12:B15', true);
        t.name = 'SpikeTable';
        t.load(['name']);
        await context.sync();
        assert(t.name === 'SpikeTable', `table name was ${t.name}`);
        return `table "${t.name}" created over A12:B15`;
      }),
  },
  {
    name: 'Undo a created table (convertToRange keeps the cells)',
    minSet: '1.2',
    run: () =>
      scratch(async (context, sheet) => {
        const rows = [
          ['sku', 'stock'],
          ['y1', 2],
          ['y2', 7],
          ['y3', 5],
        ];
        sheet.getRange('D12:E15').values = rows;
        await context.sync();

        const made = sheet.tables.add('D12:E15', true);
        made.name = 'UndoSpikeTable';
        const covered = made.getRange();
        covered.load('address');
        await context.sync();
        assert(made.name === 'UndoSpikeTable', `table name was ${made.name}`);

        /*
         * Exactly what __restore does for kind:'created-table': getItemOrNullObject,
         * so a table the user deleted themselves reads as "already gone" rather
         * than throwing. This must stay in step with `undoCreatedTable` — a check
         * that verifies a different mechanism than the one that ships is worse
         * than no check.
         */
        const found = sheet.tables.getItemOrNullObject('UndoSpikeTable');
        found.load('isNullObject');
        await context.sync();
        assert(!found.isNullObject, 'no table named UndoSpikeTable was found, so the undo would report success and change nothing');
        found.convertToRange();
        await context.sync();

        const after = sheet.getRange('D12:E15');
        after.load('values');
        const remaining = sheet.tables;
        remaining.load('items/name');
        await context.sync();

        assert(
          !remaining.items.some((t) => t.name === 'UndoSpikeTable'),
          'the table survived convertToRange, so the undo would report success and change nothing'
        );
        assert(j(after.values) === j(rows), `the cells were not preserved: ${j(after.values)}`);
        return `table removed from ${covered.address}; its ${rows.length} rows are intact`;
      }),
  },
  {
    name: 'Create a chart',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        const chart = sheet.charts.add(Excel.ChartType.line, sheet.getRange('A1:B3'), Excel.ChartSeriesBy.columns);
        chart.title.text = 'Spike chart';
        chart.setPosition('G2', 'M15');
        chart.load('name');
        await context.sync();
        assert(!!chart.name, 'chart has no name');
        return `chart "${chart.name}" placed at G2:M15`;
      }),
  },
  {
    name: 'Create a PivotTable with row + data hierarchies',
    minSet: '1.8',
    run: () =>
      scratch(async (context, sheet) => {
        const pivot = sheet.pivotTables.add('SpikePivot', sheet.getRange('A1:C3'), sheet.getRange('O2'));
        pivot.rowHierarchies.add(pivot.hierarchies.getItem('region'));
        pivot.dataHierarchies.add(pivot.hierarchies.getItem('units'));
        pivot.load('name');
        await context.sync();
        const out = pivot.layout.getDataBodyRange();
        out.load('values');
        await context.sync();
        return `pivot "${pivot.name}" created; data body = ${j(out.values)}`;
      }),
  },
  {
    name: 'Undo a created chart (delete by name)',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        const chart = sheet.charts.add(Excel.ChartType.line, sheet.getRange('A1:B3'), Excel.ChartSeriesBy.columns);
        chart.name = 'UndoSpikeChart';
        // Placed away from the data below, so a failed delete is visible on screen.
        chart.setPosition('G17', 'M30');
        await context.sync();

        /*
         * Exactly what __restore does for kind:'created-chart'. Deleting is right
         * here where it was wrong for a table: a chart floats over the cells and
         * owns none of them, so there is no data to preserve.
         */
        const found = sheet.charts.getItemOrNullObject('UndoSpikeChart');
        found.load('isNullObject');
        await context.sync();
        assert(!found.isNullObject, 'the chart could not be found by name, so the undo would change nothing');
        found.delete();
        await context.sync();

        const after = sheet.charts.getItemOrNullObject('UndoSpikeChart');
        after.load('isNullObject');
        await context.sync();
        assert(after.isNullObject, 'the chart survived delete(), so undo would report success and leave it there');
        return 'chart created, found by name, and deleted';
      }),
  },
  {
    name: 'Undo a created pivot (delete by name)',
    minSet: '1.8',
    run: () =>
      scratch(async (context, sheet) => {
        const pivot = sheet.pivotTables.add('UndoSpikePivot', sheet.getRange('A1:C3'), sheet.getRange('O17'));
        pivot.rowHierarchies.add(pivot.hierarchies.getItem('region'));
        pivot.dataHierarchies.add(pivot.hierarchies.getItem('units'));
        await context.sync();

        // Exactly what __restore does for kind:'created-pivot'.
        const found = sheet.pivotTables.getItemOrNullObject('UndoSpikePivot');
        found.load('isNullObject');
        await context.sync();
        assert(!found.isNullObject, 'the pivot could not be found by name, so the undo would change nothing');
        found.delete();
        await context.sync();

        const after = sheet.pivotTables.getItemOrNullObject('UndoSpikePivot');
        after.load('isNullObject');
        await context.sync();
        assert(after.isNullObject, 'the pivot survived delete(), so undo would report success and leave it there');

        // The source is untouched — which is the whole claim the undo makes.
        const source = sheet.getRange('A1:C3');
        source.load('values');
        await context.sync();
        assert(source.values.length === 3, `the source data was disturbed: ${j(source.values)}`);
        return 'pivot created, found by name, and deleted; source intact';
      }),
  },
  {
    name: 'Remove duplicates',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        sheet.getRange('A18:A21').values = [['a'], ['b'], ['a'], ['b']];
        await context.sync();
        sheet.getRange('A18:A21').removeDuplicates([0], false);
        await context.sync();
        const r = sheet.getRange('A18:A21');
        r.load('values');
        await context.sync();
        const kept = r.values.flat().filter((v) => v !== '' && v != null);
        assert(kept.length === 2, `expected 2 unique rows, got ${kept.length} (${j(kept)})`);
        return `4 rows -> ${kept.length} unique (${kept.join(', ')})`;
      }),
  },
  {
    name: 'Remove-duplicates shift scope (undo coverage)',
    minSet: '1.1',
    info: true,
    run: () =>
      scratch(async (context, sheet) => {
        sheet.getRange('A18:A21').values = [['a'], ['b'], ['a'], ['b']];
        sheet.getRange('A23').values = [['marker-below']];
        await context.sync();

        sheet.getRange('A18:A21').removeDuplicates([0], false);
        await context.sync();

        const below = sheet.getRange('A23');
        below.load('values');
        await context.sync();
        const kept = below.values[0]?.[0];

        /*
         * Whether the journal can undo this depends entirely on the answer.
         * If the shift stays inside the range, a range snapshot restores it.
         * If it escapes, restoring the range leaves the workbook shifted.
         */
        return kept === 'marker-below'
          ? 'CONFINED — a row below the range was untouched, so a range snapshot can undo it'
          : `ESCAPED — the marker below the range now reads ${JSON.stringify(kept)}, so a range snapshot is NOT enough`;
      }),
  },
  {
    name: 'Worksheet add / rename / delete',
    minSet: '1.1',
    run: () =>
      Excel.run(async (context) => {
        const tmp = context.workbook.worksheets.add('_AI_SPIKE_TMP');
        tmp.name = '_AI_SPIKE_TMP2';
        tmp.load('name');
        await context.sync();
        assert(tmp.name === '_AI_SPIKE_TMP2', `rename failed: ${tmp.name}`);
        tmp.delete();
        await context.sync();
        return 'added, renamed and deleted a worksheet';
      }),
  },
  {
    name: 'Bulk write 2000 x 6 (throughput sanity)',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        const rows = 2000;
        const cols = 6;
        const data = Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, k) => i * cols + k));
        const r = sheet.getRange(`A30:${colName(cols - 1)}${29 + rows}`);
        const t0 = Date.now();
        r.values = data;
        await context.sync();
        const ms = Date.now() - t0;
        return `${rows * cols} cells written in ${ms}ms (${Math.round(((rows * cols) / ms) * 1000)} cells/s)`;
      }),
  },
  {
    name: 'Snapshot read for the undo journal',
    minSet: '1.1',
    run: () =>
      scratch(async (context, sheet) => {
        const r = sheet.getRange('A1:C3');
        r.load(['values', 'formulas', 'numberFormat']);
        await context.sync();
        const snapshot = { values: r.values, formulas: r.formulas, numberFormat: r.numberFormat };
        const bytes = new TextEncoder().encode(j(snapshot)).length;
        return `values+formulas+numberFormat for 9 cells = ${bytes} bytes (scales to ~${Math.round(bytes / 9)} B/cell)`;
      }),
  },
  {
    name: 'Same-origin back-channel (taskpane -> local server)',
    info: true,
    run: async () => {
      try {
        const res = await fetch('api/ping', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ hello: 'from the taskpane' }),
          cache: 'no-store',
        });
        if (!res.ok) return `HTTP ${res.status} from the local server.`;
        const body = (await res.json()) as { echoed?: Record<string, unknown> };
        return `REACHABLE — same-origin round-trip OK, server echoed ${body.echoed ? Object.keys(body.echoed).length : 0} key(s).`;
      } catch (err) {
        return `FAILED (${(err as Error).message}) — the taskpane cannot reach the server that served it.`;
      }
    },
  },
  {
    name: 'Cross-origin LLM call direct from the taskpane (CORS)',
    info: true,
    run: async () => {
      try {
        const res = await fetch('https://api.deepseek.com/v1/models', {
          headers: { Authorization: 'Bearer probe-invalid-key' },
          cache: 'no-store',
        });
        return `ALLOWED (HTTP ${res.status}) — the provider sends CORS headers. We still proxy through the local server to keep the API key out of the WebView.`;
      } catch (err) {
        return `BLOCKED (${(err as Error).message}) — no CORS headers, so LLM calls MUST be proxied through this origin.`;
      }
    },
  },
];

/* ------------------------------------------------------------------- panel */

const PANEL = `
  <div class="diag">
    <div class="row">
      <button id="diag-run" class="primary">Run self-test</button>
      <button id="diag-copy" disabled>Copy report</button>
      <button id="diag-clean" disabled>Delete scratch sheet</button>
    </div>
    <div id="diag-summary">Ready.</div>
    <ol id="diag-results"></ol>
  </div>
`;

export interface DiagnosticsOptions {
  available: boolean;
}

export function initDiagnostics(root: HTMLElement, { available }: DiagnosticsOptions): void {
  root.innerHTML = PANEL;

  const el = {
    run: root.querySelector<HTMLButtonElement>('#diag-run') as HTMLButtonElement,
    copy: root.querySelector<HTMLButtonElement>('#diag-copy') as HTMLButtonElement,
    clean: root.querySelector<HTMLButtonElement>('#diag-clean') as HTMLButtonElement,
    summary: root.querySelector<HTMLElement>('#diag-summary') as HTMLElement,
    results: root.querySelector<HTMLOListElement>('#diag-results') as HTMLOListElement,
  };

  let report: Record<string, unknown> | null = null;

  function renderRow(check: { name: string }, status: string, detail?: string, isError?: boolean): HTMLLIElement {
    const li = document.createElement('li');
    const head = document.createElement('div');
    head.className = 'head';
    const pill = document.createElement('span');
    pill.className = `pill ${status}`;
    pill.textContent = status;
    const name = document.createElement('span');
    name.className = 'name';
    name.textContent = check.name;
    head.append(pill, name);
    li.append(head);
    if (detail) {
      const d = document.createElement('div');
      d.className = isError ? 'detail err' : 'detail';
      d.textContent = detail;
      li.append(d);
    }
    el.results.append(li);
    return li;
  }

  async function ensureScratchSheet(): Promise<void> {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();
      const stale = sheets.items.find((s) => s.name === SCRATCH);
      if (stale) stale.delete();
      sheets.add(SCRATCH).activate();
      await context.sync();
    });
  }

  async function runAll(): Promise<void> {
    el.run.disabled = true;
    el.copy.disabled = true;
    el.clean.disabled = true;
    el.results.replaceChildren();
    el.summary.textContent = 'Preparing scratch sheet…';
    const checks: unknown[] = [];
    report = { ranAt: new Date().toISOString(), checks };

    try {
      await ensureScratchSheet();
    } catch (err) {
      el.summary.textContent = 'Could not create the scratch sheet.';
      renderRow({ name: 'Create scratch sheet' }, 'FAIL', fmtError(err), true);
      el.run.disabled = false;
      return;
    }

    const tally = { PASS: 0, FAIL: 0, SKIP: 0, INFO: 0 };

    for (const check of CHECKS) {
      if (check.minSet && !Office.context.requirements.isSetSupported('ExcelApi', check.minSet)) {
        tally.SKIP++;
        checks.push({ name: check.name, status: 'SKIP', detail: `requires ExcelApi ${check.minSet}` });
        renderRow(check, 'SKIP', `requires ExcelApi ${check.minSet} — not available in this host`);
        continue;
      }

      const placeholder = renderRow(check, 'RUN', 'running…');
      try {
        const detail = await check.run();
        const status = check.info ? 'INFO' : 'PASS';
        tally[status]++;
        checks.push({ name: check.name, status, detail });
        placeholder.replaceWith(renderRow(check, status, detail));
      } catch (err) {
        tally.FAIL++;
        const detail = fmtError(err);
        checks.push({ name: check.name, status: 'FAIL', detail });
        placeholder.replaceWith(renderRow(check, 'FAIL', detail, true));
      }
      el.summary.textContent = `${tally.PASS} passed · ${tally.FAIL} failed · ${tally.SKIP} skipped`;
    }

    report.environment = {
      host: Office.context.host,
      platform: Office.context.platform,
      diagnostics: Office.context.diagnostics || null,
      excelApiSets: supportedSets(),
    };

    el.summary.textContent =
      `${tally.PASS} passed · ${tally.FAIL} failed · ${tally.SKIP} skipped · ${tally.INFO} info — ` +
      `ExcelApi up to ${supportedSets().filter((s) => s !== 'Online 1.1').pop() || 'none'}`;
    el.run.disabled = false;
    el.copy.disabled = false;
    el.clean.disabled = false;
  }

  el.run.addEventListener('click', () => {
    void runAll();
  });

  el.copy.addEventListener('click', async () => {
    if (!report) return;
    const text = JSON.stringify(report, null, 2);
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const ta = document.createElement('textarea');
      ta.value = text;
      document.body.append(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
    }
    el.copy.textContent = 'Copied ✓';
    setTimeout(() => {
      el.copy.textContent = 'Copy report';
    }, 1500);
  });

  el.clean.addEventListener('click', async () => {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();
      const target = sheets.items.find((s) => s.name === SCRATCH);
      if (target) target.delete();
      await context.sync();
    });
    el.clean.disabled = true;
    el.summary.textContent = 'Scratch sheet deleted.';
  });

  if (!available) {
    el.run.disabled = true;
    el.summary.textContent = 'Not running inside Excel.';
  } else {
    el.summary.textContent = 'Ready. Click "Run self-test".';
  }
}
