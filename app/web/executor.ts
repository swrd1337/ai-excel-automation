/*
 * Executes tool calls inside Excel via Office.js.
 *
 * This is the half of the system that must live in the taskpane: only code
 * running inside Excel can touch the workbook. The server decides *what* to do;
 * this decides *how*, and reports back.
 *
 * EXECUTOR_NAMES is sent to the server on every turn so it can only ever ask for
 * tools this build can actually run — a stale taskpane can't be sent a tool it
 * doesn't know.
 *
 * Types come from @types/office-js, which is the point: it is what catches
 * `conditionalFormats.count` (there is no such property) or `autoFilter.range`
 * (it is `getRange()`) at compile time rather than at the user's expense.
 */

export type ToolArgs = Record<string, unknown>;

/**
 * One reversible step: recorded BEFORE the change is made, replayed by
 * `__restore` if the user asks for it back.
 *
 * It is a union because not every change is a range of cells. Creating a table,
 * chart or pivot makes an object, and no amount of cell state brings that back —
 * a range snapshot would restore values that were never touched while leaving the
 * object behind, so undo would silently do nothing. `kind` is the one contract
 * between the code that records a change and the code that reverses it, which is
 * why the server stores these records without ever reading them.
 *
 * Separate `created-*` kinds rather than one kind with an `objectType`: the journal
 * is on disk, so entries outlive a build, and collapsing them would strand the
 * `created-table` records already written.
 */
export type UndoAction =
  | { kind: 'range'; sheet: string | null; address: string; state: Record<string, unknown> }
  | { kind: 'created-table'; sheet: string | null; name: string }
  | { kind: 'created-chart'; sheet: string | null; name: string }
  | { kind: 'created-pivot'; sheet: string | null; name: string };

/** The created-object kinds, which differ only in how the object is taken away. */
type CreatedAction = Extract<UndoAction, { kind: 'created-table' | 'created-chart' | 'created-pivot' }>;

export interface SnapshotResult {
  /** What the journal keeps, and what `__restore` is handed back later. */
  action: UndoAction;
  /** How the change is named to a person: undo messages, the approval card. */
  label: string;
  /** The cells as they stand now, for the approval card's "now" column. */
  preview: unknown;
  /**
   * Whether the server must ask for a post-write fingerprint before this change
   * can be undone later. True exactly when undoing rewrites cell state, which is
   * when writing an old snapshot over newer work is possible. The server reads
   * this flag rather than inspecting `action.kind`, so it still never needs to
   * understand the union.
   */
  seal: boolean;
}

type Executor = (args: ToolArgs) => Promise<unknown>;

/* ------------------------------------------------------------- address maths */

function colToNum(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function numToCol(n: number): string {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** Dimensions of an explicit "A1:C3" style range, or null for a single cell. */
export function rangeDimensions(address: string): { rowCount: number; columnCount: number } | null {
  const m = /^\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$/.exec(String(address).trim());
  if (!m) return null;
  return {
    rowCount: Math.abs(Number(m[4]) - Number(m[2])) + 1,
    columnCount: Math.abs(colToNum(m[3] as string) - colToNum(m[1] as string)) + 1,
  };
}

/**
 * Excel's own error for this is "The number of rows or columns in the input array
 * doesn't match the size or dimensions of the range", which tells the model
 * nothing actionable. Say exactly what is wrong so it can fix its own call.
 */
export function checkShape(address: string, rowCount: number, columnCount: number): string | null {
  const dims = rangeDimensions(address);
  if (!dims) return null; // single cell or unparseable: expansion handles it
  if (dims.rowCount !== rowCount || dims.columnCount !== columnCount) {
    return (
      `${address} is ${dims.rowCount}x${dims.columnCount} but ${rowCount}x${columnCount} values were given. ` +
      `Use a single-cell anchor like "A1" to let it expand, or supply exactly ${dims.rowCount}x${dims.columnCount} values.`
    );
  }
  return null;
}

/**
 * Writing a 2D array into a single-cell anchor should fill outward, and we want
 * to report the address we actually wrote. A multi-cell address is left alone so
 * a shape mismatch surfaces as an error instead of silently truncating.
 */
export function expandAddress(address: string, rowCount: number, columnCount: number): string {
  const clean = String(address).trim();
  if (clean.includes(':')) return clean;
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(clean);
  if (!m) return clean;
  const startCol = colToNum(m[1] as string);
  const startRow = Number(m[2]);
  return `${numToCol(startCol)}${startRow}:${numToCol(startCol + columnCount - 1)}${startRow + rowCount - 1}`;
}

function shapeOf(values: unknown): { rowCount: number; columnCount: number } {
  if (!Array.isArray(values)) return { rowCount: 0, columnCount: 0 };
  const rows = values.length;
  const cols = values.reduce<number>((max, row) => Math.max(max, Array.isArray(row) ? row.length : 1), 0);
  return { rowCount: rows, columnCount: cols };
}

function sheetOf(context: Excel.RequestContext, name: unknown): Excel.Worksheet {
  return name
    ? context.workbook.worksheets.getItem(String(name))
    : context.workbook.worksheets.getActiveWorksheet();
}

function formatExcelError(err: unknown): string {
  if (!err) return 'unknown error';
  const e = err as { code?: string; message?: string; debugInfo?: { message?: string } };
  const parts: string[] = [];
  if (e.code) parts.push(e.code);
  if (e.message) parts.push(e.message);
  if (e.debugInfo) parts.push(`(debug: ${e.debugInfo.message || JSON.stringify(e.debugInfo)})`);
  return parts.join(' — ') || String(err);
}

/* ------------------------------------------------------------------ context */

export interface WorkbookContext {
  activeSheet: string | null;
  sheets: { name: string; usedRange: string | null; rowCount: number; columnCount: number }[];
  selection: { address: string; values: unknown; truncated: boolean } | null;
}

/** The workbook packet fetched by the server so the model knows what it's facing. */
export async function getContext(): Promise<WorkbookContext> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    const active = sheets.getActiveWorksheet();
    active.load('name');
    await context.sync();

    const infos = sheets.items.map((s) => {
      const used = s.getUsedRangeOrNullObject(true);
      used.load(['address', 'rowCount', 'columnCount', 'isNullObject']);
      return { name: s.name, used };
    });
    await context.sync();

    const packet: WorkbookContext = {
      activeSheet: active.name,
      sheets: infos.map(({ name, used }) =>
        used.isNullObject
          ? { name, usedRange: null, rowCount: 0, columnCount: 0 }
          : { name, usedRange: used.address, rowCount: used.rowCount, columnCount: used.columnCount }
      ),
      selection: null,
    };

    try {
      const sel = context.workbook.getSelectedRange();
      sel.load(['address', 'values', 'rowCount', 'columnCount']);
      await context.sync();
      const rows = Math.min(sel.rowCount, 5);
      const cols = Math.min(sel.columnCount, 8);
      const values = sel.values as unknown[][];
      packet.selection = {
        address: sel.address,
        values: values.slice(0, rows).map((row) => row.slice(0, cols)),
        truncated: sel.rowCount > rows || sel.columnCount > cols,
      };
    } catch {
      // A selection isn't always available (e.g. focus is in a chart). Not fatal.
    }

    return packet;
  });
}

/* --------------------------------------------------------------- snapshots */

/**
 * The range a mutating tool will actually touch, AFTER any expansion.
 * write_range("A1", 3x2) writes A1:B3, not A1 — snapshotting the wrong range
 * would make undo silently incomplete.
 *
 * This stays in the taskpane so "what does this tool affect?" has exactly one
 * answer, shared with the code that does the writing.
 */
export function targetAddressFor(tool: string, args: ToolArgs): string | null {
  const address = args.address as string | undefined;
  if (!address) return null;

  if (tool === 'write_range') {
    const { rowCount, columnCount } = shapeOf(args.values);
    return rowCount && columnCount ? expandAddress(address, rowCount, columnCount) : String(address).trim();
  }
  if (tool === 'set_formula') {
    const { rowCount, columnCount } = shapeOf(args.formulas);
    return rowCount && columnCount ? expandAddress(address, rowCount, columnCount) : String(address).trim();
  }
  // format_range, and any future tool that just takes an address.
  return String(address).trim();
}

const FORMAT_PROPS = [
  'format/font/bold',
  'format/font/italic',
  'format/font/color',
  'format/fill/color',
  'format/horizontalAlignment',
];

/**
 * Everything needed to put a range back exactly as it was.
 *
 * This answers with an UndoAction rather than raw cell state so the journal has
 * one shape whatever kind of change it is holding.
 */
async function captureRange({ tool, args }: ToolArgs): Promise<SnapshotResult> {
  const address = targetAddressFor(String(tool), args as ToolArgs);
  if (!address) throw new Error(`cannot tell which range "${String(tool)}" affects, so it is not safe to run`);

  const sheet = (args as ToolArgs).sheet as string | undefined;

  return Excel.run(async (context) => {
    const range = sheetOf(context, sheet).getRange(address);
    range.load(['address', 'values', 'formulas', 'numberFormat', ...FORMAT_PROPS]);
    await context.sync();

    return {
      action: {
        kind: 'range',
        sheet: sheet || null,
        address: range.address,
        state: {
          formulas: range.formulas,
          numberFormat: range.numberFormat,
          font: {
            bold: range.format.font.bold,
            italic: range.format.font.italic,
            color: range.format.font.color,
          },
          fillColor: range.format.fill.color,
          horizontalAlignment: range.format.horizontalAlignment,
        },
      },
      label: range.address,
      // For a preview we also want what the cells hold right now, human-readable.
      preview: range.values,
      seal: true,
    };
  });
}

/* ------------------------------------------------------------- fingerprints */

/**
 * A short, stable digest of some cell state.
 *
 * Deliberately not a cryptographic hash: `node:crypto` does not exist inside
 * Excel's WebView, and this is a staleness check, not a security boundary. Two
 * rounds of FNV-1a give 64 bits, which is far more than enough to notice that a
 * person has edited a cell. A collision would mean failing to notice an edit —
 * with a handful of comparisons against 2^64 values, that is not a real risk.
 */
export function fingerprintState(value: unknown): string {
  const text = JSON.stringify(value ?? null);
  let a = 0x811c9dc5;
  let b = 0x9e3779b1;
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    a = Math.imul(a ^ c, 0x01000193) >>> 0;
    b = Math.imul(b ^ c, 0x85ebca6b) >>> 0;
  }
  return a.toString(16).padStart(8, '0') + b.toString(16).padStart(8, '0');
}

/** Every part `__restore` would put back, digested as one value. */
function fingerprintRange(range: Excel.Range): string {
  return fingerprintState([
    range.formulas,
    range.numberFormat,
    { bold: range.format.font.bold, italic: range.format.font.italic, color: range.format.font.color },
    range.format.fill.color,
    range.format.horizontalAlignment,
  ]);
}

/**
 * Whether Excel will accept this as a table name.
 *
 * Checked here rather than left to Office.js, whose rejection ("The argument is
 * invalid or missing or has an incorrect format") names neither the rule nor the
 * offending character.
 */
export function checkTableName(name: string): string | null {
  if (!name) return 'a table name is required';
  if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name)) {
    return (
      `"${name}" is not a usable table name. Start with a letter or underscore, then use only ` +
      `letters, digits, underscores and full stops — no spaces. Try "SalesTable".`
    );
  }
  if (name === 'R' || name === 'C' || /^[A-Za-z]{1,3}\d{1,7}$/.test(name)) {
    return `"${name}" is a cell reference, which Excel does not allow as a table name. Try "Table_${name}".`;
  }
  return null;
}

/**
 * The undo plan for creating a table.
 *
 * The name is settled and checked for availability BEFORE anything is made,
 * because the undo is "find the table with this name and convert it back to a
 * range". If Excel invented the name at write time we could not write that plan
 * down first — and a change we cannot describe before making it is a change we
 * refuse to make. It is also why the tool asks the model for a name at all.
 */
async function planCreateTable(args: ToolArgs): Promise<SnapshotResult> {
  const address = String(args.address ?? '').trim();
  const name = String(args.name ?? '').trim();
  const sheet = args.sheet as string | undefined;
  const hasHeaders = args.hasHeaders !== false;

  const problem = checkTableName(name);
  if (problem) throw new Error(problem);

  /*
   * A single-cell anchor would be expanded by Excel to whatever region it is
   * part of, so the range the reviewer approves could be far smaller than the
   * range the table covers. Require the full range instead of guessing.
   */
  if (!address.includes(':')) {
    throw new Error(
      `give the full range the table should cover, e.g. "A1:D20". A single-cell anchor like ` +
      `"${address}" would be expanded to whatever region it happens to be part of.`
    );
  }

  return Excel.run(async (context) => {
    const range = sheetOf(context, sheet).getRange(address);
    range.load(['address', 'values']);

    // Table names are unique across the WHOLE workbook, not just this sheet.
    const allTables = context.workbook.tables;
    allTables.load('items/name');
    await context.sync();

    const taken = allTables.items.find((t) => t.name.toLowerCase() === name.toLowerCase());
    if (taken) throw new Error(`a table named "${taken.name}" already exists in this workbook. Pick another name.`);

    return {
      action: { kind: 'created-table', sheet: sheet || null, name },
      label: `${name} — a table over ${range.address}${hasHeaders ? '' : ', with no header row'}`,
      preview: range.values,
      // Nothing to fingerprint: removing a table you created cannot overwrite
      // cell data, so a stale-check would protect nothing.
      seal: false,
    };
  });
}

/**
 * The name rule that is genuinely shared.
 *
 * Deliberately NOT `checkTableName`'s stricter rule: Excel's own chart names look
 * like "Chart 1", with a space, so the rules are not the same and applying the
 * table's would refuse names Excel accepts.
 */
export function checkObjectName(name: string, what: string): string | null {
  if (!name) return `a ${what} name is required`;
  if (name.length > 250) return `that ${what} name is ${name.length} characters long — Excel's limit is 255`;
  return null;
}

/**
 * The undo plan for creating a chart.
 *
 * Same shape as the table's, and for the same reason: the undo is "find the chart
 * with this name and delete it", so the name has to be settled — and free — before
 * the chart exists. A change we cannot describe in advance is one we refuse.
 *
 * Deleting is right here where it was wrong for a table: a chart floats over the
 * cells and owns none of them, so removing it destroys no data.
 */
async function planCreateChart(args: ToolArgs): Promise<SnapshotResult> {
  const address = String(args.address ?? '').trim();
  const name = String(args.name ?? '').trim();
  const sheet = args.sheet as string | undefined;
  const type = String(args.type ?? 'ColumnClustered');
  const seriesBy = String(args.seriesBy ?? 'Columns');

  const problem = checkObjectName(name, 'chart');
  if (problem) throw new Error(problem);

  return Excel.run(async (context) => {
    const worksheet = sheetOf(context, sheet);
    const source = worksheet.getRange(address);
    source.load(['address', 'values']);

    const existing = worksheet.charts.getItemOrNullObject(name);
    existing.load('isNullObject');
    await context.sync();

    if (!existing.isNullObject) throw new Error(`a chart named "${name}" already exists on this sheet. Pick another name.`);

    return {
      action: { kind: 'created-chart', sheet: sheet || null, name },
      label: `${name} — a ${type} chart of ${source.address}`,
      preview: source.values,
      // Removing a chart cannot overwrite cell data, so a stale-check would
      // protect nothing — same reasoning as the table.
      seal: false,
    };
  });
}

/**
 * The undo plan for creating a pivot table.
 *
 * A pivot writes computed cells into its destination, which is why the preview
 * shows what is there now: the reviewer is approving the loss of that space, and
 * undo removes the pivot and the cells it occupied.
 */
async function planCreatePivot(args: ToolArgs): Promise<SnapshotResult> {
  const name = String(args.name ?? '').trim();
  const source = String(args.source ?? '').trim();
  const destination = String(args.destination ?? '').trim();
  const sheet = args.sheet as string | undefined;

  const problem = checkObjectName(name, 'pivot table');
  if (problem) throw new Error(problem);

  return Excel.run(async (context) => {
    const worksheet = sheetOf(context, sheet);
    const target = worksheet.getRange(destination);
    target.load(['address', 'values']);

    const existing = worksheet.pivotTables.getItemOrNullObject(name);
    existing.load('isNullObject');
    await context.sync();

    if (!existing.isNullObject) {
      throw new Error(`a pivot table named "${name}" already exists on this sheet. Pick another name.`);
    }

    return {
      action: { kind: 'created-pivot', sheet: sheet || null, name },
      label: `${name} — a pivot at ${target.address} from ${source}`,
      preview: target.values,
      seal: false,
    };
  });
}

/** The plan for one change, choosing how to describe it by what the tool does. */
async function captureUndo(call: ToolArgs): Promise<SnapshotResult> {
  switch (String(call.tool)) {
    case 'create_table':
      return planCreateTable(call.args as ToolArgs);
    case 'create_chart':
      return planCreateChart(call.args as ToolArgs);
    case 'create_pivot_table':
      return planCreatePivot(call.args as ToolArgs);
    default:
      return captureRange(call);
  }
}

/**
 * Restore from `formulas`, never from `values`: assigning `values` to a cell that
 * held a formula writes the computed number and destroys the formula. `formulas`
 * returns the formula string for formula cells and the plain value elsewhere, so
 * it round-trips both.
 *
 * Before touching anything, the range is read and compared against the
 * fingerprint taken when the change was made. It has to be: a journal that
 * outlives the process can be far older than the cells it describes, and writing
 * an old snapshot over newer work is the exact data loss the journal exists to
 * prevent. When the cells have moved on we refuse and say so, rather than
 * destroy the newer work to make the undo look successful.
 */
async function undoRange(action: Extract<UndoAction, { kind: 'range' }>, expect: string | null): Promise<unknown> {
  const s = action.state as {
    formulas?: unknown;
    numberFormat?: unknown;
    font?: { bold?: unknown; italic?: unknown; color?: unknown };
    fillColor?: unknown;
    horizontalAlignment?: unknown;
  };

  return Excel.run(async (context) => {
    const range = sheetOf(context, action.sheet).getRange(action.address);
    range.load(['address', 'formulas', 'numberFormat', ...FORMAT_PROPS]);
    await context.sync();

    if (!expect) {
      throw new Error(
        'this change was recorded without a fingerprint, so there is no way to tell whether these cells have ' +
          'been edited since. Nothing was restored.'
      );
    }
    if (fingerprintRange(range) !== expect) {
      throw new Error(
        `${range.address} has changed since this task ran, so undoing it would overwrite newer work. ` +
          'Nothing was restored — undo the newer change first, or edit by hand.'
      );
    }

    if (s.formulas) range.formulas = s.formulas as string[][];
    if (s.numberFormat) range.numberFormat = s.numberFormat as string[][];

    if (s.font) {
      // Mixed formatting reads back as null; only assign what we actually captured.
      if (s.font.bold !== null && s.font.bold !== undefined) range.format.font.bold = s.font.bold as boolean;
      if (s.font.italic !== null && s.font.italic !== undefined) range.format.font.italic = s.font.italic as boolean;
      if (s.font.color) range.format.font.color = s.font.color as string;
    }
    if (s.fillColor) range.format.fill.color = s.fillColor as string;
    if (s.horizontalAlignment) range.format.horizontalAlignment = s.horizontalAlignment as Excel.HorizontalAlignment;

    range.load('address');
    await context.sync();
    return { restored: range.address };
  });
}

/**
 * Undoing a created object means taking the object away and leaving the cells.
 *
 * `getItemOrNullObject` rather than `getItem`: the user deleting the object
 * themselves is a legitimate way to end up with nothing left to undo, and that
 * should read as "already gone", not as an error. (An earlier version of the
 * table path loaded every table and searched by hand, because a truncated grep
 * made it look as though this method did not exist.)
 */
async function undoCreatedTable(action: Extract<UndoAction, { kind: 'created-table' }>): Promise<unknown> {
  return Excel.run(async (context) => {
    const table = sheetOf(context, action.sheet).tables.getItemOrNullObject(action.name);
    table.load('isNullObject');
    await context.sync();

    if (table.isNullObject) return { alreadyGone: action.name };

    /*
     * `convertToRange()`, not `delete()`: the cells held that data before the
     * table existed, so removing the table must not take them with it.
     */
    table.convertToRange();
    await context.sync();
    return { removedTable: action.name, kept: 'the cells it covered' };
  });
}

/** A chart floats over the cells and owns none of them, so removing it loses no data. */
async function undoCreatedChart(action: Extract<UndoAction, { kind: 'created-chart' }>): Promise<unknown> {
  return Excel.run(async (context) => {
    const chart = sheetOf(context, action.sheet).charts.getItemOrNullObject(action.name);
    chart.load('isNullObject');
    await context.sync();

    if (chart.isNullObject) return { alreadyGone: action.name };

    chart.delete();
    await context.sync();
    return { removedChart: action.name };
  });
}

/**
 * A pivot writes computed cells into its destination, so deleting it takes those
 * with it — which is the change that reverses creating it. Nothing is lost: the
 * source data is untouched, and the destination is whatever the pivot computed.
 */
async function undoCreatedPivot(action: Extract<UndoAction, { kind: 'created-pivot' }>): Promise<unknown> {
  return Excel.run(async (context) => {
    const pivot = sheetOf(context, action.sheet).pivotTables.getItemOrNullObject(action.name);
    pivot.load('isNullObject');
    await context.sync();

    if (pivot.isNullObject) return { alreadyGone: action.name };

    pivot.delete();
    await context.sync();
    return { removedPivot: action.name, kept: 'the source data it was built from' };
  });
}

/**
 * What a change left behind, so a later undo can tell whether anything has
 * touched those cells since. Asked for AFTER the write, and only for the kinds
 * that said `seal: true`.
 */
async function fingerprintUndo(record: ToolArgs): Promise<unknown> {
  const action = record as unknown as UndoAction | undefined;
  if (!action || action.kind !== 'range') return { expect: null };

  return Excel.run(async (context) => {
    const range = sheetOf(context, action.sheet).getRange(action.address);
    range.load(['address', 'formulas', 'numberFormat', ...FORMAT_PROPS]);
    await context.sync();
    return { expect: fingerprintRange(range) };
  });
}

/** Replay one recorded change. The server hands these back verbatim. */
async function restoreUndo(record: ToolArgs): Promise<unknown> {
  const { action, expect = null } = record as unknown as { action?: UndoAction; expect?: string | null };

  // Untrusted wire input: a record from another build could be anything at all.
  if (!action || typeof action !== 'object' || !('kind' in action)) {
    throw new Error('this undo record carries no action, so there is nothing to replay — reload the add-in taskpane');
  }

  switch (action.kind) {
    case 'range':
      return undoRange(action, expect ?? null);
    case 'created-table':
      return undoCreatedTable(action);
    case 'created-chart':
      return undoCreatedChart(action);
    case 'created-pivot':
      return undoCreatedPivot(action);
    default:
      // A taskpane from another build could send a kind this one has never
      // heard of, and guessing would be worse than saying so.
      throw new Error(`cannot undo a "${(action as { kind?: string }).kind}" change — reload the add-in taskpane`);
  }
}

/* ---------------------------------------------------------------- executors */

export const EXECUTORS: Record<string, Executor> = {
  async get_workbook_overview() {
    const ctx = await getContext();
    return { activeSheet: ctx.activeSheet, sheets: ctx.sheets };
  },

  async read_range(args) {
    const address = String(args.address);
    const sheet = args.sheet as string | undefined;
    return Excel.run(async (context) => {
      const range = sheetOf(context, sheet).getRange(address);
      range.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);
      await context.sync();
      const out: Record<string, unknown> = {
        address: range.address,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        values: range.values,
      };
      // Formulas only for smaller ranges: they double the payload for no benefit on big reads.
      if (range.rowCount * range.columnCount <= 500) out.formulas = range.formulas;
      return out;
    });
  },

  async get_selection() {
    return Excel.run(async (context) => {
      const sel = context.workbook.getSelectedRange();
      sel.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);
      await context.sync();
      return {
        address: sel.address,
        rowCount: sel.rowCount,
        columnCount: sel.columnCount,
        values: sel.values,
        formulas: sel.formulas,
      };
    });
  },

  async write_range(args) {
    const address = String(args.address);
    const sheet = args.sheet as string | undefined;
    const values = args.values;
    const { rowCount, columnCount } = shapeOf(values);
    if (!rowCount || !columnCount) throw new Error('values must be a non-empty 2D array');
    const mismatch = checkShape(address, rowCount, columnCount);
    if (mismatch) throw new Error(mismatch);

    return Excel.run(async (context) => {
      const target = expandAddress(address, rowCount, columnCount);
      const range = sheetOf(context, sheet).getRange(target);
      range.values = values as (string | number | boolean)[][];
      range.load('address');
      await context.sync();
      return { wroteTo: range.address, rowCount, columnCount };
    });
  },

  async set_formula(args) {
    const address = String(args.address);
    const sheet = args.sheet as string | undefined;
    const formulas = args.formulas;
    const { rowCount, columnCount } = shapeOf(formulas);
    if (!rowCount || !columnCount) throw new Error('formulas must be a non-empty 2D array');
    const flat = (formulas as unknown[][]).flat();
    const notFormula = flat.find((f) => typeof f !== 'string' || !f.startsWith('='));
    if (notFormula !== undefined) throw new Error(`every formula must start with "=", got ${JSON.stringify(notFormula)}`);
    const mismatch = checkShape(address, rowCount, columnCount);
    if (mismatch) throw new Error(mismatch);

    return Excel.run(async (context) => {
      const target = expandAddress(address, rowCount, columnCount);
      const range = sheetOf(context, sheet).getRange(target);
      range.formulas = formulas as string[][];
      range.load(['address', 'values']);
      await context.sync();
      return { wroteTo: range.address, computedValues: range.values };
    });
  },

  async format_range(args) {
    const address = String(args.address);
    const sheet = args.sheet as string | undefined;
    const { bold, italic, fontColor, fillColor, numberFormat, horizontalAlignment } = args as {
      bold?: boolean;
      italic?: boolean;
      fontColor?: string;
      fillColor?: string;
      numberFormat?: string;
      horizontalAlignment?: Excel.HorizontalAlignment;
    };

    return Excel.run(async (context) => {
      const range = sheetOf(context, sheet).getRange(address);
      const applied: string[] = [];

      if (bold !== undefined) {
        range.format.font.bold = bold;
        applied.push(`bold=${bold}`);
      }
      if (italic !== undefined) {
        range.format.font.italic = italic;
        applied.push(`italic=${italic}`);
      }
      if (fontColor !== undefined) {
        range.format.font.color = fontColor;
        applied.push(`fontColor=${fontColor}`);
      }
      if (fillColor !== undefined) {
        range.format.fill.color = fillColor;
        applied.push(`fillColor=${fillColor}`);
      }
      if (numberFormat !== undefined) {
        // The published type is string[][], but Excel accepts one code for the
        // whole range, which is what the model is asked to supply.
        range.numberFormat = numberFormat as unknown as string[][];
        applied.push(`numberFormat=${numberFormat}`);
      }
      if (horizontalAlignment !== undefined) {
        range.format.horizontalAlignment = horizontalAlignment;
        applied.push(`horizontalAlignment=${horizontalAlignment}`);
      }

      if (!applied.length) throw new Error('no formatting properties were supplied');
      range.load('address');
      await context.sync();
      return { formatted: range.address, applied };
    });
  },

  async sort_range(args) {
    const address = String(args.address);
    const sheet = args.sheet as string | undefined;
    const keys = (args.keys || []) as { column: number; ascending?: boolean }[];
    if (!keys.length) throw new Error('at least one sort key is required');
    const hasHeaders = args.hasHeaders !== false;

    return Excel.run(async (context) => {
      const range = sheetOf(context, sheet).getRange(address);
      range.sort.apply(
        keys.map((k) => ({ key: k.column, ascending: k.ascending !== false })),
        false,
        hasHeaders
      );
      range.load('address');
      await context.sync();
      return { sorted: range.address, keys: keys.length, hasHeaders };
    });
  },

  async create_table(args) {
    const address = String(args.address ?? '').trim();
    const name = String(args.name ?? '').trim();
    const sheet = args.sheet as string | undefined;
    const hasHeaders = args.hasHeaders !== false;

    const problem = checkTableName(name);
    if (problem) throw new Error(problem);

    return Excel.run(async (context) => {
      const table = sheetOf(context, sheet).tables.add(address, hasHeaders);
      table.name = name;
      table.load('name');
      // A Table has no address of its own; ask its range. What Excel actually
      // covered is worth reporting, since it can differ from what was asked for.
      const covered = table.getRange();
      covered.load('address');
      await context.sync();
      return { table: table.name, address: covered.address, hasHeaders };
    });
  },

  async create_chart(args) {
    const address = String(args.address ?? '').trim();
    const name = String(args.name ?? '').trim();
    const sheet = args.sheet as string | undefined;
    const type = String(args.type ?? 'ColumnClustered') as Excel.ChartType;
    const seriesBy = String(args.seriesBy ?? 'Columns') as Excel.ChartSeriesBy;

    const problem = checkObjectName(name, 'chart');
    if (problem) throw new Error(problem);

    return Excel.run(async (context) => {
      const worksheet = sheetOf(context, sheet);
      const chart = worksheet.charts.add(type, worksheet.getRange(address), seriesBy);
      chart.name = name;
      // The position is optional and cosmetic: a chart with no position lands
      // wherever Excel puts it, which is still a usable chart.
      if (args.position) chart.setPosition(String(args.position));
      chart.load('name');
      await context.sync();
      return { chart: chart.name, type, source: address, seriesBy };
    });
  },

  async create_pivot_table(args) {
    const name = String(args.name ?? '').trim();
    const source = String(args.source ?? '').trim();
    const destination = String(args.destination ?? '').trim();
    const sheet = args.sheet as string | undefined;
    const rows = (args.rows || []) as string[];
    const values = (args.values || []) as string[];

    const problem = checkObjectName(name, 'pivot table');
    if (problem) throw new Error(problem);

    return Excel.run(async (context) => {
      const worksheet = sheetOf(context, sheet);
      const pivot = worksheet.pivotTables.add(name, worksheet.getRange(source), worksheet.getRange(destination));

      /*
       * Hierarchies are addressed by the *field name* in the source headers, not
       * by index, so a name the model invented surfaces as an error naming the
       * field rather than as an empty pivot.
       */
      for (const field of rows) pivot.rowHierarchies.add(pivot.hierarchies.getItem(field));
      for (const field of values) pivot.dataHierarchies.add(pivot.hierarchies.getItem(field));
      pivot.load('name');
      await context.sync();

      return { pivot: pivot.name, source, destination, rows: rows.length, values: values.length };
    });
  },

  async remove_duplicates(args) {
    const address = String(args.address ?? '').trim();
    const sheet = args.sheet as string | undefined;
    const columns = (args.columns || []) as number[];
    const hasHeaders = args.hasHeaders !== false;

    if (!columns.length) {
      throw new Error(
        "name at least one column to compare, as an index within the range: 0 is the range's first column."
      );
    }
    if (columns.some((c) => !Number.isInteger(c) || c < 0)) {
      throw new Error(`columns must be whole numbers counting from 0, got ${JSON.stringify(columns)}`);
    }

    return Excel.run(async (context) => {
      const range = sheetOf(context, sheet).getRange(address);
      range.load(['address', 'rowCount', 'columnCount']);
      await context.sync();

      /*
       * Checked here because Excel's own failure is misleading: an index past the
       * end of the range comes back as a generic invalid-argument error, which
       * tells the model nothing about which index was wrong or why.
       */
      const tooFar = columns.filter((c) => c >= range.columnCount);
      if (tooFar.length) {
        throw new Error(
          `column ${tooFar.join(', ')} is outside ${range.address}, which has ${range.columnCount} ` +
            `column(s) — indices count from the range's first column, not from column A.`
        );
      }

      /*
       * The undo is the ordinary range snapshot. Diagnostics measured whether that
       * is honest — "Remove-duplicates shift scope" — and found the row shift stays
       * INSIDE the range, so a snapshot of the range covers everything it moves.
       * If that ever reports otherwise, this tool stops being undoable and must be
       * withdrawn rather than left looking safe.
       */
      const result = range.removeDuplicates(columns, hasHeaders);
      result.load(['removed', 'uniqueRemaining']);
      await context.sync();

      // What Excel actually did, rather than a bare success: the model can then
      // report how many rows went instead of guessing.
      return {
        address: range.address,
        removed: result.removed,
        remaining: result.uniqueRemaining,
        comparedColumns: columns,
      };
    });
  },

  /*
   * Internal, neither offered to the model nor advertised as a tool — the server
   * dispatches these directly to implement preview and undo.
   */
  __snapshot: captureUndo,
  __fingerprint: fingerprintUndo,
  __restore: restoreUndo,
};

/** Every call this taskpane can service, including the server-internal helpers. */
export const EXECUTOR_NAMES: string[] = Object.keys(EXECUTORS);

/*
 * The server asks for the workbook context itself rather than having the UI
 * supply it, which is what lets the UI be anything that can POST.
 *
 * It is deliberately NOT an EXECUTORS entry: the taskpane shell answers it
 * directly so that a context read — which happens on every single turn — is not
 * logged as workbook activity or counted as a tool call. It lives here so the
 * name is still owned by the taskpane rather than written out a second time.
 */
export const CONTEXT_TOOL = '__workbook_context';

/* --------------------------------------------------------------- dispatcher */

export interface ExecutorCall {
  id: string;
  name: string;
  args: ToolArgs;
  error?: string;
}

export interface ExecutorResult {
  id: string;
  name: string;
  content?: unknown;
  error?: string;
}

/** Runs the calls in order — later calls often depend on earlier writes landing. */
export async function executeToolCalls(calls: ExecutorCall[]): Promise<ExecutorResult[]> {
  const results: ExecutorResult[] = [];
  for (const call of calls) {
    if (call.error) {
      results.push({ id: call.id, name: call.name, error: call.error });
      continue;
    }
    const fn = EXECUTORS[call.name];
    if (!fn) {
      results.push({ id: call.id, name: call.name, error: `this taskpane cannot execute "${call.name}"` });
      continue;
    }
    try {
      const content = await fn(call.args || {});
      results.push({ id: call.id, name: call.name, content });
    } catch (err) {
      results.push({ id: call.id, name: call.name, error: formatExcelError(err) });
    }
  }
  return results;
}
