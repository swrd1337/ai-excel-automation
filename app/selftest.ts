/*
 * End-to-end self-test — no Excel, no DeepSeek, no network.
 *
 * A scripted model stands in for DeepSeek and an in-memory sheet stands in for
 * Excel, but the wiring between them is the real thing: the HTTP server, the
 * NDJSON chat stream, the long-poll executor channel, the server-side tool loop,
 * argument validation, capability filtering, the safety journal and the activity
 * feed.
 *
 * The only parts not covered are the Office.js calls themselves (M0 proved those
 * against real Excel) and Electron's window chrome.
 *
 * Run:  node app/selftest.js
 */

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { chmodSync, existsSync, mkdirSync, mkdtempSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { createServer, request as httpRequest } from 'node:http';
import { request as httpsRequest } from 'node:https';
import type { Server } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { createApp, buildSystemPrompt, startServer, type AppOptions, type ChatEvent } from './server.js';
import { validateArgs, toolsFor, getTool, TOOLS } from './tools.js';
import type { ConfigState, CredentialPatch } from './config.js';
import { describeCertificate, ensureCertificate, isTrusted, resolveOpenssl } from './tls.js';
import { describeManifest, installManifest, readManifestVersion, renderManifest, setupState as describeSetup } from './setup.js';
import { expandAddress, checkShape, checkTableName, EXECUTOR_NAMES, CONTEXT_TOOL } from './web/executor.js';
import { isOwnLoopbackOrigin } from './electron/cert-policy.js';
import { estimateCost, isPeak, resolveModel } from './pricing.js';
import { foldDeltas, type AiClient, type CompleteRequest, type Completion, type StreamDelta } from './ai.js';

const SESSION = 'test';

/*
 * The value the credential tests send. It has to satisfy the app's own validator
 * (`sk-` followed by at least 16 characters), so it is necessarily key-shaped —
 * and it is ASSEMBLED rather than written as one literal on purpose.
 *
 * It used to be `sk-` followed by the alphabet: obviously fake to a person, and
 * indistinguishable from a leaked key to a machine. GitHub's push protection
 * blocked the very first push of this repository over it, and the only ways out
 * are to maintain an allow-list entry or to stop writing key-shaped literals. A
 * scanner cannot be taught the difference, so this is built at runtime and there
 * is nothing for it to match.
 */
const PLACEHOLDER_KEY = ['sk', 'fixture', 'not', 'areal', 'key'].join('-');
/*
 * What a current taskpane advertises, derived from the real executor rather than
 * hand-listed. The hand-written copy drifted the moment `__fingerprint` was added
 * — and a stale list means every test exercises a taskpane that does not exist.
 */
const CAPABILITIES = [...EXECUTOR_NAMES];
// What a taskpane loaded before the safety layer existed advertises.
const STALE_CAPABILITIES = CAPABILITIES.filter((c) => !c.startsWith('__'));

let passed = 0;
let failed = 0;

function test(name: string, fn: () => void): void {
  try {
    fn();
    passed++;
    console.log(`  ok    ${name}`);
  } catch (err) {
    failed++;
    console.log(`  FAIL  ${name}\n        ${(err as Error).message}`);
  }
}

async function testAsync(name: string, fn: () => Promise<void>): Promise<void> {
  try {
    await fn();
    passed++;
    console.log(`  ok    ${name}`);
  } catch (err) {
    failed++;
    console.log(`  FAIL  ${name}\n        ${(err as Error).message}`);
  }
}

/* ------------------------------------------------------------ fake taskpane */

/** Stands in for Excel: one sheet in memory, plus any objects created over it. */
function fakeExcel() {
  const state: {
    sheets: Record<string, unknown[][]>;
    tables: Record<string, string>;
    charts: Record<string, string>;
    pivots: Record<string, string>;
    seen: string[];
  } = {
    sheets: { Sheet1: [['name', 'qty'], ['widget', 4]] },
    tables: {},
    charts: {},
    pivots: {},
    seen: [],
  };

  /** Stands in for the taskpane's fingerprint: a digest of what the sheet holds now. */
  const fingerprint = (): string => `fp:${JSON.stringify(state.sheets.Sheet1)}`;

  const api = {
    state,
    async execute(call: { name: string; args: Record<string, unknown> }): Promise<unknown> {
      state.seen.push(call.name);
      const addr = call.args.address ? String(call.args.address) : '';
      switch (call.name) {
        case '__workbook_context':
          return {
            activeSheet: 'Sheet1',
            sheets: [{ name: 'Sheet1', usedRange: 'Sheet1!A1:B2', rowCount: 2, columnCount: 2 }],
            selection: { address: 'Sheet1!A1', values: [['name', 'qty']] },
          };
        case '__snapshot': {
          const inner = (call.args.args ?? {}) as { address?: string; source?: string; name?: string };
          const tool = String(call.args.tool);

          /*
           * Mirrors the real taskpane: a created object is undone by name, not by
           * restoring cells, so its plan is a different kind of record. Checked
           * before the address guard below, because a pivot has a source and a
           * destination rather than the single address the range path needs.
           */
          const shelf =
            tool === 'create_table'
              ? state.tables
              : tool === 'create_chart'
                ? state.charts
                : tool === 'create_pivot_table'
                  ? state.pivots
                  : null;

          if (shelf) {
            const kind =
              tool === 'create_table' ? 'created-table' : tool === 'create_chart' ? 'created-chart' : 'created-pivot';
            if (!inner.name) throw new Error(`${tool} needs a name`);
            if (shelf[inner.name]) throw new Error(`"${inner.name}" already exists`);
            return {
              action: { kind, sheet: null, name: inner.name },
              label: `${inner.name} — a ${kind.replace('created-', '')}`,
              preview: state.sheets.Sheet1,
              seal: false,
            };
          }

          if (!inner.address) throw new Error('cannot tell which range this affects');
          return {
            action: { kind: 'range', sheet: null, address: inner.address, state: { formulas: state.sheets.Sheet1 } },
            label: `Sheet1!${inner.address}`,
            preview: state.sheets.Sheet1,
            seal: true,
          };
        }
        case '__fingerprint': {
          const action = call.args.action as { kind?: string } | undefined;
          return { expect: action?.kind === 'range' ? fingerprint() : null };
        }
        case '__restore': {
          const record = call.args as {
            action?: { kind: string; address?: string; state?: { formulas: unknown[][] }; name?: string };
            expect?: string | null;
          };
          const action = record.action;
          if (!action) throw new Error('the undo record carries no action');

          if (action.kind === 'range') {
            // The same rule the real taskpane enforces: never write an old
            // snapshot over cells that have moved on since.
            if (!record.expect) throw new Error('recorded without a fingerprint');
            if (record.expect !== fingerprint()) {
              throw new Error(`${action.address} has changed since this task ran`);
            }
            state.sheets.Sheet1 = action.state!.formulas;
            return { restored: action.address };
          }
          if (action.kind === 'created-table') {
            delete state.tables[action.name as string];
            // convertToRange keeps the cells; only the table goes away.
            return { removedTable: action.name, kept: 'the cells it covered' };
          }
          /*
           * Charts and pivots are deleted outright, unlike a table: a chart owns no
           * cells at all, and a pivot's cells are its own output, recomputed from
           * the source it was built on.
           */
          if (action.kind === 'created-chart') {
            delete state.charts[action.name as string];
            return { removedChart: action.name };
          }
          if (action.kind === 'created-pivot') {
            delete state.pivots[action.name as string];
            return { removedPivot: action.name, kept: 'the source data it was built from' };
          }
          throw new Error(`cannot undo a "${action.kind}" change`);
        }
        case 'get_workbook_overview':
          return { activeSheet: 'Sheet1', sheets: [{ name: 'Sheet1' }] };
        case 'read_range':
          return { address: `Sheet1!${addr}`, values: state.sheets.Sheet1 };
        case 'write_range': {
          const values = call.args.values;
          if (!Array.isArray(values) || !values.length) throw new Error('values required');
          state.sheets.Sheet1 = values as unknown[][];
          return { wroteTo: `Sheet1!${addr}`, rowCount: (values as unknown[]).length };
        }
        case 'set_formula':
          return { wroteTo: `Sheet1!${addr}`, computedValues: [[42]] };
        case 'format_range':
          return { formatted: `Sheet1!${addr}`, applied: ['bold=true'] };
        case 'sort_range': {
          const keys = (call.args.keys || []) as { column: number; ascending?: boolean }[];
          if (!keys.length) throw new Error('at least one sort key is required');
          return { sorted: `Sheet1!${addr}`, keys: keys.length, hasHeaders: call.args.hasHeaders !== false };
        }
        case 'create_table': {
          const name = String(call.args.name ?? '');
          if (!name) throw new Error('a table name is required');
          if (state.tables[name]) throw new Error(`a table named "${name}" already exists`);
          state.tables[name] = addr;
          return { table: name, address: `Sheet1!${addr}`, hasHeaders: call.args.hasHeaders !== false };
        }
        case 'create_chart': {
          const name = String(call.args.name ?? '');
          if (!name) throw new Error('a chart name is required');
          if (state.charts[name]) throw new Error(`a chart named "${name}" already exists`);
          state.charts[name] = addr;
          return { chart: name, type: call.args.type ?? 'ColumnClustered', source: addr };
        }
        case 'create_pivot_table': {
          const name = String(call.args.name ?? '');
          if (!name) throw new Error('a pivot table name is required');
          if (state.pivots[name]) throw new Error(`a pivot table named "${name}" already exists`);
          state.pivots[name] = String(call.args.destination ?? '');
          return { pivot: name, source: call.args.source, destination: call.args.destination };
        }
        case 'remove_duplicates': {
          const columns = (call.args.columns || []) as number[];
          if (!columns.length) throw new Error('at least one column is required');
          const rows = state.sheets.Sheet1;
          const header = call.args.hasHeaders !== false ? rows.slice(0, 1) : [];
          const body = call.args.hasHeaders !== false ? rows.slice(1) : rows;
          const seen = new Set<string>();
          const kept: unknown[][] = [];
          for (const row of body) {
            // Only the named columns decide a duplicate — that is the whole point
            // of the argument, so the fake must honour it or a test proves nothing
            // about a range whose other columns differ.
            const key = JSON.stringify(columns.map((c) => row[c]));
            if (seen.has(key)) continue;
            seen.add(key);
            kept.push(row);
          }
          state.sheets.Sheet1 = [...header, ...kept];
          return { removed: body.length - kept.length, remaining: kept.length, comparedColumns: columns };
        }
        default:
          throw new Error(`fake Excel cannot run "${call.name}"`);
      }
    },
  };
  return api;
}

type FakeExcel = ReturnType<typeof fakeExcel>;

/** Runs the real long-poll loop against the server, like the taskpane does. */
function startFakeTaskpane(
  base: string,
  excel: FakeExcel,
  { sessionId = SESSION, capabilities = CAPABILITIES }: { sessionId?: string; capabilities?: string[] } = {}
) {
  let running = true;
  const stats = { handled: 0, errors: [] as string[] };

  const run = async (): Promise<void> => {
    while (running) {
      const params = new URLSearchParams({
        sessionId,
        capabilities: capabilities.join(','),
        // Short hold: a 25s poll would make teardown wait 25s for the connection to drain.
        timeout: '250',
      });
      const res = await fetch(`${base}/api/executor/next?${params}`, { cache: 'no-store' });
      if (res.status === 204) continue;
      if (!res.ok) throw new Error(`poll failed: HTTP ${res.status}`);
      const { call } = (await res.json()) as { call: { id: string; name: string; args: Record<string, unknown> } };

      let content: unknown = null;
      let error: string | null = null;
      try {
        content = await excel.execute(call);
      } catch (err) {
        error = (err as Error).message;
        stats.errors.push(`${call.name}: ${error}`);
      }
      stats.handled += 1;

      await fetch(`${base}/api/executor/result`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sessionId, id: call.id, content, error }),
        cache: 'no-store',
      });
    }
  };

  const loop = run().catch((err: Error) => {
    if (running) stats.errors.push(`loop: ${err.message}`);
  });

  return {
    stats,
    async stop(): Promise<void> {
      running = false;
      await Promise.race([loop, new Promise((r) => setTimeout(r, 200))]);
    },
  };
}

/** The long-poll only registers on its first request, so wait for it. */
async function waitForExecutor(base: string, sessionId = SESSION): Promise<void> {
  for (let i = 0; i < 60; i++) {
    const res = await fetch(`${base}/api/status?sessionId=${sessionId}`, { cache: 'no-store' });
    const s = (await res.json()) as { excelConnected: boolean };
    if (s.excelConnected) return;
    await new Promise((r) => setTimeout(r, 25));
  }
  throw new Error('the fake taskpane never registered as connected');
}

/* ---------------------------------------------------------- scripted model */

function textCompletion(content: string): Completion {
  return {
    text: content,
    toolCalls: [],
    finishReason: 'stop',
    usage: null,
    assistantMessage: { role: 'assistant', content },
  };
}

function callCompletion(name: string, args: Record<string, unknown>, raw?: string): Completion {
  const id = `call_${name}_${Math.random().toString(36).slice(2, 7)}`;
  const rawArgs = raw !== undefined ? raw : JSON.stringify(args);
  let parsed: Record<string, unknown> = {};
  let argError: string | null = null;
  try {
    parsed = JSON.parse(rawArgs) as Record<string, unknown>;
  } catch (err) {
    argError = `arguments are not valid JSON (${(err as Error).message})`;
  }
  return {
    text: '',
    toolCalls: [{ id, name, args: parsed, argError, rawArgs }],
    finishReason: 'tool_calls',
    usage: null,
    assistantMessage: {
      role: 'assistant',
      content: null,
      tool_calls: [{ id, type: 'function', function: { name, arguments: rawArgs } }],
    },
  };
}

function scriptedAi(
  steps: Completion[]
): AiClient & { seen: { roles: string[]; tools: string[]; contents: string[] }[] } {
  const seen: { roles: string[]; tools: string[]; contents: string[] }[] = [];
  return {
    model: 'scripted',
    seen,
    async complete({ messages, tools, onDelta }: CompleteRequest): Promise<Completion> {
      seen.push({
        roles: messages.map((m) => m.role),
        tools: (tools || []).map((t) => t.function.name),
        contents: messages.map((m) => (typeof m.content === 'string' ? m.content : '')),
      });
      const step = steps[Math.min(seen.length - 1, steps.length - 1)] as Completion;

      // Stream the prose, as a real model would, so the delta path is exercised.
      if (step.text && onDelta) {
        for (const word of step.text.split(' ')) onDelta(`${word} `);
      }
      return step;
    },
  };
}

function fakeSearch(results: { title: string; url: string; snippet: string }[]) {
  const calls: { query: string; limit?: number }[] = [];
  return {
    provider: 'fake',
    calls,
    async search(query: string, limit?: number) {
      calls.push({ query, limit });
      return { query, provider: 'fake', answer: 'a synthesised answer', results };
    },
  };
}

/* ---------------------------------------------------------------- harness */

/**
 * A journal file of its own, in a temp directory.
 *
 * Every test gets one, because the default is the real ~/.ai-excel/journal.json
 * and a test that writes there would step on the user's ability to undo — the
 * same class of mistake as pointing a test server at the live port.
 */
function tempJournal(): string {
  return join(mkdtempSync(join(tmpdir(), 'ai-excel-journal-')), 'journal.json');
}

/**
 * A credential store in memory.
 *
 * The real one writes ~/.ai-excel/config.json, which is the user's own key. A
 * test that saved there could replace a working key with a fake one, so no test
 * ever gets the real store — the same rule as the journal and the port.
 */
function fakeCredentials(initial?: Partial<ConfigState>) {
  let state: ConfigState = {
    file: '(in memory — the real one is ~/.ai-excel/config.json)',
    deepseek: { set: true, masked: 'sk-test…key1', source: 'file', envVar: null, model: 'deepseek-chat' },
    search: { set: false, masked: null, source: null, envVar: null, provider: 'tavily' },
    ...initial,
  };
  const applied: CredentialPatch[] = [];
  return {
    applied,
    read: () => state,
    apply(patch: CredentialPatch) {
      applied.push(patch);
      if (patch.deepseekApiKey !== undefined) {
        state = {
          ...state,
          deepseek: {
            ...state.deepseek,
            set: patch.deepseekApiKey !== null,
            masked: patch.deepseekApiKey === null ? null : 'sk-new…key2',
            source: patch.deepseekApiKey === null ? null : 'file',
          },
        };
      }
      if (patch.searchApiKey !== undefined) {
        state = {
          ...state,
          search: {
            ...state.search,
            set: patch.searchApiKey !== null,
            masked: patch.searchApiKey === null ? null : 'tvly-…key3',
            source: patch.searchApiKey === null ? null : 'file',
            provider: patch.searchProvider ?? state.search.provider,
          },
        };
      }
      return { ai: replacementAi, search: null };
    },
  };
}

/** Stands in for the client built from a freshly saved key. */
const replacementAi: AiClient = {
  model: 'replacement-model',
  async complete() {
    return { text: 'replaced', toolCalls: [], finishReason: 'stop', usage: null, assistantMessage: { role: 'assistant', content: 'replaced' } };
  },
};

/**
 * A request that tolerates the app's own certificate.
 *
 * Node's `fetch` rejects a self-signed chain, and `startServer` genuinely serves
 * the certificate it found on this machine — so a test of what it bound has to
 * speak TLS without being able to verify it. That is the point of the test: it is
 * checking which addresses answer, not who signed them.
 */
function rawStatus(url: string): Promise<number> {
  return new Promise((resolve, reject) => {
    const send = url.startsWith('https') ? httpsRequest : httpRequest;
    const req = send(url, { rejectUnauthorized: false }, (res) => {
      res.resume();
      resolve(res.statusCode ?? 0);
    });
    req.on('error', reject);
    req.end();
  });
}

async function startApp({
  ai,
  search = null,
  journalFile = tempJournal(),
  credentials,
  setupped,
}: {
  ai: AiClient;
  search?: ReturnType<typeof fakeSearch> | null;
  journalFile?: string;
  credentials?: AppOptions['credentials'];
  setupped?: AppOptions['setupped'];
}) {
  const app = createApp({ ai, search, journalFile, credentials, setupped });
  const server = createServer((req, res) => {
    void app(req, res);
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  return { server, base: `http://127.0.0.1:${(server.address() as { port: number }).port}` };
}

/**
 * server.close() waits for in-flight requests, and a held long-poll is exactly
 * that. Destroy connections first so teardown is instant.
 */
async function stopServer(server: Server): Promise<void> {
  server.closeAllConnections?.();
  await new Promise((r) => server.close(r));
}

async function post(base: string, path: string, body: unknown): Promise<{ status: number; body: any }> {
  const res = await fetch(`${base}${path}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  return { status: res.status, body: await res.json() };
}

/** Consumes the NDJSON chat stream, exactly as the window does. */
async function chatEvents(
  base: string,
  message: string,
  { review = false, search = false, sessionId = SESSION }: { review?: boolean; search?: boolean; sessionId?: string } = {}
): Promise<ChatEvent[]> {
  const res = await fetch(`${base}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ sessionId, message, review, search }),
  });

  if (!res.ok) {
    const body = (await res.json()) as { error?: string };
    return [{ type: 'error', error: body.error || `HTTP ${res.status}` }];
  }

  const events: ChatEvent[] = [];
  const decoder = new TextDecoder();
  let buffer = '';

  for await (const chunk of res.body as unknown as AsyncIterable<Uint8Array>) {
    buffer += decoder.decode(chunk, { stream: true });
    const lines = buffer.split('\n');
    buffer = lines.pop() ?? '';
    for (const line of lines) {
      if (line.trim()) events.push(JSON.parse(line) as ChatEvent);
    }
  }
  return events;
}

const replyOf = (events: ChatEvent[]): string =>
  (events.find((e) => e.type === 'done') as { reply?: string } | undefined)?.reply ?? '';
const errorOf = (events: ChatEvent[]): string =>
  (events.find((e) => e.type === 'error') as { error?: string } | undefined)?.error ?? '';
const deltasOf = (events: ChatEvent[]): string =>
  events
    .filter((e): e is Extract<ChatEvent, { type: 'delta' }> => e.type === 'delta')
    .map((e) => e.text)
    .join('');
const toolEvents = (events: ChatEvent[]) => events.filter((e) => e.type === 'tool');

const status = async (base: string, sessionId = SESSION): Promise<any> =>
  (await fetch(`${base}/api/status?sessionId=${sessionId}`, { cache: 'no-store' })).json();

/** The turn parks on an approval, so wait for it to appear before answering. */
async function waitForApproval(base: string, sessionId = SESSION): Promise<any> {
  for (let i = 0; i < 120; i++) {
    const res = await fetch(`${base}/api/approvals?sessionId=${sessionId}`, { cache: 'no-store' });
    const { approvals } = (await res.json()) as { approvals: any[] };
    if (approvals.length) return approvals[0];
    await new Promise((r) => setTimeout(r, 25));
  }
  throw new Error('no approval request arrived');
}

/* -------------------------------------------------------------------- tests */

console.log('\npure helpers');

test('expandAddress fills outward from a single-cell anchor', () => {
  assert.equal(expandAddress('A1', 3, 2), 'A1:B3');
  assert.equal(expandAddress('B2', 1, 1), 'B2:B2');
  assert.equal(expandAddress('AA10', 2, 3), 'AA10:AC11');
});

test('expandAddress leaves an explicit range untouched', () => {
  assert.equal(expandAddress('A1:C3', 9, 9), 'A1:C3');
  assert.equal(expandAddress('$A$1:$B$2', 5, 5), '$A$1:$B$2');
});

test('validateArgs accepts good args and rejects each kind of bad one', () => {
  const schema = getTool('write_range')!.parameters;
  assert.deepEqual(validateArgs(schema, { address: 'A1', values: [['x']] }), []);
  assert.ok(validateArgs(schema, { values: [['x']] }).some((e) => e.includes('address is required')));
  assert.ok(validateArgs(schema, { address: 'A1', values: 'nope' }).some((e) => e.includes('must be an array')));
  assert.ok(validateArgs(schema, { address: 'A1', values: [['x']], bogus: 1 }).some((e) => e.includes('bogus')));
});

test('the taskpane can run every internal call the server dispatches', () => {
  /*
   * Hard-coded names on purpose: this is the other side of the wire from
   * `app/server.ts`, so renaming one without the other fails here rather than at
   * a user's expense — which is exactly what invariant 4 is about.
   *
   * `__workbook_context` is deliberately not an EXECUTORS entry, because the
   * taskpane shell answers it directly so a per-turn context read is not logged
   * as workbook activity. That is why it is checked separately.
   */
  assert.equal(CONTEXT_TOOL, '__workbook_context');
  for (const name of ['__snapshot', '__fingerprint', '__restore']) {
    assert.ok(EXECUTOR_NAMES.includes(name), `the taskpane has no "${name}", so the server's call would always fail`);
  }

  // And the other direction: an internal helper must never reach the model.
  assert.ok(!toolsFor({ capabilities: CAPABILITIES, serverTools: [] }).some((t) => t.function.name.startsWith('__')));
});

test('toolsFor respects capabilities and only offers server tools when configured', () => {
  const withSearch = toolsFor({ capabilities: ['read_range'], serverTools: ['web_search'] }).map((t) => t.function.name);
  assert.deepEqual(withSearch.sort(), ['read_range', 'web_search']);

  const withoutSearch = toolsFor({ capabilities: ['read_range'], serverTools: [] }).map((t) => t.function.name);
  assert.deepEqual(withoutSearch, ['read_range']);

  const noCaps = toolsFor({ capabilities: [], serverTools: [] }).map((t) => t.function.name);
  assert.ok(noCaps.includes('write_range'), 'with no capabilities advertised, all taskpane tools are offered');
});

test('checkShape explains a range/values mismatch in actionable terms', () => {
  assert.equal(checkShape('A1', 2, 2), null, 'a single-cell anchor expands, so it is always fine');
  assert.equal(checkShape('A1:B2', 2, 2), null, 'an exact match is fine');

  const err = checkShape('A1:B2', 3, 2);
  assert.match(err as string, /A1:B2 is 2x2 but 3x2 values/);
  assert.match(err as string, /single-cell anchor/, 'the model needs to be told how to fix it');
});

test('checkTableName rejects what Excel would reject with an unusable message', () => {
  assert.equal(checkTableName('SalesTable'), null);
  assert.equal(checkTableName('_q1.2026'), null, 'underscores and full stops are allowed');

  // Office.js answers all of these with "The argument is invalid or missing or
  // has an incorrect format", which names neither the rule nor the character.
  assert.match(checkTableName('Sales Table') as string, /no spaces/);
  assert.match(checkTableName('A1') as string, /cell reference/);
  assert.match(checkTableName('') as string, /required/);
});

test('the certificate exception applies ONLY to our own loopback origin', () => {
  assert.equal(isOwnLoopbackOrigin('https://localhost:3443/ui.html', 3443), true);
  assert.equal(isOwnLoopbackOrigin('https://127.0.0.1:3443/ui.html', 3443), true);

  // A different local port is a different service, so it is not covered.
  assert.equal(isOwnLoopbackOrigin('https://localhost:9999/ui.html', 3443), false);

  // Suffix tricks must not slip through: exact host match only.
  assert.equal(isOwnLoopbackOrigin('https://localhost.evil.com:3443/x', 3443), false);
  assert.equal(isOwnLoopbackOrigin('https://notlocalhost:3443/x', 3443), false);

  // A wildcard bind address is not loopback.
  assert.equal(isOwnLoopbackOrigin('https://0.0.0.0:3443/x', 3443), false);

  // Remote hosts and non-web schemes fail closed.
  assert.equal(isOwnLoopbackOrigin('https://example.com/ui.html', 3443), false);
  assert.equal(isOwnLoopbackOrigin('file:///etc/passwd', 3443), false);
  assert.equal(isOwnLoopbackOrigin('not a url', 3443), false);
});

test('the system prompt never claims a capability the model does not have', () => {
  const withSearch = buildSystemPrompt({ hasSearch: true });
  assert.match(withSearch, /You have web_search/);
  assert.doesNotMatch(withSearch, /NO web access/);

  const without = buildSystemPrompt({ hasSearch: false });
  assert.doesNotMatch(without, /You have web_search/, 'must not claim search when it is disabled');
  assert.match(without, /NO web access/);
  assert.match(without, /ask the user to supply it/, 'it should be honest about the gap');
});

console.log('\nstreaming');

test('streamed text deltas concatenate in order', () => {
  const folded = foldDeltas([
    { content: 'Added ' },
    { content: 'a header ' },
    { content: 'row.' },
    { finish_reason: 'stop' },
  ]);
  assert.equal(folded.text, 'Added a header row.');
  assert.equal(folded.finishReason, 'stop');
});

test('tool-call arguments are reassembled from streamed fragments', () => {
  // This is the shape DeepSeek actually sends: the name arrives in pieces and
  // the arguments JSON is dribbled out a few characters at a time.
  const folded = foldDeltas([
    { tool_calls: [{ index: 0, id: 'call_1', function: { name: 'write_', arguments: '{"add' } }] },
    { tool_calls: [{ index: 0, function: { name: 'range', arguments: 'ress":"A1",' } }] },
    { tool_calls: [{ index: 0, function: { arguments: '"values":[["x"]]}' } }] },
    { finish_reason: 'tool_calls' },
  ]);

  assert.equal(folded.toolCalls.length, 1);
  assert.equal(folded.toolCalls[0]!.name, 'write_range', 'a split name must be joined, not truncated');
  assert.deepEqual(JSON.parse(folded.toolCalls[0]!.rawArgs), { address: 'A1', values: [['x']] });
  assert.equal(folded.finishReason, 'tool_calls');
});

test('two concurrent tool calls are kept apart by index', () => {
  const folded = foldDeltas([
    {
      tool_calls: [
        { index: 0, id: 'a', function: { name: 'read_range', arguments: '{"address":"A1"}' } },
        { index: 1, id: 'b', function: { name: 'get_selection', arguments: '{}' } },
      ],
    },
  ]);
  assert.deepEqual(
    folded.toolCalls.map((c) => c.name),
    ['read_range', 'get_selection']
  );
});

test('a delta stream with no tool calls yields none', () => {
  const folded = foldDeltas([{ content: 'hello' }]);
  assert.deepEqual(folded.toolCalls, []);
});

console.log('\npricing');

test('peak hours are 01-04 and 06-10 UTC on weekdays only', () => {
  assert.equal(isPeak(new Date('2026-09-23T02:00:00Z')), true, 'Wed 02:00 UTC is peak');
  assert.equal(isPeak(new Date('2026-09-23T07:00:00Z')), true, 'Wed 07:00 UTC is peak');
  assert.equal(isPeak(new Date('2026-09-23T05:00:00Z')), false, 'the 04:00-06:00 gap is off-peak');
  assert.equal(isPeak(new Date('2026-09-23T12:00:00Z')), false, 'midday is off-peak');
  assert.equal(isPeak(new Date('2026-09-26T02:00:00Z')), false, 'Saturday is off-peak all day');
  assert.equal(isPeak(new Date('2026-09-27T07:00:00Z')), false, 'Sunday is off-peak all day');
});

test('cost is charged per token class, and doubles at peak', () => {
  const offPeak = new Date('2026-09-23T12:00:00Z');
  const peak = new Date('2026-09-23T02:00:00Z');
  const million = { promptTokens: 1_000_000, cacheHitTokens: 0, cacheMissTokens: 1_000_000, completionTokens: 1_000_000 };

  // Flash: $0.15 per 1M cache-miss input + $0.60 per 1M output at off-peak.
  const off = estimateCost(million, 'deepseek-flash', offPeak);
  assert.equal(Number(off.usd.toFixed(6)), 0.75);

  const on = estimateCost(million, 'deepseek-flash', peak);
  assert.equal(Number(on.usd.toFixed(6)), 1.5, 'peak is exactly double');
});

test('cached input is charged at the much lower cache-hit rate', () => {
  const at = new Date('2026-09-23T12:00:00Z');
  const cached = estimateCost(
    { promptTokens: 1_000_000, cacheHitTokens: 1_000_000, cacheMissTokens: 0, completionTokens: 0 },
    'deepseek-flash',
    at
  );
  assert.equal(Number(cached.usd.toFixed(6)), 0.003);
});

test('missing cache reporting is treated as cache-miss, not as free', () => {
  const at = new Date('2026-09-23T12:00:00Z');
  const unknown = estimateCost({ promptTokens: 1_000_000, completionTokens: 0 }, 'deepseek-flash', at);
  assert.equal(Number(unknown.usd.toFixed(6)), 0.15, 'must over-estimate rather than flatter the number');
});

test('an unpublished model name is resolved but flagged as an estimate', () => {
  assert.deepEqual(resolveModel('deepseek-flash'), { key: 'deepseek-flash', exact: true });
  assert.deepEqual(resolveModel('deepseek-chat'), { key: 'deepseek-flash', exact: false });
  assert.equal(estimateCost({}, 'something-new', new Date()).exact, false);
});

console.log('\nexecutor channel');

await testAsync('a message flows to Excel through the channel and back', async () => {
  const ai = scriptedAi([
    callCompletion('write_range', { address: 'A1', values: [['Hello', 1]] }),
    textCompletion('Wrote Hello into A1:B1.'),
  ]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    const events = await chatEvents(base, 'put Hello in A1');
    assert.equal(errorOf(events), '', 'unexpected error');
    assert.match(replyOf(events), /Wrote Hello/);
    assert.deepEqual(excel.state.sheets.Sheet1, [['Hello', 1]], 'the fake sheet must actually change');
    assert.ok(excel.state.seen.includes('__workbook_context'), 'context is fetched from Excel, not the UI');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('the reply streams as deltas, and a done event carries the final text', async () => {
  const ai = scriptedAi([
    callCompletion('write_range', { address: 'A1', values: [['x']] }),
    textCompletion('Added a header row in A1.'),
  ]);
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, fakeExcel());
  try {
    await waitForExecutor(base);
    const events = await chatEvents(base, 'add a header');

    const deltas = deltasOf(events);
    assert.match(deltas, /Added a header row/, 'the text must arrive as deltas, not only at the end');
    assert.equal(deltas.trim(), replyOf(events).trim(), 'the deltas must reconstruct the final reply');
    assert.ok(events.some((e) => e.type === 'assistant_start'), 'a bubble must be opened before deltas');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('web_search resolves on the server and never reaches Excel', async () => {
  const ai = scriptedAi([
    callCompletion('web_search', { query: 'top selling EVs 2026' }),
    callCompletion('write_range', { address: 'A1', values: [['EV', 'Units']] }),
    textCompletion('Searched and wrote the results into A1:B1.'),
  ]);
  const excel = fakeExcel();
  const search = fakeSearch([{ title: 'EV sales', url: 'https://example.com', snippet: 'Model Y leads' }]);
  const { server, base } = await startApp({ ai, search });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    const events = await chatEvents(base, 'find the top selling EVs and put them in the sheet', { search: true });
    assert.match(replyOf(events), /Searched and wrote/);
    assert.equal(search.calls.length, 1, 'the search should have run once');
    assert.equal(search.calls[0]!.query, 'top selling EVs 2026');
    assert.ok(!excel.state.seen.includes('web_search'), 'Excel must never be asked to run a server tool');
    assert.ok(
      events.some((e) => e.type === 'search'),
      'the UI should be told a search happened'
    );

    const offered = ai.seen[0]!.tools;
    assert.ok(offered.includes('web_search') && offered.includes('write_range'), 'the model must see both kinds of tool');
    assert.equal(
      (await status(base)).usage.searches,
      1,
      'a search costs no tokens, so it must be counted separately or it looks free'
    );
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('web_search stays off unless the request asks for it, even with a key', async () => {
  const ai = scriptedAi([textCompletion('Answered without searching.')]);
  const { server, base } = await startApp({ ai, search: fakeSearch([]) });
  const taskpane = startFakeTaskpane(base, fakeExcel());
  try {
    await waitForExecutor(base);
    // A configured provider is not consent to use it.
    await chatEvents(base, 'look something up');

    assert.ok(
      !ai.seen[0]!.tools.includes('web_search'),
      'a configured search key must not enable search on its own'
    );
    assert.doesNotMatch(
      buildSystemPrompt({ hasSearch: false }),
      /You have web_search/,
      'the prompt and the tool list must agree'
    );
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('a web_search call is refused when search is off', async () => {
  const ai = scriptedAi([callCompletion('web_search', { query: 'sneaky' }), textCompletion('Answered directly.')]);
  const search = fakeSearch([]);
  const { server, base } = await startApp({ ai, search });
  const taskpane = startFakeTaskpane(base, fakeExcel());
  try {
    await waitForExecutor(base);
    await chatEvents(base, 'look something up');

    assert.equal(search.calls.length, 0, 'the provider must not be called when search is off');
    assert.ok(
      !(await status(base)).usage.searches,
      'a refused search must not be counted as one'
    );
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('the model is not offered web_search when no search key is configured', async () => {
  const ai = scriptedAi([textCompletion('I cannot look that up.')]);
  const { server, base } = await startApp({ ai, search: null });
  const taskpane = startFakeTaskpane(base, fakeExcel());
  try {
    await waitForExecutor(base);
    await chatEvents(base, 'look something up', { search: true });
    assert.ok(!ai.seen[0]!.tools.includes('web_search'), 'web_search must not be offered without a provider');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('with no taskpane connected the user gets a clear error, not a hang', async () => {
  const ai = scriptedAi([textCompletion('unused')]);
  const { server, base } = await startApp({ ai });
  try {
    const events = await chatEvents(base, 'do something');
    assert.match(errorOf(events), /Excel is not connected/i);
    assert.match(errorOf(events), /add-in/i, 'the message should say how to fix it');
  } finally {
    await stopServer(server);
  }
});

await testAsync('invalid arguments are never dispatched to Excel, and the model can repair them', async () => {
  const ai = scriptedAi([
    callCompletion('write_range', { values: [['oops']] }), // no address
    callCompletion('write_range', { address: 'B2', values: [['fixed']] }),
    textCompletion('Corrected.'),
  ]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    const events = await chatEvents(base, 'write something');
    assert.match(replyOf(events), /Corrected/);
    const writes = excel.state.seen.filter((n) => n === 'write_range');
    assert.equal(writes.length, 1, 'only the repaired call should reach Excel');
    assert.deepEqual(excel.state.sheets.Sheet1, [['fixed']]);
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('malformed tool-call JSON is caught and never dispatched', async () => {
  const ai = scriptedAi([callCompletion('write_range', {}, '{"address": "A1", '), textCompletion('Recovered.')]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    await chatEvents(base, 'write');
    assert.ok(!excel.state.seen.includes('write_range'), 'malformed args must not reach Excel');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('a tool outside the catalog never reaches Excel', async () => {
  const ai = scriptedAi([callCompletion('delete_everything', {}), textCompletion('Refused.')]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel, { capabilities: [...CAPABILITIES, 'delete_everything'] });
  try {
    await waitForExecutor(base);
    await chatEvents(base, 'destroy it all');
    assert.ok(!excel.state.seen.includes('delete_everything'), 'an unknown tool must not run');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('the search budget is enforced', async () => {
  const ai = scriptedAi([callCompletion('web_search', { query: 'forever' })]); // never stops
  const excel = fakeExcel();
  const search = fakeSearch([]);
  const { server, base } = await startApp({ ai, search });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    const events = await chatEvents(base, 'search forever', { search: true });
    assert.ok(search.calls.length <= 3, `expected at most 3 searches, got ${search.calls.length}`);
    assert.match(replyOf(events), /could not finish|stopped/i);
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('Excel failing a tool does not break the turn', async () => {
  const ai = scriptedAi([
    callCompletion('read_range', { address: 'ZZZ999' }),
    textCompletion('That range could not be read.'),
  ]);
  const excel = fakeExcel();
  excel.execute = async (call) => {
    if (call.name === '__workbook_context') {
      return {
        activeSheet: 'Sheet1',
        sheets: [{ name: 'Sheet1', usedRange: 'Sheet1!A1:B2', rowCount: 2, columnCount: 2 }],
      };
    }
    throw new Error('InvalidArgument: the range is not valid');
  };
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    const events = await chatEvents(base, 'read a nonsense range');
    assert.match(replyOf(events), /could not be read/);
    assert.ok(
      toolEvents(events).some((e) => (e as { ok: boolean }).ok === false),
      'the failure must be reported to the UI'
    );
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

console.log('\npreview and undo');

const ORIGINAL = [
  ['name', 'qty'],
  ['widget', 4],
];

await testAsync('a write is snapshotted first, and undo puts the original back', async () => {
  const ai = scriptedAi([
    callCompletion('write_range', { address: 'A1', values: [['CHANGED', 99]] }),
    textCompletion('Wrote it.'),
  ]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    assert.match(replyOf(await chatEvents(base, 'change A1')), /Wrote it/);
    assert.deepEqual(excel.state.sheets.Sheet1, [['CHANGED', 99]]);
    assert.equal((await status(base)).undoAvailable, 1, 'the task should be undoable');

    const undo = await post(base, '/api/undo', { sessionId: SESSION });
    assert.equal(undo.body.undone, 1);
    assert.deepEqual(excel.state.sheets.Sheet1, ORIGINAL, 'undo must restore the original values');
    assert.equal((await status(base)).undoAvailable, 0, 'the journal should be empty after undo');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('a snapshot is taken BEFORE the write, not after', async () => {
  const ai = scriptedAi([
    callCompletion('write_range', { address: 'A1', values: [['CHANGED']] }),
    textCompletion('done'),
  ]);
  const excel = fakeExcel();
  const order: string[] = [];
  const realExecute = excel.execute.bind(excel);
  excel.execute = async (call) => {
    order.push(call.name);
    return realExecute(call);
  };
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    await chatEvents(base, 'change it');
    const snap = order.indexOf('__snapshot');
    const write = order.indexOf('write_range');
    assert.ok(snap !== -1 && write !== -1, `expected both calls, saw ${order.join(', ')}`);
    assert.ok(snap < write, `snapshot must precede the write, order was ${order.join(', ')}`);
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('a read-only turn is not journalled, so undo always steps over real changes', async () => {
  const ai = scriptedAi([callCompletion('read_range', { address: 'A1' }), textCompletion('Here it is.')]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    await chatEvents(base, 'what is in A1?');
    assert.equal((await status(base)).undoAvailable, 0);
    const undo = await post(base, '/api/undo', { sessionId: SESSION });
    assert.equal(undo.body.undone, 0);
    assert.match(undo.body.message, /nothing to undo/i);
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('review mode applies a change once approved', async () => {
  const ai = scriptedAi([
    callCompletion('write_range', { address: 'A1', values: [['APPROVED']] }),
    textCompletion('Applied it.'),
  ]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);

    const pending = chatEvents(base, 'change it', { review: true });
    const approval = await waitForApproval(base);
    assert.equal(approval.tool, 'write_range');
    assert.deepEqual(approval.before, ORIGINAL, 'the card must show what the cells hold now');
    assert.deepEqual(approval.args.values, [['APPROVED']]);

    await post(base, '/api/approvals/decision', { sessionId: SESSION, id: approval.id, decision: 'approve' });
    const events = await pending;

    assert.match(replyOf(events), /Applied it/);
    assert.deepEqual(excel.state.sheets.Sheet1, [['APPROVED']]);
    assert.equal((await status(base)).undoAvailable, 1, 'an approved change is still undoable');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('review mode blocks a change that is declined', async () => {
  const ai = scriptedAi([
    callCompletion('write_range', { address: 'A1', values: [['SHOULD NOT LAND']] }),
    textCompletion('Understood, leaving it alone.'),
  ]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);

    const pending = chatEvents(base, 'change it', { review: true });
    const approval = await waitForApproval(base);
    await post(base, '/api/approvals/decision', { sessionId: SESSION, id: approval.id, decision: 'decline' });
    const events = await pending;

    assert.match(replyOf(events), /leaving it alone/);
    assert.deepEqual(excel.state.sheets.Sheet1, ORIGINAL, 'a declined change must not be applied');
    assert.ok(!excel.state.seen.includes('write_range'), 'the write must never be dispatched');
    assert.equal((await status(base)).undoAvailable, 0, 'a declined change must not be journalled');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('an approval that expires is reported as a silence, not as a refusal', async () => {
  /*
   * This path used to be unreachable in a test — the timeout was a 120-second
   * constant — and it was wrong in a way nobody could see: an expired approval
   * told the model "The user declined this change", so the model stopped and
   * asked what the user would prefer, for a change the user most likely never
   * saw. Shortening the timeout is what makes it checkable.
   */
  const previous = process.env.AI_EXCEL_APPROVAL_TIMEOUT_MS;
  process.env.AI_EXCEL_APPROVAL_TIMEOUT_MS = '40';

  const ai = scriptedAi([
    callCompletion('write_range', { address: 'A1', values: [['SHOULD NOT LAND']] }),
    textCompletion('That approval expired; shall I try again?'),
  ]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);

    const events = await chatEvents(base, 'change it', { review: true });
    const expired = events.find((e) => e.type === 'approval');
    assert.ok(expired, `expected an approval event, saw ${events.map((e) => e.type).join(', ')}`);
    assert.equal(expired.ok, false);
    assert.match(String(expired.detail), /expired/, `the window must say it expired: ${expired.detail}`);
    assert.doesNotMatch(String(expired.detail), /declin/, 'nobody declined it');

    assert.deepEqual(excel.state.sheets.Sheet1, ORIGINAL, 'an expired approval must not write');
    assert.ok(!excel.state.seen.includes('write_range'), 'the write must never be dispatched');
    assert.equal((await status(base)).undoAvailable, 0, 'an expired change must not be journalled');

    /*
     * And what the model was told on its second call — the part that was wrong.
     * It has to know the difference, because the two call for different replies.
     */
    const told = ai.seen[ai.seen.length - 1]?.contents.join('\n') ?? '';
    assert.match(told, /expired/, 'the model must be told the approval expired');
    assert.doesNotMatch(told, /user declined/i, 'the model must NOT be told the user declined');
  } finally {
    if (previous === undefined) delete process.env.AI_EXCEL_APPROVAL_TIMEOUT_MS;
    else process.env.AI_EXCEL_APPROVAL_TIMEOUT_MS = previous;
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('if the range cannot be snapshotted the write is refused, not attempted', async () => {
  const ai = scriptedAi([
    callCompletion('write_range', { address: 'EXPLODE', values: [['x']] }),
    textCompletion('I could not make that change safely.'),
  ]);
  const excel = fakeExcel();
  const realExecute = excel.execute.bind(excel);
  excel.execute = async (call) => {
    if (call.name === '__snapshot') throw new Error('simulated read failure');
    return realExecute(call);
  };
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    const events = await chatEvents(base, 'write to a broken range');
    assert.match(replyOf(events), /could not make that change safely/);
    assert.deepEqual(excel.state.sheets.Sheet1, ORIGINAL, 'nothing should have been written');
    assert.ok(!excel.state.seen.includes('write_range'), 'an unundoable change must not run at all');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('a taskpane too old to snapshot is refused, never written to', async () => {
  const ai = scriptedAi([
    callCompletion('write_range', { address: 'A1', values: [['SHOULD NOT LAND']] }),
    textCompletion('I could not apply that.'),
  ]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel, { capabilities: STALE_CAPABILITIES });
  try {
    await waitForExecutor(base);
    const events = await chatEvents(base, 'write something');

    assert.ok(!excel.state.seen.includes('__snapshot'), 'must not even ask a pane that cannot snapshot');
    assert.ok(!excel.state.seen.includes('write_range'), 'must not write without a way to undo it');
    assert.deepEqual(excel.state.sheets.Sheet1, ORIGINAL);

    assert.ok(
      toolEvents(events).some((e) => /out of date/.test((e as { detail?: string }).detail || '')),
      'the stream the user reads must say to reload the add-in'
    );
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('the build id is exposed so a stale taskpane can detect itself', async () => {
  const { server, base } = await startApp({ ai: scriptedAi([textCompletion('x')]) });
  try {
    const res = await fetch(`${base}/api/version`);
    const { build } = (await res.json()) as { build: string };
    assert.match(build, /^[0-9a-f]{12}$/, `expected a 12-char hash, got ${build}`);
    assert.equal(res.headers.get('X-App-Build'), build, 'every response must carry the build id');
  } finally {
    await stopServer(server);
  }
});

await testAsync('sort_range is dispatched, snapshotted first, and undoable', async () => {
  const ai = scriptedAi([
    callCompletion('sort_range', { address: 'A1:B3', keys: [{ column: 1, ascending: false }] }),
    textCompletion('Sorted by the second column, descending.'),
  ]);
  const excel = fakeExcel();
  const order: string[] = [];
  const realExecute = excel.execute.bind(excel);
  excel.execute = async (call) => {
    order.push(call.name);
    return realExecute(call);
  };

  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    assert.match(replyOf(await chatEvents(base, 'sort by qty descending')), /Sorted by the second column/);

    const snap = order.indexOf('__snapshot');
    const sort = order.indexOf('sort_range');
    assert.ok(snap !== -1 && sort !== -1, `expected both calls, saw ${order.join(', ')}`);
    assert.ok(snap < sort, 'a sort must be snapshotted before it runs — it permutes values in place');
    assert.equal((await status(base)).undoAvailable, 1, 'a sort should be undoable');

    const undo = await post(base, '/api/undo', { sessionId: SESSION });
    assert.equal(undo.body.undone, 1);
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('create_table records its undo by name, and undo removes the table but keeps the cells', async () => {
  const ai = scriptedAi([
    callCompletion('create_table', { address: 'A1:B2', name: 'SalesTable' }),
    textCompletion('A1:B2 is now the table SalesTable.'),
  ]);
  const excel = fakeExcel();
  const order: string[] = [];
  const realExecute = excel.execute.bind(excel);
  excel.execute = async (call) => {
    order.push(call.name);
    return realExecute(call);
  };
  const cells = JSON.stringify(excel.state.sheets.Sheet1);

  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    assert.match(replyOf(await chatEvents(base, 'make A1:B2 a table called SalesTable')), /SalesTable/);

    assert.deepEqual(Object.keys(excel.state.tables), ['SalesTable'], 'the table must exist after the turn');

    const snap = order.indexOf('__snapshot');
    const made = order.indexOf('create_table');
    assert.ok(snap !== -1 && made !== -1, `expected both calls, saw ${order.join(', ')}`);
    assert.ok(snap < made, 'the undo plan must be written down before the table is made — it names the object');

    const undo = await post(base, '/api/undo', { sessionId: SESSION });
    assert.equal(undo.body.undone, 1, 'a created table must be undoable');
    assert.deepEqual(Object.keys(excel.state.tables), [], 'undo must take the table away');
    assert.equal(
      JSON.stringify(excel.state.sheets.Sheet1),
      cells,
      'undo must convert the table back to a range, not delete the data under it'
    );
    assert.equal((await status(base)).undoAvailable, 0, 'the journal must be empty after undo');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('create_chart names its undo before the chart exists, and undo deletes it', async () => {
  const ai = scriptedAi([
    callCompletion('create_chart', { address: 'A1:B2', name: 'SalesChart', type: 'BarClustered' }),
    textCompletion('Added SalesChart.'),
  ]);
  const excel = fakeExcel();
  const order: string[] = [];
  const realExecute = excel.execute.bind(excel);
  excel.execute = async (call) => {
    order.push(call.name);
    return realExecute(call);
  };
  const cells = JSON.stringify(excel.state.sheets.Sheet1);

  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    assert.match(replyOf(await chatEvents(base, 'chart A1:B2 as SalesChart')), /SalesChart/);
    assert.deepEqual(Object.keys(excel.state.charts), ['SalesChart']);

    const snap = order.indexOf('__snapshot');
    const made = order.indexOf('create_chart');
    assert.ok(snap !== -1 && made !== -1, `expected both calls, saw ${order.join(', ')}`);
    assert.ok(snap < made, 'the plan names the chart, so it has to be written down before the chart exists');

    const undo = await post(base, '/api/undo', { sessionId: SESSION });
    assert.equal(undo.body.undone, 1, 'a created chart must be undoable');
    assert.deepEqual(Object.keys(excel.state.charts), [], 'undo must delete the chart');
    assert.equal(
      JSON.stringify(excel.state.sheets.Sheet1),
      cells,
      'a chart floats over the cells, so deleting it must not touch them'
    );
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('create_pivot_table is undoable by name, and a duplicate name is refused up front', async () => {
  const ai = scriptedAi([
    callCompletion('create_pivot_table', { name: 'ByRegion', source: 'A1:B2', destination: 'D1' }),
    textCompletion('Built ByRegion.'),
    // The same name again: the planner must refuse before anything is created.
    callCompletion('create_pivot_table', { name: 'ByRegion', source: 'A1:B2', destination: 'H1' }),
    textCompletion('That name was taken.'),
  ]);
  const excel = fakeExcel();

  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    await chatEvents(base, 'pivot A1:B2 into D1 as ByRegion');
    assert.deepEqual(Object.keys(excel.state.pivots), ['ByRegion']);

    const second = await chatEvents(base, 'do that again at H1');
    assert.match(JSON.stringify(second), /already exists|error/i, 'a taken name must be refused, not silently reused');
    assert.equal(excel.state.pivots.ByRegion, 'D1', 'the refused call must not have moved the pivot');

    const undo = await post(base, '/api/undo', { sessionId: SESSION });
    assert.equal(undo.body.undone, 1, 'a created pivot must be undoable');
    assert.deepEqual(Object.keys(excel.state.pivots), [], 'undo must delete the pivot');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('remove_duplicates is snapshotted first, and its undo restores the removed rows', async () => {
  /*
   * The undo here is the ordinary range snapshot, which is only honest because
   * Diagnostics measured the row shift as CONFINED to the range — "Remove-duplicates
   * shift scope". This test covers the wiring; that check is what covers Office.js.
   */
  const before = [
    ['name', 'qty'],
    ['widget', 4],
    ['gadget', 2],
    ['widget', 9],
  ];
  const ai = scriptedAi([
    callCompletion('remove_duplicates', { address: 'A1:B4', columns: [0] }),
    textCompletion('Removed the duplicate widgets.'),
  ]);
  const excel = fakeExcel();
  excel.state.sheets.Sheet1 = before.map((r) => [...r]);
  const order: string[] = [];
  const realExecute = excel.execute.bind(excel);
  excel.execute = async (call) => {
    order.push(call.name);
    return realExecute(call);
  };

  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    await chatEvents(base, 'remove duplicate rows by name');

    assert.deepEqual(
      excel.state.sheets.Sheet1,
      [['name', 'qty'], ['widget', 4], ['gadget', 2]],
      'the header and the first of each duplicate must survive, keyed on column 0 only'
    );

    const snap = order.indexOf('__snapshot');
    const dedupe = order.indexOf('remove_duplicates');
    assert.ok(snap !== -1 && dedupe !== -1, `expected both calls, saw ${order.join(', ')}`);
    assert.ok(snap < dedupe, 'the snapshot must be taken before the rows are removed');

    const undo = await post(base, '/api/undo', { sessionId: SESSION });
    assert.equal(undo.body.undone, 1, 'remove_duplicates must be undoable');
    assert.deepEqual(excel.state.sheets.Sheet1, before, 'undo must put the removed row back');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

test('every mutating tool in the catalog has an executor and an undo plan', () => {
  /*
   * Invariant 4: the schema is the contract. A tool the model can call but the
   * taskpane cannot run — or cannot undo — is a bug that only shows up mid-turn,
   * in front of the user, with a workbook already changed.
   */
  const described = TOOLS.filter((t) => t.executesOn !== 'server').map((t) => t.name);
  for (const name of described) {
    assert.ok(EXECUTOR_NAMES.includes(name), `"${name}" is offered to the model but the taskpane has no executor`);
  }

  /*
   * And the undo side: every tool that mutates must be one the planner knows how
   * to describe. `captureRange` is the fallback, so anything the range snapshot
   * genuinely covers is fine — this asserts the created-object tools are not
   * silently falling through to it, which would produce a snapshot that restores
   * cell values while leaving the object behind.
   */
  const created = ['create_table', 'create_chart', 'create_pivot_table'];
  for (const name of created) {
    const describedTool = TOOLS.find((t) => t.name === name);
    assert.ok(describedTool, `"${name}" should be in the catalog`);
    assert.equal(describedTool!.mutates, true, `"${name}" creates something, so it mutates`);
  }
});

await testAsync('the undo journal survives a restart, and the workbook is what makes it valid', async () => {
  const journalFile = tempJournal();
  // The workbook outlives the app, exactly as a real .xlsx does.
  const excel = fakeExcel();
  const ai = scriptedAi([
    callCompletion('write_range', { address: 'A1', values: [['Hello', 1]] }),
    textCompletion('Written.'),
  ]);

  const first = await startApp({ ai, journalFile });
  const tp1 = startFakeTaskpane(first.base, excel);
  try {
    await waitForExecutor(first.base);
    await chatEvents(first.base, 'write something');
    assert.equal((await status(first.base)).undoAvailable, 1, 'the write should be undoable');
  } finally {
    await tp1.stop();
    await stopServer(first.server);
  }

  const onDisk = JSON.parse(readFileSync(journalFile, 'utf8')) as {
    sessions: Record<string, { entries: Record<string, unknown>[] }[]>;
  };
  assert.equal(onDisk.sessions[SESSION]?.length, 1, 'the task must be written to the journal file, not just memory');
  assert.ok(
    (onDisk.sessions[SESSION]?.[0]?.entries[0] as { expect?: string })?.expect,
    'the entry needs a fingerprint, or undo will refuse it later'
  );
  assert.deepEqual(excel.state.sheets.Sheet1, [['Hello', 1]], 'the workbook keeps the write');

  // A brand-new server: nothing in memory, so anything it knows came off disk.
  const second = await startApp({ ai, journalFile });
  const tp2 = startFakeTaskpane(second.base, excel);
  try {
    await waitForExecutor(second.base);
    assert.equal((await status(second.base)).undoAvailable, 1, 'a fresh server must find the journal on disk');

    const undo = await post(second.base, '/api/undo', { sessionId: SESSION });
    assert.equal(undo.body.undone, 1, `undo failed: ${JSON.stringify(undo.body)}`);
    assert.deepEqual(
      excel.state.sheets.Sheet1,
      [['name', 'qty'], ['widget', 4]],
      'undo across a restart must put the original back'
    );
  } finally {
    await tp2.stop();
    await stopServer(second.server);
  }
});

await testAsync('a persisted undo is refused when the cells have changed since', async () => {
  const journalFile = tempJournal();
  const excel = fakeExcel();
  const ai = scriptedAi([
    callCompletion('write_range', { address: 'A1', values: [['Hello', 1]] }),
    textCompletion('Written.'),
  ]);

  const first = await startApp({ ai, journalFile });
  const tp1 = startFakeTaskpane(first.base, excel);
  try {
    await waitForExecutor(first.base);
    await chatEvents(first.base, 'write something');
  } finally {
    await tp1.stop();
    await stopServer(first.server);
  }

  // Someone edited the workbook while the app was closed. Restoring the old
  // snapshot now would destroy that, so the undo has to refuse instead.
  excel.state.sheets.Sheet1 = [['edited by hand', 99]];

  const second = await startApp({ ai, journalFile });
  const tp2 = startFakeTaskpane(second.base, excel);
  try {
    await waitForExecutor(second.base);
    const undo = await post(second.base, '/api/undo', { sessionId: SESSION });

    assert.equal(undo.body.undone, 0, 'nothing may be restored over newer work');
    assert.deepEqual(excel.state.sheets.Sheet1, [['edited by hand', 99]], 'the newer edit must survive');
    assert.match(String(undo.body.failed?.join(' ')), /changed since this task ran/);
  } finally {
    await tp2.stop();
    await stopServer(second.server);
  }
});

await testAsync('a corrupt journal file does not stop the app starting', async () => {
  const journalFile = tempJournal();
  // A crash mid-write is exactly what leaves this behind.
  writeFileSync(journalFile, '{"sessions": {"test": [', 'utf8');

  const ai = scriptedAi([callCompletion('write_range', { address: 'A1', values: [['x']] }), textCompletion('Done.')]);
  const excel = fakeExcel();
  const { server, base } = await startApp({ ai, journalFile });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    assert.equal((await status(base)).undoAvailable, 0, 'an unreadable journal reads as nothing to undo');
    // And the app still works, which is the point.
    await chatEvents(base, 'write something');
    assert.equal((await status(base)).undoAvailable, 1, 'a new entry must replace the unreadable file');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('a change with no replayable undo plan is refused rather than written', async () => {
  const ai = scriptedAi([callCompletion('write_range', { address: 'A1', values: [['x']] }), textCompletion('Done.')]);
  const excel = fakeExcel();
  const realExecute = excel.execute.bind(excel);
  const before = JSON.stringify(excel.state.sheets.Sheet1);

  /*
   * A taskpane from another build that answers __snapshot, but with nothing this
   * server can hand back to __restore later. Refusing is the same rule as a
   * snapshot that fails outright: an unundoable change is worse than none.
   */
  excel.execute = async (call) => (call.name === '__snapshot' ? { label: 'mystery' } : realExecute(call));

  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, excel);
  try {
    await waitForExecutor(base);
    await chatEvents(base, 'write something');

    assert.ok(!excel.state.seen.includes('write_range'), 'the write must never be dispatched');
    assert.equal(JSON.stringify(excel.state.sheets.Sheet1), before, 'the sheet must be untouched');
    assert.equal((await status(base)).undoAvailable, 0, 'nothing was written, so there is nothing to undo');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('the same-origin probe round-trips a request body', async () => {
  const { server, base } = await startApp({ ai: scriptedAi([textCompletion('x')]) });
  try {
    const res = await fetch(`${base}/api/ping`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ probe: 'value' }),
    });
    const body = (await res.json()) as { echoed?: unknown };
    assert.deepEqual(
      body.echoed,
      { probe: 'value' },
      'the taskpane diagnostics check asserts on this echo; losing it made that check report "0 keys" while still passing'
    );
  } finally {
    await stopServer(server);
  }
});

await testAsync('the workbook context never accumulates in the conversation', async () => {
  const ai = scriptedAi([textCompletion('First answer.'), textCompletion('Second answer.')]);
  const { server, base } = await startApp({ ai });
  const taskpane = startFakeTaskpane(base, fakeExcel());
  try {
    await waitForExecutor(base);
    await chatEvents(base, 'what is in this workbook?');
    await chatEvents(base, 'and now?');

    // The context is per-turn state. If it were stored, the second request would
    // carry two snapshots and every turn after that one more.
    const lastCall = ai.seen[ai.seen.length - 1]!;
    const occurrences = lastCall.contents.join('\n').split('--- current workbook state ---').length - 1;
    assert.equal(occurrences, 1, `expected exactly one context block, found ${occurrences}`);

    // And the user's own words must survive intact.
    assert.ok(
      lastCall.contents.some((c) => c === 'and now?'),
      'the stored user message must be the words alone, not words plus a snapshot'
    );
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

test('the manifest keeps the things that silently break the add-in', () => {
  /*
   * `npm run validate:manifest` runs the official schema validator, but it needs
   * the network and a tool download, so it cannot be part of an offline suite.
   * These are the specific failures it found in this manifest — Excel loaded it
   * happily anyway, which is exactly why they need a test rather than a habit.
   */
  const template = readFileSync('app/web/manifest.template.xml', 'utf8');
  const rendered = renderManifest(template, 'https://localhost:3443');

  // Comments are stripped first: this file's own header mentions the element
  // names in prose, and a check that reads prose instead of markup passes for the
  // wrong reason — which it did, on the first run of this test.
  const markup = rendered.replace(/<!--[\s\S]*?-->/g, '');

  // 1. <Requirements> must sit between <Hosts> and <DefaultSettings>. After
  //    <Permissions> the schema rejects the manifest and Excel quietly ignores it.
  const order = ['<Hosts>', '</Hosts>', '<Requirements>', '</Requirements>', '<DefaultSettings>', '<Permissions>'];
  const positions = order.map((tag) => markup.indexOf(tag));
  assert.ok(
    positions.every((p) => p !== -1),
    `a required element is missing: ${order.filter((t, i) => positions[i] === -1).join(', ')}`
  );
  for (let i = 1; i < positions.length; i++) {
    assert.ok(positions[i] > positions[i - 1], `<${order[i - 1]}> must come before <${order[i]}>`);
  }

  // 2. The schema's floor is 1.0.0.0. This manifest shipped at 0.2.0.0, below it.
  const version = readManifestVersion(markup) as string;
  const [major = 0, minor = 0, patch = 0, build = 0] = version.split('.').map(Number);
  assert.ok(
    [major, minor, patch, build].every((n) => Number.isInteger(n)) && major >= 1,
    `manifest version ${version} is below the 1.0 minimum the schema requires`
  );

  /*
   * 3. No requirement set the client may not honour. A VersionOverrides block
   *    declaring a shared runtime used to live here; it bought less than it cost
   *    and a manifest asking for something the client cannot do is a real way to
   *    stop the add-in loading.
   */
  assert.ok(!markup.includes('<VersionOverrides'), 'the shared-runtime override was removed deliberately');
  assert.ok(!markup.includes('SharedRuntime'), 'do not require SharedRuntime again without re-testing in Excel');

  // And the endpoint everything depends on is still the taskpane.
  assert.ok(markup.includes('/taskpane.html"'), 'the manifest must still load taskpane.html');
});

/* --------------------------------------------------------- certs and setup */

test('a certificate for localhost is generated, validated, and then reused', () => {
  if (!resolveOpenssl()) throw new Error('openssl is not available, so the app cannot generate a certificate');

  const dir = join(mkdtempSync(join(tmpdir(), 'ai-excel-certs-')), 'certs');
  assert.equal(describeCertificate(dir).ready, false, 'an empty directory has no certificate');

  // trust:false is not optional here. The suite runs on the user's machine, and
  // this is the one step that would leave a trusted certificate behind —
  // unlike the journal and the manifest, that is not a temp file.
  const first = ensureCertificate(dir, { trust: false });
  assert.equal(first.action, 'generated');
  assert.equal(first.ready, true, `generation failed: ${first.detail}`);

  for (const [file, mode] of [
    ['localhost.key', '600'],
    ['localhost.crt', '644'],
  ] as const) {
    const actual = (statSync(join(dir, file)).mode & 0o777).toString(8);
    assert.equal(actual, mode, `${file} should be ${mode}, is ${actual}`);
  }

  // One self-signed certificate, no CA: the private key of a CA can vouch for any
  // hostname, and nothing here needs that power.
  for (const gone of ['ca.crt', 'ca.key']) {
    assert.equal(existsSync(join(dir, gone)), false, `${gone} is from the retired CA design`);
  }

  const again = ensureCertificate(dir, { trust: false });
  assert.equal(again.action, 'reused', 'a good certificate must not be regenerated');
  assert.equal(again.notAfter, first.notAfter);
});

test('a certificate that does not cover localhost is refused rather than served', () => {
  if (!resolveOpenssl()) throw new Error('openssl is not available');

  const dir = join(mkdtempSync(join(tmpdir(), 'ai-excel-certs-')), 'certs');
  ensureCertificate(dir, { trust: false });

  // Swap in a certificate for another name: exactly what a stale file or a
  // copy-paste from another project would look like.
  writeFileSync(join(dir, 'openssl.cnf'), readFileSync(join(dir, 'openssl.cnf'), 'utf8').replace('DNS:localhost', 'DNS:example.com'));
  execFileSync(resolveOpenssl() as string, ['req', '-x509', '-new', '-key', 'localhost.key', '-sha256', '-days', '3650',
    '-subj', '/CN=example.com', '-config', 'openssl.cnf', '-extensions', 'leaf_ext', '-out', 'localhost.crt'], { cwd: dir });

  const state = describeCertificate(dir);
  assert.equal(state.ready, false);
  assert.match(state.detail, /does not cover localhost/);
});

test('sideloading reports itself unsupported off macOS instead of guessing', () => {
  // An empty dir means "no folder on this platform" — the app must say so
  // rather than claim success.
  const state = describeManifest('<OfficeApp/>', '');
  assert.equal(state.installed, false);
  assert.equal(state.path, null);
  assert.match(state.detail, /only automated on macOS/);
});

await testAsync('the setup endpoint reports state, and cannot change anything', async () => {
  /*
   * Read-only is the point now. The app writes the certificate and the manifest
   * before this server exists, which is what stops a stub-model test server —
   * `--smoke` — from being able to touch the user's machine at all. There is
   * nothing to POST to.
   */
  const root = mkdtempSync(join(tmpdir(), 'ai-excel-setup-'));
  const dir = join(root, 'wef');
  const certDir = join(root, 'certs');
  const rendered = renderManifest(readFileSync('app/web/manifest.template.xml', 'utf8'), 'https://localhost:3443');

  const withSetup = await startApp({
    ai: scriptedAi([textCompletion('x')]),
    setupped: { manifestTemplate: readFileSync('app/web/manifest.template.xml', 'utf8'), manifestUrl: 'https://localhost:3443', tlsSource: 'test' },
  });
  try {
    const before = await (await fetch(`${withSetup.base}/api/setup`, { cache: 'no-store' })).json();
    assert.equal(before.available, true);
    assert.equal(before.state.manifest.blocked, false);
    assert.equal(before.state.tlsSource, 'test');
    // It reports the REAL machine's manifest, because that is what the panel is
    // for — and it never writes to it.
    assert.equal(typeof before.state.manifest.installed, 'boolean');

    const refused = await post(withSetup.base, '/api/setup', {});
    assert.equal(refused.status, 405, 'there is nothing to run: the app already did it');
  } finally {
    await stopServer(withSetup.server);
  }

  // A server with no setup information — a test, or a stub — says so.
  const without = await startApp({ ai: scriptedAi([textCompletion('x')]) });
  try {
    const res = await (await fetch(`${without.base}/api/setup`, { cache: 'no-store' })).json();
    assert.equal(res.available, false);
    assert.equal(res.state, null);
  } finally {
    await stopServer(without.server);
  }

  // The pieces the panel leans on still behave with no wef folder and no certs.
  assert.equal(describeManifest(rendered, dir).installed, false);
  assert.equal(describeCertificate(certDir).ready, false);
});

await testAsync('the server listens on loopback only, on both families', async () => {
  /*
   * This used to bind every interface. That is not a theoretical problem: the
   * app promises nothing leaves the machine, the certificate exception is
   * loopback-scoped, and the credentials endpoint trusts its origin — all of
   * which a LAN peer could otherwise reach.
   */
  const started = await startServer({
    port: 0,
    ai: scriptedAi([textCompletion('x')]),
    search: null,
  });
  try {
    assert.deepEqual(
      [...started.addresses].sort(),
      ['127.0.0.1', '::1'],
      `expected both loopback families, bound ${started.addresses.join(', ')}`
    );
    assert.ok(started.port > 0, 'port 0 must be resolved so both listeners agree');

    for (const address of ['127.0.0.1', '::1']) {
      const host = address.includes(':') ? `[${address}]` : address;
      const code = await rawStatus(`${started.scheme}://${host}:${started.port}/api/status`);
      assert.equal(code, 200, `${address} must answer`);
    }
  } finally {
    started.server.close();
  }
});

test('a certificate in the keychain is not the same as a trusted one', () => {
  if (!resolveOpenssl()) throw new Error('openssl is not available');

  const dir = join(mkdtempSync(join(tmpdir(), 'ai-excel-trust-')), 'certs');
  // Generated, valid, usable — and untrusted, because nothing asked macOS to
  // trust it. This is exactly what shipped: the app served it, and Excel refused
  // the page with "the content is blocked because it isn't signed by a valid
  // security certificate", which blames neither the certificate nor the trust
  // setting. If this ever returns true for a fresh pair, the check has stopped
  // meaning anything.
  ensureCertificate(dir, { trust: false });
  assert.equal(isTrusted(join(dir, 'localhost.crt')), false, 'a fresh certificate must not verify as trusted');
  assert.equal(isTrusted(join(dir, 'does-not-exist.crt')), false);
});

test('the setup state says so when the certificate is not trusted', () => {
  const dir = join(mkdtempSync(join(tmpdir(), 'ai-excel-setup-')), 'wef');
  const certificate = { dir: '/tmp/certs', ready: true, detail: 'valid until 2036', notAfter: null, trusted: false };
  const state = describeSetup('<OfficeApp/>', certificate, dir);

  assert.equal(state.certificate.trusted, false);
  assert.ok(
    state.remaining.some((step) => /not trusted by macOS/.test(step)),
    `an untrusted certificate must be reported: ${JSON.stringify(state.remaining)}`
  );

  // ...and stays quiet once it is trusted and the manifest is current.
  const content = '<OfficeApp/>';
  installManifest(content, dir);
  const good = describeSetup(content, { ...certificate, trusted: true }, dir);
  assert.deepEqual(good.remaining, []);
});

test('a folder the app cannot open is reported as blocked, not as a missing add-in', () => {
  const dir = join(mkdtempSync(join(tmpdir(), 'ai-excel-blocked-')), 'wef');
  mkdirSync(dir, { recursive: true });
  writeFileSync(join(dir, 'manifest.xml'), '<OfficeApp/>');
  chmodSync(join(dir, 'manifest.xml'), 0o000);

  try {
    const state = describeManifest('<OfficeApp/>', dir);
    // The bug this replaces: an unreadable file was reported as "the Excel add-in
    // could not be installed", which sends you looking at the add-in instead of at
    // whatever is blocking the read.
    assert.equal(state.blocked, true, 'a permission failure is not an absent manifest');
    assert.equal(state.installed, false);
    assert.match(state.detail, /could not be opened/);
    assert.match(state.detail, /sandbox/, 'the message has to name the likely cause');
  } finally {
    chmodSync(join(dir, 'manifest.xml'), 0o600);
  }
});

/* -------------------------------------------------------------- credentials */

await testAsync('a saved API key is never sent back to the window', async () => {
  const credentials = fakeCredentials();
  const ai = scriptedAi([textCompletion('hello')]);
  const { server, base } = await startApp({ ai, credentials });
  const secret = PLACEHOLDER_KEY;

  try {
    const before = await (await fetch(`${base}/api/config`, { cache: 'no-store' })).json();
    assert.equal(before.configurable, true);
    assert.equal(before.config.deepseek.masked, 'sk-test…key1');

    const saved = await post(base, '/api/config', { deepseekApiKey: secret });
    assert.equal(saved.status, 200);

    /*
     * The rule this test exists for: the raw key lives in the config file and in
     * the model client, and nowhere else. If it ever comes back over the wire it
     * is in the renderer's memory, in devtools, and in every screenshot.
     */
    const whole = JSON.stringify(saved.body) + JSON.stringify((await (await fetch(`${base}/api/config`)).json()));
    assert.ok(!whole.includes(secret), `the key came back to the window: ${whole.slice(0, 200)}`);
    assert.ok(!whole.includes(secret.slice(-12)), 'nor may a recognisable tail of it come back');
    assert.equal(saved.body.config.deepseek.masked, 'sk-new…key2');
    assert.equal(credentials.applied.length, 1);
  } finally {
    await stopServer(server);
  }
});

await testAsync('saving a key takes effect on the next turn, with no restart', async () => {
  const credentials = fakeCredentials();
  // The first client refuses, exactly as an unconfigured server does.
  const unconfigured = scriptedAi([textCompletion('should never be reached')]);
  const { server, base } = await startApp({ ai: unconfigured, credentials });
  const taskpane = startFakeTaskpane(base, fakeExcel());

  try {
    await waitForExecutor(base);
    const saved = await post(base, '/api/config', { deepseekApiKey: PLACEHOLDER_KEY });
    assert.equal(saved.status, 200);

    // The status line must show the new model, or the window looks unchanged.
    const s = await status(base);
    assert.equal(s.model, 'replacement-model');
    assert.equal(s.configurable, true);

    // And the swap has to be the client the tool loop actually uses.
    const events = await chatEvents(base, 'hello');
    assert.equal(replyOf(events), 'replaced', 'the turn must run on the client built from the new key');
    assert.equal(unconfigured.seen.length, 0, 'the old client must not be used again');
  } finally {
    await taskpane.stop();
    await stopServer(server);
  }
});

await testAsync('a malformed key is refused before anything is written', async () => {
  const credentials = fakeCredentials();
  const { server, base } = await startApp({ ai: scriptedAi([textCompletion('x')]), credentials });

  try {
    // The mistake the CLI was built to catch, now caught in the window too.
    const pasted = await post(base, '/api/config', { deepseekApiKey: 'pbpaste | node app/setkey.js' });
    assert.equal(pasted.status, 400);
    assert.match(pasted.body.error, /whitespace/);

    const wrongShape = await post(base, '/api/config', { deepseekApiKey: 'not-a-deepseek-key' });
    assert.equal(wrongShape.status, 400);
    assert.match(wrongShape.body.error, /sk-/);

    const badProvider = await post(base, '/api/config', { searchApiKey: 'x'.repeat(40), searchProvider: 'bing' });
    assert.equal(badProvider.status, 400);
    assert.match(badProvider.body.error, /unknown search provider/);

    assert.equal((await post(base, '/api/config', {})).status, 400, 'an empty patch is not a change');
    assert.deepEqual(credentials.applied, [], 'nothing may be applied when validation fails');

    // The shape of the body is untrusted too.
    assert.equal((await post(base, '/api/config', { deepseekApiKey: 42 })).status, 400);
  } finally {
    await stopServer(server);
  }
});

await testAsync('a cross-origin credential change is refused, and the keys are gone when unmanageable', async () => {
  const credentials = fakeCredentials();
  const { server, base } = await startApp({ ai: scriptedAi([textCompletion('x')]), credentials });

  try {
    // A page on another site can still POST to loopback; this is the endpoint
    // where that must not matter.
    const res = await fetch(`${base}/api/config`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Origin: 'https://evil.example' },
      body: JSON.stringify({ deepseekApiKey: PLACEHOLDER_KEY }),
    });
    assert.equal(res.status, 403);
    assert.deepEqual(credentials.applied, []);

    // The same request from our own origin is fine.
    const ours = await fetch(`${base}/api/config`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Origin: `http://${new URL(base).host}` },
      body: JSON.stringify({ deepseekApiKey: PLACEHOLDER_KEY }),
    });
    assert.equal(ours.status, 200);
  } finally {
    await stopServer(server);
  }
});

await testAsync('a server handed its model does not manage keys at all', async () => {
  /*
   * This is the shape `--smoke` starts: a stub model, no config file. If this
   * endpoint worked here, opening the settings dialog during a test run would
   * build a real, keyed client from the user's own config.
   */
  const { server, base } = await startApp({ ai: scriptedAi([textCompletion('x')]) });
  try {
    const res = await (await fetch(`${base}/api/config`, { cache: 'no-store' })).json();
    assert.equal(res.configurable, false);
    assert.equal(res.config, null);

    const refused = await post(base, '/api/config', { deepseekApiKey: PLACEHOLDER_KEY });
    assert.equal(refused.status, 503);

    assert.equal((await status(base)).configurable, false);
  } finally {
    await stopServer(server);
  }
});

await testAsync('an environment variable is reported as the value in charge', async () => {
  const credentials = fakeCredentials({
    deepseek: { set: true, masked: 'sk-env…key9', source: 'env', envVar: 'DEEPSEEK_API_KEY', model: null },
  });
  const { server, base } = await startApp({ ai: scriptedAi([textCompletion('x')]), credentials });
  try {
    const res = await (await fetch(`${base}/api/config`, { cache: 'no-store' })).json();
    // Otherwise saving a key here looks like it did nothing.
    assert.equal(res.config.deepseek.source, 'env');
    assert.equal(res.config.deepseek.envVar, 'DEEPSEEK_API_KEY');
  } finally {
    await stopServer(server);
  }
});

await testAsync('bad requests are rejected with 400', async () => {
  const { server, base } = await startApp({ ai: scriptedAi([textCompletion('x')]) });
  try {
    assert.equal((await post(base, '/api/chat', { message: 'no session' })).status, 400);
    assert.equal((await post(base, '/api/chat', { sessionId: 'x' })).status, 400);
    const badJson = await fetch(`${base}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: '{not json',
    });
    assert.equal(badJson.status, 400);
    const s = await status(base);
    assert.ok(s.tools.length >= 7);
    assert.equal(s.search, null);
  } finally {
    await stopServer(server);
  }
});

/* ------------------------------------------------------------------- report */

console.log(`\n${passed} passed, ${failed} failed\n`);
process.exit(failed ? 1 : 0);
