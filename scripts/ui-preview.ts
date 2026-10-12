/*
 * UI preview server.
 *
 * Runs the REAL server with a scripted model and a fake Excel, so the desktop
 * window can be developed and reviewed without Excel open, a DeepSeek key, or a
 * search key. It seeds two things and then stays up:
 *
 *   1. a completed turn  — user message, tool lines, streamed assistant reply
 *   2. a turn parked on the review gate — an approval card waiting for a decision
 *
 *   npm run preview
 *   npm run desktop -- --smoke          # attaches to the running server
 *
 * Node does NOT read the macOS keychain, so its own calls to this HTTPS server
 * need the dev CA supplied explicitly — hence the npm script, which sets:
 *   NODE_EXTRA_CA_CERTS=~/.office-addin-dev-certs/ca.crt
 *
 * Nothing here is used by the app itself.
 */

import { createServer as createSecureServer } from 'node:https';
import { readFile } from 'node:fs/promises';
import { homedir, tmpdir } from 'node:os';
import { join } from 'node:path';

import { createApp, type ChatEvent } from '../app/server.js';
import type { AiClient, CompleteRequest, Completion, WireUsage } from '../app/ai.js';
import type { ConfigState, CredentialPatch } from '../app/config.js';

/*
 * NOT 3443. The user's Excel taskpane polls that port, so a test server there
 * would be reachable by their real workbook. Keep tests on their own port.
 */
const PORT = Number(process.env.PORT || 3499);

/*
 * NOT the real journal either, and for the same reason: this harness invents
 * changes that never happened, and writing them to ~/.ai-excel/journal.json
 * would leave the real app offering to undo them. Found by crashing into the
 * next problem — the sandbox refused the write and logged EPERM.
 */
const JOURNAL_FILE = process.env.PREVIEW_JOURNAL || join(tmpdir(), 'ai-excel-preview-journal.json');
const SESSION = 'default';
const CRT_PATH = process.env.TLS_CERT || join(homedir(), '.office-addin-dev-certs', 'localhost.crt');
const KEY_PATH = process.env.TLS_KEY || join(homedir(), '.office-addin-dev-certs', 'localhost.key');

const CAPABILITIES = [
  'get_workbook_overview',
  'read_range',
  'get_selection',
  'write_range',
  'set_formula',
  'format_range',
  'sort_range',
  'create_table',
  '__workbook_context',
  '__snapshot',
  '__fingerprint',
  '__restore',
];

/* ------------------------------------------------------------ scripted model */

const USAGE = (prompt: number, completion: number, cacheHit: number): WireUsage => ({
  prompt_tokens: prompt,
  completion_tokens: completion,
  total_tokens: prompt + completion,
  prompt_cache_hit_tokens: cacheHit,
  prompt_cache_miss_tokens: prompt - cacheHit,
});

function textCompletion(content: string): Completion {
  return {
    text: content,
    toolCalls: [],
    finishReason: 'stop',
    usage: USAGE(4820, 96, 4096),
    assistantMessage: { role: 'assistant', content },
  };
}

function callCompletion(name: string, args: Record<string, unknown>): Completion {
  const id = `call_${name}_${Math.random().toString(36).slice(2, 7)}`;
  const rawArgs = JSON.stringify(args);
  return {
    text: '',
    toolCalls: [{ id, name, args, argError: null, rawArgs }],
    finishReason: 'tool_calls',
    usage: USAGE(2410, 58, 2048),
    assistantMessage: {
      role: 'assistant',
      content: null,
      tool_calls: [{ id, type: 'function', function: { name, arguments: rawArgs } }],
    },
  };
}

const SCRIPT: Completion[] = [
  callCompletion('get_workbook_overview', {}),
  callCompletion('write_range', { address: 'A1', values: [['Region', 'Units', 'Price']] }),
  callCompletion('format_range', { address: 'A1:C1', bold: true, fillColor: '#1F6FEB', fontColor: '#FFFFFF' }),
  textCompletion('Added a header row in Sheet1!A1:C1 and formatted it in bold with a blue fill.'),
  // Turn 2 stops here, parked on the review gate.
  callCompletion('write_range', { address: 'E1', values: [['Summary', null], ['Total units', 1240]] }),
  // A search, so a window-initiated turn exercises the search event path too.
  callCompletion('web_search', { query: 'top selling EVs 2026' }),
  // A prose reply, so a window-initiated turn has something to stream.
  // Structured markdown, so the renderer is exercised rather than assumed.
  textCompletion(
    [
      "Here's what **Sheet1** holds, per the last `read_range` call.",
      // Single newlines: these must render as separate lines, not one paragraph.
      '**Current total:** 1240 units',
      '**Largest region:** North at 420',
      '',
      '- **Region** — North, South, East, West',
      '- **Units** — 420, 380, 240, 200',
      '',
      '| Region | Units | Price |',
      '| --- | ---: | ---: |',
      '| North | 420 | 2.50 |',
      '| South | 380 | 3.75 |',
      '',
      'Total units `1240`. Want me to write this to another sheet?',
    ].join('\n')
  ),
];

let step = 0;
const ai: AiClient = {
  model: 'preview-model',
  async complete({ onDelta }: CompleteRequest): Promise<Completion> {
    const next = SCRIPT[Math.min(step, SCRIPT.length - 1)] as Completion;
    step += 1;

    /*
     * A real model takes time and streams. Without either, the window shows no
     * streaming at all and the usage panel reports a meaningless tokens/sec —
     * neither of which is what the user will see.
     */
    const completion = next.usage?.completion_tokens ?? 0;
    const perChunk = Math.max(16, Math.round(completion / 4));

    if (next.text && onDelta) {
      const words = next.text.split(' ');
      for (const word of words) {
        onDelta(`${word} `);
        await new Promise((r) => setTimeout(r, Math.max(12, perChunk)));
      }
    } else {
      await new Promise((r) => setTimeout(r, Math.max(200, completion * 18)));
    }

    return next;
  },
};

/* --------------------------------------------------------------- fake Excel */

const sheet: (string | number)[][] = [
  ['Region', 'Units', 'Price'],
  ['North', 420, 2.5],
  ['South', 380, 3.75],
  ['East', 240, 4.1],
  ['West', 200, 2.25],
];

async function execute(operation: { name: string; args: Record<string, unknown> }): Promise<unknown> {
  const args = operation.args as { address?: string; values?: unknown[][]; tool?: string; args?: Record<string, unknown> };

  if (operation.name === '__workbook_context') {
    return {
      activeSheet: 'Sheet1',
      sheets: [{ name: 'Sheet1', usedRange: 'Sheet1!A1:C5', rowCount: 5, columnCount: 3 }],
      selection: { address: 'Sheet1!A1', values: [['Region', 'Units', 'Price']] },
    };
  }
  if (operation.name === '__snapshot') {
    // Real captureRange returns only the target range's current contents, so
    // the fake must too — returning the whole sheet made the card look wrong.
    const after = args.args?.values as unknown[][] | undefined;
    const rows = Array.isArray(after) ? after.length : 1;
    const cols = Array.isArray(after) && Array.isArray(after[0]) ? (after[0] as unknown[]).length : 1;
    const empty = Array.from({ length: rows }, () => Array.from({ length: cols }, () => ''));
    // The shape the real taskpane answers with: an undo plan plus what the card
    // should show. `label` is the target, which is an address here and would be
    // a table name for a created object.
    return {
      action: { kind: 'range', sheet: null, address: args.args?.address, state: { formulas: sheet.map((r) => [...r]) } },
      label: `Sheet1!${args.args?.address}`,
      preview: empty,
      seal: true,
    };
  }
  if (operation.name === '__fingerprint') {
    return { expect: `fp:${JSON.stringify(sheet)}` };
  }
  if (operation.name === '__restore') {
    const record = operation.args as { action?: { address?: string; state?: { formulas: (string | number)[][] } }; expect?: string };
    const action = record.action;
    if (!action || !record.expect) throw new Error('recorded without a fingerprint');
    if (record.expect !== `fp:${JSON.stringify(sheet)}`) throw new Error(`${action.address} has changed since this task ran`);
    sheet.length = 0;
    sheet.push(...(action.state?.formulas ?? []));
    return { restored: action.address };
  }
  if (operation.name === 'get_workbook_overview') {
    return { activeSheet: 'Sheet1', sheets: [{ name: 'Sheet1', usedRange: 'Sheet1!A1:C5' }] };
  }
  if (operation.name === 'write_range') return { wroteTo: `Sheet1!${args.address}` };
  if (operation.name === 'format_range') return { formatted: `Sheet1!${args.address}` };
  if (operation.name === 'read_range') return { address: `Sheet1!${args.address}`, values: sheet };
  return { ok: true };
}

/* ------------------------------------------------------------------ harness */

let running = true;
async function taskpaneLoop(): Promise<void> {
  while (running) {
    try {
      const params = new URLSearchParams({
        sessionId: SESSION,
        capabilities: CAPABILITIES.join(','),
        timeout: '2000',
      });
      const res = await fetch(`https://localhost:${PORT}/api/executor/next?${params}`, { cache: 'no-store' });
      if (res.status === 204) continue;
      if (!res.ok) continue;

      const { call } = (await res.json()) as { call: { id: string; name: string; args: Record<string, unknown> } };

      let content: unknown = null;
      let error: string | null = null;
      try {
        content = await execute(call);
      } catch (err) {
        error = (err as Error).message;
      }
      await fetch(`https://localhost:${PORT}/api/executor/result`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sessionId: SESSION, id: call.id, content, error }),
      });
    } catch {
      await new Promise((r) => setTimeout(r, 500));
    }
  }
}

/** Consumes the NDJSON chat stream and returns the final events. */
async function chat(message: string, review: boolean): Promise<ChatEvent[]> {
  const res = await fetch(`https://localhost:${PORT}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ sessionId: SESSION, message, review }),
  });

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

/* --------------------------------------------------------------------- main */

const search = {
  provider: 'preview',
  async search(query: string) {
    return {
      query,
      provider: 'preview',
      answer: 'A synthesised answer for preview purposes.',
      results: [{ title: 'Example source', url: 'https://example.com', snippet: 'Preview snippet.' }],
    };
  },
};

/*
 * Fake credentials, so the settings dialog can be seen and driven without the
 * real ~/.ai-excel/config.json. Nothing here touches disk, no key is real, and
 * the whole point is that the harness keeps working after someone saves a key
 * in it.
 */
let previewConfig: ConfigState = {
  file: '~/.ai-excel/config.json (preview — nothing is written here)',
  deepseek: { set: true, masked: 'sk-prev…view', source: 'file', envVar: null, model: 'deepseek-chat' },
  search: { set: false, masked: null, source: null, envVar: null, provider: 'tavily' },
};

const mask = (key: string): string => (key.length <= 12 ? `${key.slice(0, 3)}…` : `${key.slice(0, 6)}…${key.slice(-4)}`);

const credentials = {
  read: (): ConfigState => previewConfig,
  apply(patch: CredentialPatch) {
    if (patch.deepseekApiKey !== undefined) {
      previewConfig = {
        ...previewConfig,
        deepseek: {
          ...previewConfig.deepseek,
          set: patch.deepseekApiKey !== null,
          masked: patch.deepseekApiKey === null ? null : mask(patch.deepseekApiKey),
          source: patch.deepseekApiKey === null ? null : 'file',
        },
      };
    }
    if (patch.searchApiKey !== undefined || patch.searchProvider !== undefined) {
      const provider = patch.searchProvider ?? previewConfig.search.provider;
      const key = patch.searchApiKey;
      previewConfig = {
        ...previewConfig,
        search: {
          provider,
          set: key === undefined ? previewConfig.search.set : key !== null,
          masked: key === undefined ? previewConfig.search.masked : key === null ? null : mask(key),
          source: key === undefined ? previewConfig.search.source : key === null ? null : 'file',
          envVar: null,
        },
      };
    }
    // The scripted model and fake search stay as they are: this is a UI harness.
    return { ai, search };
  },
};

const app = createApp({ ai, search, journalFile: JOURNAL_FILE, credentials });
const server = createSecureServer({ cert: await readFile(CRT_PATH), key: await readFile(KEY_PATH) }, (req, res) => {
  void app(req, res).catch(() => {
    if (!res.headersSent) res.writeHead(500);
    res.end();
  });
});

try {
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    // Loopback only, like the real server: a preview reachable from the network
    // is the same mistake as a test server on the live port.
    server.listen(PORT, '127.0.0.1', resolve);
  });
} catch (err) {
  if ((err as NodeJS.ErrnoException).code === 'EADDRINUSE') {
    console.error(`Port ${PORT} is already in use — another preview is still running.`);
    console.error('Stop it first:  pkill -f ui-preview');
  } else {
    console.error(`Could not start the preview server: ${(err as Error).message}`);
  }
  process.exit(1);
}
console.log(`UI preview server on https://localhost:${PORT}`);
console.log(`  journal: ${JOURNAL_FILE} — NOT ~/.ai-excel/journal.json`);

void taskpaneLoop();
await new Promise((r) => setTimeout(r, 400)); // let the fake taskpane register

const first = await chat('Add a header row and format it', false);
const reply = (first.find((e) => e.type === 'done') as { reply?: string } | undefined)?.reply;
console.log(`  seeded turn 1: ${reply}`);

// Fire and forget: this one parks on the review gate, which is the point.
void chat('Add a summary block', true).catch(() => {});
console.log('  seeded turn 2: parked on the review gate');

console.log('');
console.log(`  open https://localhost:${PORT}/ui/ — or run the desktop app and it will attach.`);
