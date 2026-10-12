/*
 * Server — the hub. Zero runtime dependencies, Node built-ins only.
 *
 * ARCHITECTURE: the UI lives outside Excel; the server mediates.
 *
 *   desktop window ──POST /api/chat (NDJSON stream)──▶ server ──▶ DeepSeek
 *                                                       │
 *                        taskpane ──GET /api/executor/next──┘  (long-poll,
 *                        executes via Office.js                same origin)
 *                        ──POST /api/executor/result──────────▶│
 *   desktop window ◀──── {delta}…{tool}…{done} ──────────────────┘
 *
 * The chat response is a stream because replies are streamed from the model:
 * the window prints text as it is written, and sees each tool call as it runs.
 * Everything the model needs — the workbook context AND every tool that touches
 * Excel — is dispatched to whatever taskpane is connected to that session.
 *
 * All of it is same-origin, so there is no CORS, no mixed content and no port
 * juggling.
 *
 * Run:  node app/server.js
 */

import { createHash } from 'node:crypto';
import { createServer, type IncomingMessage, type ServerResponse } from 'node:http';
import { createServer as createSecureServer } from 'node:https';
import { existsSync } from 'node:fs';
import { readFile, stat } from 'node:fs/promises';
import { homedir } from 'node:os';
import { extname, join, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { Server } from 'node:http';

import { createDeepSeekClient, DEFAULT_MODEL, type AiClient, type WireMessage, type WireUsage } from './ai.js';
import { toolsFor, getTool, validateArgs, TOOLS, SERVER_TOOL_NAMES, type ToolDefinition } from './tools.js';
import { createSearchClient, type SearchClient } from './search.js';
import {
  applyCredentialPatch,
  checkCredentialPatch,
  describeConfig,
  resolveConfig,
  type ConfigState,
  type CredentialPatch,
} from './config.js';
import { createJournalStore, MAX_JOURNAL_TASKS, type JournalTask } from './journal.js';
import { estimateCost, PRICING_AS_OF, PRICING_SOURCE } from './pricing.js';
import { certificatePaths, describeCertificate, isTrusted } from './tls.js';
import { describeManifest, isExcelRunning, manifestDir, renderManifest } from './setup.js';

const WEB_DIR = fileURLToPath(new URL('./web/', import.meta.url));
const PORT = Number(process.env.PORT || 3443);
const MAX_BODY_BYTES = 1_000_000;
const MAX_TOOL_ROUNDS = 12; // runaway guard: a turn may not exceed this many Excel round trips
const MAX_SEARCH_ROUNDS = 3; // server-side search calls allowed within a single turn
const MAX_HISTORY = 40;
const MAX_TOOL_RESULT_CHARS = 20_000;
const EXECUTOR_TIMEOUT_MS = 60_000; // how long we wait for Excel to answer one call
const EXECUTOR_POLL_MS = 25_000; // long-poll hold
const EXECUTOR_STALE_MS = 45_000; // no poll for this long => Excel is gone

/** The default: long enough to read a diff and think, short enough to notice. */
const APPROVAL_TIMEOUT_DEFAULT_MS = 120_000;

/**
 * How long an approval waits for an answer.
 *
 * Read per call rather than captured in a constant, so a test can shorten it. A
 * 120-second timeout is not something a test can wait for, and an expiry path
 * that is never exercised is an expiry path that does not work — which is exactly
 * how the message below went out wrong for so long.
 *
 * The override is also usable in anger: a long turn should not die because
 * somebody stepped away from the machine.
 */
function approvalTimeoutMs(): number {
  const raw = Number(process.env.AI_EXCEL_APPROVAL_TIMEOUT_MS);
  return Number.isFinite(raw) && raw > 0 ? raw : APPROVAL_TIMEOUT_DEFAULT_MS;
}

/**
 * Which certificate to serve with.
 *
 * **Trusted ones only, and it has to be checked rather than assumed.** A
 * certificate can sit in the login keychain and still be untrusted, and Office
 * answers that with "the content is blocked because it isn't signed by a valid
 * security certificate" — an error naming neither the certificate nor the app.
 * That is exactly how this shipped broken once.
 *
 * Ours first, then the Office dev certs if they are what this machine trusts.
 */
function resolveTls(): { cert: string; key: string; source: string; trusted: boolean } | null {
  if (process.env.TLS_CERT && process.env.TLS_KEY) {
    if (existsSync(process.env.TLS_CERT) && existsSync(process.env.TLS_KEY)) {
      return {
        cert: process.env.TLS_CERT,
        key: process.env.TLS_KEY,
        source: 'TLS_CERT / TLS_KEY',
        trusted: isTrusted(process.env.TLS_CERT),
      };
    }
    console.error('[tls] TLS_CERT or TLS_KEY does not exist; ignoring them');
  }

  const own = certificatePaths();
  if (existsSync(own.cert) && existsSync(own.key)) {
    const trusted = isTrusted(own.cert);
    if (!trusted) {
      console.error(
        '[tls] the generated certificate is not trusted by macOS, so Excel will refuse to load the add-in.'
      );
    }
    return { cert: own.cert, key: own.key, source: trusted ? 'generated for this app (trusted)' : 'generated for this app (NOT TRUSTED)', trusted };
  }

  const dev = {
    cert: join(homedir(), '.office-addin-dev-certs', 'localhost.crt'),
    key: join(homedir(), '.office-addin-dev-certs', 'localhost.key'),
  };
  if (existsSync(dev.cert) && existsSync(dev.key)) {
    const trusted = isTrusted(dev.cert);
    return { ...dev, source: trusted ? 'office-addin-dev-certs (trusted)' : 'office-addin-dev-certs (not trusted)', trusted };
  }

  return null;
}

const MIME: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.xml': 'application/xml; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

const BASE_PROMPT = `You are an assistant that operates Microsoft Excel on the user's behalf through tools.

Rules:
- Use A1 notation for addresses ("B2:D20"). Ranges are inclusive.
- Read before you write. Never overwrite a range without first reading it, unless the user
  explicitly asked you to replace its contents.
- Values are rows-of-columns: [["Name","Qty"],["Widget",4]].
- Formulas must begin with "=" and use comma separators.
- Prefer a single tool call over several when the task allows.
- Never guess at data you have not read. If the request is ambiguous about which sheet or range
  it means, ask a short clarifying question instead of acting.`;

const SEARCH_AVAILABLE = `
- You have web_search. Use it whenever you need facts from outside the workbook — statistics,
  prices, specs, game data, documentation, current events. NEVER ask the user to paste data you
  could look up yourself; look it up, then write it into the sheet.
- After searching, use the results. If the sources disagree, say so briefly. Do not invent
  numbers that the search did not return.`;

const SEARCH_UNAVAILABLE = `
- You have NO web access in this session. If a task needs information from outside the workbook,
  say so plainly and ask the user to supply it. Do not imply you searched, and do not invent
  outside facts.`;

const CLOSING = `
- When you have finished, reply with one or two sentences describing what you changed, naming
  the exact sheet and range. Do not describe your reasoning.
- If a tool returns an error, fix your arguments and try again rather than giving up.`;

/**
 * The prompt must match the tools actually on offer: telling the model it can
 * search when it can't produces either invented results or a pointless tool call.
 */
export function buildSystemPrompt({ hasSearch = false }: { hasSearch?: boolean } = {}): string {
  return BASE_PROMPT + (hasSearch ? SEARCH_AVAILABLE : SEARCH_UNAVAILABLE) + CLOSING;
}

/* --------------------------------------------------------------------- types */

export class HttpError extends Error {
  status: number;
  constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}

export interface ExecutorCall {
  id: string;
  name: string;
  args: Record<string, unknown>;
}

export interface ToolOutcome {
  content?: unknown;
  error?: string;
}

/**
 * One entry in the undo journal.
 *
 * The server never looks inside `action`. It was written by the taskpane and is
 * replayed by the taskpane, so the union of undo kinds lives in exactly one
 * place (`app/web/executor.ts`) rather than being mirrored on both sides of the
 * wire, where the two copies would drift apart.
 *
 * `expect` is a fingerprint of what the change left behind, taken just after the
 * write. Undo compares it before restoring: if the cells have moved on, someone
 * has edited them since and restoring would destroy that work.
 */
export interface UndoRecord {
  action: Record<string, unknown>;
  label: string;
  expect?: string | null;
}

export interface UsageTotals {
  calls: number;
  searches: number;
  promptTokens: number;
  completionTokens: number;
  cacheHitTokens: number;
  cacheMissTokens: number;
  elapsedMs: number;
  usd: number;
  estimated: boolean;
  pricedAs: string | null;
}

export type UsageSummary = UsageTotals & {
  totalTokens: number;
  cacheHitRate: number;
  tokensPerSecond: number;
  peak: boolean;
  pricingAsOf: string;
  pricingSource: string;
};

interface ApprovalRequest {
  id: string;
  at: number;
  tool: string;
  /**
   * What the change is aimed at, in the words the taskpane used to describe it:
   * a range address for a cell change, a table name for a created object.
   */
  target: string;
  args: Record<string, unknown>;
  before: unknown;
}

interface Session {
  messages: WireMessage[];
  rounds: number;
  queue: ExecutorCall[];
  waiter: (() => void) | null;
  pending: Map<string, (outcome: ToolOutcome) => void>;
  capabilities: string[];
  lastPollAt: number;
  seq: number;
  journal: JournalTask<UndoRecord>[];
  approvals: ApprovalRequest[];
  approvalWaiters: Map<string, (decision: string) => void>;
  usage: UsageTotals;
}

export type ChatEvent =
  | { type: 'status'; text: string }
  | { type: 'assistant_start' }
  | { type: 'delta'; text: string }
  | { type: 'tool'; name: string; ok: boolean; detail?: string }
  | { type: 'search'; query: string }
  | { type: 'approval'; name: string; ok: boolean; detail?: string }
  | { type: 'done'; reply: string; stats: UsageSummary; elapsedMs: number }
  | { type: 'error'; error: string };

type Emit = (event: ChatEvent) => void;

/**
 * Credential management, when this server is allowed to have any.
 *
 * A separate object rather than something `createApp` does itself, so the server
 * never learns where the config lives or which vendors exist — and so a server
 * handed a model by its caller (the smoke test's stub) cannot build a real keyed
 * one.
 */
export interface Credentials {
  /** The masked view the window may see. Never contains a key. */
  read(): ConfigState;
  /** Persist a validated patch and hand back clients rebuilt from it. */
  apply(patch: CredentialPatch): { ai: AiClient; search: SearchClient | null };
}

export interface AppOptions {
  ai: AiClient;
  search?: SearchClient | null;
  /**
   * Where the undo journal lives. Injectable because a test must never write to
   * the real ~/.ai-excel/journal.json — a test that steps on the user's own
   * journal is the same class of mistake as a test server on the live port.
   */
  journalFile?: string;
  /** Omit to make this server refuse to manage keys at all. */
  credentials?: Credentials;
  /**
   * Where things are, for the window to *report* on. Read-only on purpose: the
   * app writes the certificate and the manifest before the server starts, so a
   * server — including one a test started with a stub model — has no way to
   * change anything about this machine. The panel reports; it does not act.
   */
  setupped?: { manifestTemplate: string; manifestUrl: string; tlsSource: string | null };
}

export interface ValidChat {
  sessionId: string;
  message: string;
  reset: boolean;
  review: boolean;
  search: boolean;
}

export type RequestHandler = (req: IncomingMessage, res: ServerResponse) => Promise<void>;

/* ------------------------------------------------------------------ helpers */

/*
 * A build id derived from the TASKPANE's own files — no manual version bumping.
 *
 * Deliberately not the whole web/ directory: the desktop UI is built by Vite and
 * changes constantly during development, and hashing it would make the taskpane
 * reload itself for no reason. Only the files the taskpane runs matter here.
 */
const TASKPANE_FILES = ['taskpane.html', 'taskpane.js', 'executor.js', 'diagnostics.js'];

let buildIdPromise: Promise<string> | null = null;
export function getBuildId(): Promise<string> {
  if (!buildIdPromise) {
    buildIdPromise = (async () => {
      const hash = createHash('sha1');
      for (const file of TASKPANE_FILES) {
        try {
          const s = await stat(join(WEB_DIR, file));
          hash.update(`${file}:${s.size}:${s.mtimeMs}`);
        } catch {
          hash.update(`${file}:missing`);
        }
      }
      return hash.digest('hex').slice(0, 12);
    })().catch(() => 'unknown');
  }
  return buildIdPromise;
}

function resolveStatic(pathname: string): string | null {
  let rel = decodeURIComponent(pathname).replace(/^[/\\]+/, '');
  if (rel === '' || rel.endsWith('/')) rel += 'index.html'; // /ui/ -> /ui/index.html
  if (rel.endsWith('.ts')) return null; // never serve sources
  const full = normalize(join(WEB_DIR, rel));
  return full.startsWith(WEB_DIR) ? full : null;
}

function originFor(req: IncomingMessage): string {
  const host = (req.headers['x-forwarded-host'] as string) || req.headers.host || `localhost:${PORT}`;
  // Asked of the socket rather than of a constant, so this stays right whichever
  // certificate (or none) the server ended up with.
  const socket = req.socket as { encrypted?: boolean };
  const proto = (req.headers['x-forwarded-proto'] as string) || (socket.encrypted ? 'https' : 'http');
  return `${proto}://${host}`;
}

/**
 * Whether a request that claims an origin really came from it. `Origin` is set by
 * the browser and cannot be forged by a page, so comparing its host to the Host
 * we were reached on is enough to tell our own window from a page on another site.
 */
function isSameOrigin(req: IncomingMessage, origin: string): boolean {
  try {
    return new URL(origin).host === req.headers.host;
  } catch {
    return false;
  }
}

/** The credential patch is untrusted JSON: check its shape before trusting a field. */
function validateCredentialBody(body: unknown): CredentialPatch {
  if (!body || typeof body !== 'object' || Array.isArray(body)) throw new HttpError(400, 'expected a JSON object');
  const raw = body as Record<string, unknown>;
  const patch: CredentialPatch = {};

  if ('deepseekApiKey' in raw) {
    if (raw.deepseekApiKey !== null && typeof raw.deepseekApiKey !== 'string') {
      throw new HttpError(400, 'deepseekApiKey must be a string, or null to clear it');
    }
    patch.deepseekApiKey = raw.deepseekApiKey as string | null;
  }
  if ('searchApiKey' in raw) {
    if (raw.searchApiKey !== null && typeof raw.searchApiKey !== 'string') {
      throw new HttpError(400, 'searchApiKey must be a string, or null to clear it');
    }
    patch.searchApiKey = raw.searchApiKey as string | null;
  }
  if ('searchProvider' in raw) {
    if (typeof raw.searchProvider !== 'string') throw new HttpError(400, 'searchProvider must be a string');
    patch.searchProvider = raw.searchProvider;
  }

  return patch;
}

async function readJsonBody(req: IncomingMessage): Promise<Record<string, unknown>> {
  const chunks: Buffer[] = [];
  let total = 0;
  for await (const chunk of req) {
    const buf = chunk as Buffer;
    total += buf.length;
    if (total > MAX_BODY_BYTES) throw new HttpError(413, 'request body too large');
    chunks.push(buf);
  }
  if (!total) return {};
  try {
    return JSON.parse(Buffer.concat(chunks).toString('utf8')) as Record<string, unknown>;
  } catch (err) {
    throw new HttpError(400, `invalid JSON body: ${(err as Error).message}`);
  }
}

function sendJson(res: ServerResponse, status: number, payload: unknown): void {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
    'Content-Length': Buffer.byteLength(body),
  });
  res.end(body);
}

function renderContext(context: unknown): string {
  const ctx = context as
    | {
        sheets?: { name: string; usedRange?: string | null; rowCount?: number; columnCount?: number }[];
        activeSheet?: string;
        selection?: { address: string; values?: unknown };
      }
    | null;
  if (!ctx) return '';

  const lines: string[] = [];
  if (ctx.sheets?.length) {
    lines.push('Workbook sheets:');
    for (const s of ctx.sheets) {
      lines.push(`  - "${s.name}" used range ${s.usedRange || '(empty)'} (${s.rowCount || 0} rows x ${s.columnCount || 0} cols)`);
    }
  }
  if (ctx.activeSheet) lines.push(`Active sheet: "${ctx.activeSheet}"`);
  if (ctx.selection) {
    lines.push(`Current selection: ${ctx.selection.address}`);
    if (ctx.selection.values) {
      lines.push(`Selection values: ${JSON.stringify(ctx.selection.values).slice(0, 2000)}`);
    }
  }
  return lines.join('\n');
}

function clipResult(text: unknown): string {
  const s = typeof text === 'string' ? text : JSON.stringify(text ?? '');
  return s.length > MAX_TOOL_RESULT_CHARS ? `${s.slice(0, MAX_TOOL_RESULT_CHARS)}\n...(truncated)` : s;
}

/* ------------------------------------------------------------ usage tracking */

function freshUsage(): UsageTotals {
  return {
    calls: 0,
    // Searches cost no tokens but consume the provider's own quota, so the
    // usage panel would otherwise show a search-heavy session as free.
    searches: 0,
    promptTokens: 0,
    completionTokens: 0,
    cacheHitTokens: 0,
    cacheMissTokens: 0,
    elapsedMs: 0,
    usd: 0,
    estimated: false,
    pricedAs: null,
  };
}

/**
 * One model call. Cost is priced per call rather than from the running totals,
 * so a session that spans the peak boundary is charged correctly for each half.
 */
function recordUsage(usage: WireUsage | null, elapsedMs: number, model: string): UsageTotals {
  const u = freshUsage();
  u.calls = 1;
  u.elapsedMs = elapsedMs;
  if (!usage) return u;

  const promptTokens = usage.prompt_tokens || 0;
  const cacheHitTokens = usage.prompt_cache_hit_tokens || 0;
  const cacheMissTokens = usage.prompt_cache_miss_tokens ?? Math.max(0, promptTokens - cacheHitTokens);

  u.promptTokens = promptTokens;
  u.completionTokens = usage.completion_tokens || 0;
  u.cacheHitTokens = cacheHitTokens;
  u.cacheMissTokens = cacheMissTokens;

  const cost = estimateCost(
    {
      promptTokens,
      completionTokens: u.completionTokens,
      cacheHitTokens,
      cacheMissTokens,
    },
    model
  );
  u.usd = cost.usd;
  u.estimated = !cost.exact;
  u.pricedAs = cost.model;
  return u;
}

function addUsage(total: UsageTotals, delta: UsageTotals): void {
  total.calls += delta.calls;
  total.searches += delta.searches;
  total.promptTokens += delta.promptTokens;
  total.completionTokens += delta.completionTokens;
  total.cacheHitTokens += delta.cacheHitTokens;
  total.cacheMissTokens += delta.cacheMissTokens;
  total.elapsedMs += delta.elapsedMs;
  total.usd += delta.usd;
  total.estimated = total.estimated || delta.estimated;
  total.pricedAs = delta.pricedAs || total.pricedAs;
}

/** Raw counters plus the derived figures the UI shows. */
function summariseUsage(u: UsageTotals, model: string): UsageSummary {
  const seconds = u.elapsedMs / 1000;
  const cost = estimateCost(
    {
      promptTokens: u.promptTokens,
      completionTokens: u.completionTokens,
      cacheHitTokens: u.cacheHitTokens,
      cacheMissTokens: u.cacheMissTokens,
    },
    model
  );
  return {
    ...u,
    totalTokens: u.promptTokens + u.completionTokens,
    cacheHitRate: u.promptTokens ? u.cacheHitTokens / u.promptTokens : 0,
    // Wall clock, so it includes network latency: a lower bound on real speed.
    tokensPerSecond: seconds > 0 ? u.completionTokens / seconds : 0,
    estimated: u.estimated || !cost.exact,
    peak: cost.peak,
    pricedAs: u.pricedAs || cost.model,
    pricingAsOf: PRICING_AS_OF,
    pricingSource: PRICING_SOURCE,
  };
}

/* ----------------------------------------------------------------------- app */

export function createApp({
  ai: initialAi,
  search: initialSearch = null,
  journalFile,
  credentials,
  setupped,
}: AppOptions): RequestHandler {
  /*
   * `ai` and `search` are reassignable: the window can change a key, and the next
   * turn has to use the new client without restarting the app. Everything below
   * reads them from this closure, so a swap reaches every call site.
   */
  let ai = initialAi;
  let search = initialSearch;

  const sessions = new Map<string, Session>();
  const journal = createJournalStore<UndoRecord>(journalFile);

  function session(id: string): Session {
    let s = sessions.get(id);
    if (!s) {
      s = {
        messages: [],
        rounds: 0,
        queue: [],
        waiter: null,
        pending: new Map(),
        capabilities: [],
        lastPollAt: 0,
        seq: 0,
        // Loaded from disk, so "Undo last task" survives a restart. The journal
        // describes the workbook, not the conversation, which is also why
        // `reset` does not clear it.
        journal: journal.load(id),
        approvals: [],
        approvalWaiters: new Map(),
        usage: freshUsage(),
      };
      sessions.set(id, s);
    }
    return s;
  }

  function excelConnected(s: Session): boolean {
    return Date.now() - s.lastPollAt < EXECUTOR_STALE_MS;
  }

  /** Hand a call to the connected taskpane and wait for its answer. */
  function dispatch(s: Session, call: ExecutorCall, timeoutMs = EXECUTOR_TIMEOUT_MS): Promise<ToolOutcome> {
    if (!excelConnected(s)) {
      return Promise.reject(
        new HttpError(409, 'Excel is not connected. Open the add-in taskpane in Excel (Insert → Add-ins → My Add-ins).')
      );
    }
    return new Promise<ToolOutcome>((resolve) => {
      const timer = setTimeout(() => {
        s.pending.delete(call.id);
        resolve({ error: `Excel did not respond within ${Math.round(timeoutMs / 1000)}s` });
      }, timeoutMs);

      s.pending.set(call.id, (value) => {
        clearTimeout(timer);
        s.pending.delete(call.id);
        resolve(value);
      });

      s.queue.push(call);
      const waiter = s.waiter;
      if (waiter) {
        s.waiter = null;
        waiter();
      }
    });
  }

  /**
   * Ask the UI to approve one change and wait for the answer. The turn is parked
   * here in the meantime — the taskpane must not write until we hear back.
   */
  function requestApproval(s: Session, request: Omit<ApprovalRequest, 'id' | 'at'>): Promise<string> {
    return new Promise<string>((resolve) => {
      const id = `appr-${(s.seq += 1)}`;
      s.approvals.push({ id, at: Date.now(), ...request });

      const settle = (decision: string): void => {
        clearTimeout(timer);
        s.approvalWaiters.delete(id);
        s.approvals = s.approvals.filter((a) => a.id !== id);
        resolve(decision);
      };

      const timer = setTimeout(() => settle('timeout'), approvalTimeoutMs());
      s.approvalWaiters.set(id, settle);
    });
  }

  /**
   * What the model is told when a change did not happen — and it must never say
   * "the user declined" about an approval that merely expired.
   *
   * It did, until this was fixed: an unanswered card produced "The user declined
   * this change. Do not retry it; ask what they would prefer instead." The model
   * then behaved as though the user had actively rejected the change and stopped
   * — for a change the user most likely never saw, because they were away from
   * the machine. A refusal and a silence are different events and the model can
   * act sensibly on both, but only if it is told which one happened.
   */
  function notAppliedMessage(decision: string): string {
    if (decision === 'timeout') {
      return (
        'This change was NOT applied: the approval request expired with no answer. Nobody declined it. ' +
        'Tell the user it expired, and ask whether to try it again.'
      );
    }
    return 'The user declined this change. Do not retry it; ask what they would prefer instead.';
  }

  /** The same distinction, in the words the window shows. */
  function notAppliedDetail(decision: string, label: string): string {
    if (decision === 'timeout') return `expired with no answer, so nothing was written (${label})`;
    return `declined, so nothing was written (${label})`;
  }

  /** Validated before the response starts streaming, so bad input is a real 400. */
  function validateChatBody(body: Record<string, unknown>): ValidChat {
    const { sessionId, message, reset = false, review = false, search = false } = body as {
      sessionId?: string;
      message?: string;
      reset?: boolean;
      review?: boolean;
      search?: boolean;
    };
    if (!sessionId || typeof sessionId !== 'string') throw new HttpError(400, 'sessionId is required');
    if (typeof message !== 'string' || !message.trim()) throw new HttpError(400, '"message" is required');
    return {
      sessionId,
      message,
      reset: Boolean(reset),
      review: Boolean(review),
      // Absent means off, so any client that does not ask for search does not get it.
      search: Boolean(search),
    };
  }

  interface ValidatedCall extends ExecutorCall {
    error?: string;
  }

  function validateCall(toolCall: {
    id: string;
    name: string;
    args: Record<string, unknown>;
    argError: string | null;
  }): ValidatedCall {
    const tool = getTool(toolCall.name);
    if (!tool) return { id: toolCall.id, name: toolCall.name, args: {}, error: `unknown tool "${toolCall.name}"` };
    if (toolCall.argError) return { id: toolCall.id, name: toolCall.name, args: {}, error: toolCall.argError };
    const errors = validateArgs(tool.parameters, toolCall.args);
    return errors.length
      ? { id: toolCall.id, name: toolCall.name, args: toolCall.args, error: `invalid arguments: ${errors.join('; ')}` }
      : { id: toolCall.id, name: toolCall.name, args: toolCall.args };
  }

  async function executeServerTool(call: ValidatedCall): Promise<unknown> {
    if (call.name === 'web_search') {
      if (!search) throw new Error('web search is not configured on this server');
      return search.search(String(call.args.query), call.args.maxResults as number | undefined);
    }
    throw new Error(`no server executor for "${call.name}"`);
  }

  async function handleChat(chat: ValidChat, emit: Emit): Promise<void> {
    const { sessionId, message, reset, review, search: wantsSearch } = chat;

    if (reset) sessions.delete(sessionId);
    const s = session(sessionId);

    s.rounds = 0;

    /*
     * Everything this turn changes is collected here so it can be undone as one
     * unit. A task with no entries (a question, a read) is never journalled, so
     * undo always steps over real changes.
     */
    const task: JournalTask<UndoRecord> = { at: Date.now(), entries: [] };
    const commit = (reply: string): void => {
      if (task.entries.length) {
        s.journal.push(task);
        if (s.journal.length > MAX_JOURNAL_TASKS) s.journal.shift();
        journal.save(sessionId, s.journal);
      }
      s.rounds = 0;
      emit({ type: 'done', reply, stats: summariseUsage(s.usage, ai.model), elapsedMs: Date.now() - task.at });
    };

    // The workbook context is fetched from Excel, not supplied by the UI — that
    // way any UI gets it for free and none of them need Office.js.
    emit({ type: 'status', text: 'Reading the workbook…' });
    let ctxText = '';
    try {
      const ctx = await dispatch(s, { id: `ctx-${s.seq}`, name: '__workbook_context', args: {} }, 15_000);
      if (ctx && ctx.content && !ctx.error) ctxText = renderContext(ctx.content);
    } catch (err) {
      if (err instanceof HttpError) throw err; // Excel is not connected: say so plainly
    }

    /*
     * Only the user's own words go into history.
     *
     * The workbook context used to be concatenated onto this message, so every
     * turn left a full snapshot in the conversation. Ten turns in, the model was
     * reading ten descriptions of the sheet — stale ones included, which is a
     * correctness problem and not just a token cost. It is per-turn state, so it
     * travels with the system prompt instead and is rebuilt from Excel each time.
     */
    s.messages.push({ role: 'user', content: message.trim() });

    if (s.messages.length > MAX_HISTORY) s.messages.splice(0, s.messages.length - MAX_HISTORY);

    // A configured key only makes search *possible*; the toggle decides whether
    // it is offered. The prompt and the tool list must agree, so both read this.
    const searchEnabled = wantsSearch && Boolean(search);

    const offered = toolsFor({
      capabilities: s.capabilities,
      serverTools: searchEnabled ? SERVER_TOOL_NAMES : [],
    });
    let searches = 0;
    let lastCallUsage: WireUsage | null = null;

    const contextBlock = ctxText ? `\n\n--- current workbook state ---\n${ctxText}` : '';

    for (let round = 0; round <= MAX_TOOL_ROUNDS; round++) {
      // One bubble per round that produces prose, so intermediate narration and
      // the final answer do not run together in the window.
      let openedRound = false;
      const startedAt = Date.now();

      // Only the first round gets it: it describes the state before the turn
      // began, and by later rounds the model's own tool results are fresher.
      const systemPrompt = buildSystemPrompt({ hasSearch: searchEnabled }) + (round === 0 ? contextBlock : '');

      const result = await ai.complete({
        messages: [{ role: 'system', content: systemPrompt }, ...s.messages],
        tools: offered,
        onDelta: (text) => {
          if (!openedRound) {
            openedRound = true;
            emit({ type: 'assistant_start' });
          }
          emit({ type: 'delta', text });
        },
      });

      addUsage(s.usage, recordUsage(result.usage, Date.now() - startedAt, ai.model));
      if (result.usage) lastCallUsage = result.usage;
      void lastCallUsage;
      s.messages.push(result.assistantMessage);

      if (!result.toolCalls.length) {
        return commit(result.text || '(the model returned no text)');
      }

      const calls = result.toolCalls.map(validateCall);

      // 1. Server-side tools resolve here, inline.
      const allowedServerTools = searchEnabled ? SERVER_TOOL_NAMES : [];

      for (const call of calls.filter((c) => !c.error && getTool(c.name)?.executesOn === 'server')) {
        let content: unknown;

        if (!allowedServerTools.includes(call.name)) {
          /*
           * The catalog still contains this tool, so a hallucinated call would
           * otherwise execute despite never being offered. Refuse it explicitly.
           */
          content = { error: `"${call.name}" is not enabled right now. Answer without it.` };
        } else if (searches >= MAX_SEARCH_ROUNDS) {
          content = { error: 'Search budget exhausted. Answer using what you already have.' };
        } else {
          emit({ type: 'search', query: String(call.args.query ?? '') });
          searches += 1;
          s.usage.searches += 1;
          try {
            content = await executeServerTool(call);
          } catch (err) {
            content = { error: (err as Error).message };
          }
        }

        s.messages.push({ role: 'tool', tool_call_id: call.id, content: clipResult(content) });
      }

      // 2. Everything else goes to Excel.
      const excelCalls = calls.filter((c) => c.error || getTool(c.name)?.executesOn !== 'server');
      if (!excelCalls.length) continue;

      if (round === MAX_TOOL_ROUNDS) {
        return commit(`I stopped after ${MAX_TOOL_ROUNDS} steps without finishing. Nothing further was changed.`);
      }

      for (const call of excelCalls) {
        const tool = getTool(call.name) as ToolDefinition;

        if (call.error) {
          s.messages.push({ role: 'tool', tool_call_id: call.id, content: call.error });
          emit({ type: 'tool', name: call.name, ok: false, detail: call.error });
          continue;
        }

        let undo: UndoRecord | null = null;
        let needsSeal = false;

        if (tool.mutates) {
          /*
           * Refuse rather than write unundoably. A taskpane loaded before this
           * check existed cannot snapshot, and we must not guess that it can.
           */
          if (!s.capabilities.includes('__snapshot')) {
            const detail =
              'The add-in taskpane is running an older version and cannot snapshot changes, so this was ' +
              'refused rather than applied without a way to undo it. Close and reopen the taskpane in ' +
              'Excel (Insert → Add-ins → My Add-ins), then try again.';
            s.messages.push({ role: 'tool', tool_call_id: call.id, content: detail });
            emit({ type: 'tool', name: call.name, ok: false, detail: 'taskpane out of date — reload the add-in pane' });
            continue;
          }

          // Plan the change BEFORE writing it. If we cannot, we do not write —
          // an unundoable change is worse than a refused one.
          const snap = await dispatch(s, {
            id: `snap-${s.seq}-${call.id}`,
            name: '__snapshot',
            args: { tool: call.name, args: call.args },
          });

          const recorded = snap.content as
            | { action?: Record<string, unknown>; label?: string; preview?: unknown; seal?: boolean }
            | null
            | undefined;

          if (snap.error || !recorded?.action) {
            const why = snap.error || 'the add-in taskpane did not say how to undo it';
            const detail = `could not preview the change (${why}), so it was not applied`;
            s.messages.push({ role: 'tool', tool_call_id: call.id, content: detail });
            emit({ type: 'tool', name: call.name, ok: false, detail });
            continue;
          }

          undo = { action: recorded.action, label: recorded.label || String(call.args.address ?? call.name) };
          needsSeal = Boolean(recorded.seal);

          if (review) {
            const decision = await requestApproval(s, {
              tool: call.name,
              target: undo.label,
              args: call.args,
              before: recorded.preview,
            });
            emit({
              type: 'approval',
              name: call.name,
              ok: decision === 'approve',
              detail: decision === 'approve' ? `approved ${undo.label}` : notAppliedDetail(decision, undo.label),
            });
            if (decision !== 'approve') {
              s.messages.push({
                role: 'tool',
                tool_call_id: call.id,
                content: notAppliedMessage(decision),
              });
              continue;
            }
          }
        }

        const outcome = await dispatch(s, call);

        if (!outcome.error && undo) {
          /*
           * Ask what the cells hold NOW, so a later undo can tell whether
           * anything has touched them in between. Only the kinds that restore
           * cell state need this — removing an object you created cannot
           * overwrite data — and the taskpane says which, so this side still
           * never reads inside `action`.
           *
           * A failure here leaves `expect` null, which makes the undo refuse
           * rather than guess. The write has already happened, so refusing to
           * journal it would only lose the step.
           */
          if (needsSeal) {
            const sealed = await dispatch(s, {
              id: `seal-${s.seq}-${call.id}`,
              name: '__fingerprint',
              args: { action: undo.action },
            });
            undo.expect = sealed.error ? null : ((sealed.content as { expect?: string | null })?.expect ?? null);
          }
          task.entries.push(undo);
        }

        /*
         * Reported after the fact so `ok` reflects what actually happened. This
         * used to be emitted before the dispatch with ok: true, which told the
         * user a call had succeeded when Excel had in fact rejected it.
         */
        {
          const detail = outcome.error ? outcome.error : JSON.stringify(call.args).slice(0, 120);
          emit({ type: 'tool', name: call.name, ok: !outcome.error, detail });
        }

        s.messages.push({
          role: 'tool',
          tool_call_id: call.id,
          content: outcome.error ? `error: ${outcome.error}` : clipResult(outcome.content),
        });
      }
    }

    return commit('I could not finish that request. Try rephrasing it, or break it into smaller steps.');
  }

  return async function handle(req: IncomingMessage, res: ServerResponse): Promise<void> {
    const origin = originFor(req);
    const url = new URL(req.url || '/', `http://${req.headers.host || 'localhost'}`);
    const { pathname } = url;

    // On every response, so a long-lived taskpane can notice it is out of date.
    res.setHeader('X-App-Build', await getBuildId());

    try {
      if (pathname === '/api/ping') {
        // The taskpane's same-origin diagnostics check asserts on `echoed`, so a
        // request body must actually make the round trip. Dropping this field
        // during the TypeScript rewrite made that check report "0 keys" while
        // still passing — proving less than it claimed.
        const body = req.method === 'POST' ? await readJsonBody(req).catch(() => ({})) : {};
        return sendJson(res, 200, { ok: true, origin, pid: process.pid, build: await getBuildId(), echoed: body });
      }

      if (pathname === '/api/version') {
        return sendJson(res, 200, { build: await getBuildId() });
      }

      if (pathname === '/api/status') {
        const id = url.searchParams.get('sessionId');
        const s = id ? sessions.get(id) : undefined;
        return sendJson(res, 200, {
          model: ai.model || DEFAULT_MODEL,
          search: search ? search.provider : null,
          excelConnected: s ? excelConnected(s) : false,
          capabilities: s ? s.capabilities : [],
          // Falls back to disk rather than the in-memory session: on a freshly
          // started server nothing has run yet, and "Undo (1)" has to be true
          // before the first turn or the safety net looks empty when it is not.
          undoAvailable: s ? s.journal.length : id ? journal.load(id).length : 0,
          approvalsPending: s ? s.approvals.length : 0,
          usage: summariseUsage(s ? s.usage : freshUsage(), ai.model),
          configurable: Boolean(credentials),
          tlsSource: setupped?.tlsSource ?? null,
          tools: TOOLS.map((t) => ({ name: t.name, executesOn: t.executesOn || 'taskpane', mutates: !!t.mutates })),
        });
      }

      // --- connecting Excel: reporting only ---------------------------------
      /*
       * Read-only, and deliberately so. The app writes the certificate and the
       * manifest before this server exists, so there is nothing here for a
       * server to do — which is what keeps a stub-model test server from being
       * able to touch the user's machine at all.
       */
      if (pathname === '/api/setup') {
        if (!setupped) return sendJson(res, 200, { available: false, state: null });
        if (req.method !== 'GET') throw new HttpError(405, 'use GET');

        const cert = describeCertificate();
        return sendJson(res, 200, {
          available: true,
          state: {
            certificate: { ...cert, trusted: cert.ready && isTrusted(certificatePaths().cert) },
            manifest: describeManifest(renderManifest(setupped.manifestTemplate, setupped.manifestUrl)),
            manifestDir: manifestDir(),
            excelRunning: isExcelRunning(),
            tlsSource: setupped.tlsSource,
          },
        });
      }

      // --- credentials ------------------------------------------------------
      if (pathname === '/api/config') {
        if (!credentials) {
          if (req.method !== 'GET') {
            throw new HttpError(
              503,
              'This server was started with credentials supplied to it, so it does not manage any. ' +
                'Set keys with:  node app/setkey.js'
            );
          }
          return sendJson(res, 200, { configurable: false, config: null });
        }

        if (req.method === 'GET') {
          return sendJson(res, 200, { configurable: true, config: credentials.read() });
        }
        if (req.method !== 'POST') throw new HttpError(405, 'use GET or POST');

        /*
         * Loopback is reachable from any page in a browser, and this is the one
         * endpoint that changes credentials, so a cross-origin POST is refused.
         * A request with no Origin header is allowed on purpose: the tests and the
         * CLI are not browsers and send none, and a browser always sends one on a
         * cross-origin POST.
         */
        const origin = req.headers.origin;
        if (origin && !isSameOrigin(req, origin)) {
          throw new HttpError(403, 'credential changes are only accepted from this app');
        }

        const patch = validateCredentialBody(await readJsonBody(req));
        const { error, warning } = checkCredentialPatch(patch);
        if (error) throw new HttpError(400, error);

        const applied = credentials.apply(patch);
        ai = applied.ai;
        search = applied.search;

        return sendJson(res, 200, { ok: true, config: credentials.read(), warning: warning ?? null });
      }

      // --- the taskpane's inbound channel ---------------------------------
      if (pathname === '/api/executor/next') {
        const id = url.searchParams.get('sessionId');
        if (!id) throw new HttpError(400, 'sessionId is required');
        const s = session(id);
        s.lastPollAt = Date.now();

        const caps = url.searchParams.get('capabilities');
        s.capabilities = caps ? caps.split(',').filter(Boolean) : [];

        if (s.queue.length) return sendJson(res, 200, { call: s.queue.shift() });

        const timeoutMs = Math.min(Number(url.searchParams.get('timeout') || EXECUTOR_POLL_MS), 30_000);
        const call = await new Promise<ExecutorCall | null>((resolve) => {
          const deliver = (): void => {
            clearTimeout(timer);
            resolve(s.queue.shift() || null);
          };
          const timer = setTimeout(() => {
            if (s.waiter === deliver) s.waiter = null;
            resolve(null);
          }, timeoutMs);
          s.waiter = deliver;
        });

        if (!call) {
          res.writeHead(204, { 'Cache-Control': 'no-store' });
          res.end();
          return;
        }
        return sendJson(res, 200, { call });
      }

      if (pathname === '/api/executor/result') {
        if (req.method !== 'POST') throw new HttpError(405, 'use POST');
        const body = await readJsonBody(req);
        const s = sessions.get(String(body.sessionId));
        const pending = s?.pending.get(String(body.id));
        if (pending) pending({ content: body.content, error: body.error as string | undefined });
        return sendJson(res, 200, { ok: true });
      }

      // --- preview + undo --------------------------------------------------
      if (pathname === '/api/approvals') {
        const id = url.searchParams.get('sessionId');
        const s = id ? sessions.get(id) : undefined;
        return sendJson(res, 200, { approvals: s ? s.approvals : [] });
      }

      if (pathname === '/api/approvals/decision') {
        if (req.method !== 'POST') throw new HttpError(405, 'use POST');
        const body = await readJsonBody(req);
        const s = sessions.get(String(body.sessionId));
        const settle = s?.approvalWaiters.get(String(body.id));
        if (settle) settle(body.decision === 'approve' ? 'approve' : 'decline');
        return sendJson(res, 200, { ok: Boolean(settle) });
      }

      if (pathname === '/api/undo') {
        if (req.method !== 'POST') throw new HttpError(405, 'use POST');
        const body = await readJsonBody(req);
        /*
         * Hydrated from disk, not read out of the map: after a restart the
         * journal exists but no session has been created, and being able to undo
         * something from before the restart is the entire reason it is on disk.
         */
        const s = session(String(body.sessionId));
        const task = s.journal.pop();

        if (!task) return sendJson(res, 200, { undone: 0, remaining: 0, message: 'Nothing to undo.' });

        // Persisted before the attempts, not after: a crash mid-undo must not
        // leave the same task available to be undone twice.
        journal.save(String(body.sessionId), s.journal);

        let restored = 0;
        const failed: string[] = [];
        // Newest first: later writes may overlap earlier ranges.
        for (const entry of [...task.entries].reverse()) {
          /*
           * `action` plus `expect`, not the whole record: `label` is for the
           * user, and the taskpane is what decides whether these cells are still
           * the ones it changed.
           */
          const outcome = await dispatch(s, {
            id: `undo-${(s.seq += 1)}`,
            name: '__restore',
            args: { action: entry.action, expect: entry.expect ?? null },
          });
          if (outcome.error) failed.push(`${entry.label} — ${outcome.error}`);
          else restored += 1;
        }

        return sendJson(res, 200, {
          undone: restored,
          failed,
          remaining: s.journal.length,
          message: failed.length
            ? `Restored ${restored} change(s), but ${failed.length} could not be restored: ${failed.join('; ')}`
            : `Undid the last task — restored ${restored} change(s).`,
        });
      }

      // --- the UI -----------------------------------------------------------
      if (pathname === '/api/chat') {
        if (req.method !== 'POST') throw new HttpError(405, 'use POST');
        const body = await readJsonBody(req);
        // Validated BEFORE the stream opens, so malformed input is a real 400
        // rather than an error event inside a 200.
        const chat = validateChatBody(body);

        /*
         * Streamed NDJSON rather than one JSON body: the reply is streamed from
         * the model, so the window can print it as it is written and show each
         * tool call as it happens.
         */
        res.writeHead(200, {
          'Content-Type': 'application/x-ndjson; charset=utf-8',
          'Cache-Control': 'no-store',
          'X-Accel-Buffering': 'no',
        });

        let closed = false;
        const emit: Emit = (event) => {
          if (closed) return;
          res.write(`${JSON.stringify(event)}\n`);
        };

        req.on('close', () => {
          closed = true;
        });

        try {
          await handleChat(chat, emit);
        } catch (err) {
          if (!(err instanceof HttpError)) console.error('[chat]', err);
          emit({ type: 'error', error: err instanceof HttpError ? err.message : (err as Error).message });
        } finally {
          res.end();
        }
        return;
      }

      if (pathname === '/manifest.xml') {
        const template = await readFile(join(WEB_DIR, 'manifest.template.xml'), 'utf8');
        res.writeHead(200, { 'Content-Type': MIME['.xml'] as string, 'Cache-Control': 'no-store' });
        res.end(template.replaceAll('{{BASE_URL}}', origin));
        return;
      }

      if (pathname === '/') {
        res.writeHead(302, { Location: '/ui/' });
        res.end();
        return;
      }

      // The window used to be a single static page; keep old links working.
      if (pathname === '/ui' || pathname === '/ui.html') {
        res.writeHead(302, { Location: '/ui/' });
        res.end();
        return;
      }

      const file = resolveStatic(pathname);
      if (!file) throw new HttpError(403, 'forbidden');
      try {
        const data = await readFile(file);
        res.writeHead(200, {
          'Content-Type': MIME[extname(file).toLowerCase()] || 'application/octet-stream',
          'Cache-Control': 'no-store, must-revalidate',
        });
        res.end(data);
        return;
      } catch {
        throw new HttpError(404, 'not found');
      }
    } catch (err) {
      if (err instanceof HttpError) return sendJson(res, err.status, { error: err.message });
      console.error(`[request failed] ${req.method} ${req.url}:`, err);
      return sendJson(res, 500, { error: (err as Error).message });
    }
  };
}

/* ----------------------------------------------------------------------- cli */

function buildAi(): AiClient & { keySource?: string | null } {
  const cfg = resolveConfig();
  if (cfg.deepseek.apiKey) {
    const source = process.env.DEEPSEEK_API_KEY ? 'the DEEPSEEK_API_KEY environment variable' : cfg.configFile;
    return {
      ...createDeepSeekClient({ apiKey: cfg.deepseek.apiKey, model: cfg.deepseek.model || DEFAULT_MODEL }),
      keySource: source,
    };
  }
  return {
    model: 'unconfigured',
    keySource: null,
    async complete() {
      throw new HttpError(
        503,
        'No DeepSeek API key is configured. Open Settings (the gear in the header) to add one, ' +
          'or run:  node app/setkey.js'
      );
    },
  };
}

function buildSearch(): SearchClient | null {
  const cfg = resolveConfig();
  if (!cfg.search.apiKey) return null;
  try {
    return createSearchClient({ provider: cfg.search.provider, apiKey: cfg.search.apiKey });
  } catch (err) {
    console.error(`[search] disabled: ${(err as Error).message}`);
    return null;
  }
}

export interface StartedServer {
  /**
   * Closes every listener. An object rather than a `Server` because loopback is
   * bound per address family and that takes more than one listener.
   */
  server: { close: () => void };
  scheme: string;
  port: number;
  url: string;
  /** Which loopback addresses actually answered, IPv4 first. */
  addresses: string[];
  ai: AiClient & { keySource?: string | null };
  search: SearchClient | null;
  tls: boolean;
  /** Where the certificate came from, for the startup line and the setup panel. */
  tlsSource: string | null;
}

/*
 * Where to listen. `localhost` resolves to either family depending on the client
 * and the resolver, so both are bound.
 */
const LOOPBACK_ADDRESSES = ['127.0.0.1', '::1'];

/**
 * Listen on loopback ONLY.
 *
 * `server.listen(port)` with no host binds every interface — this server was
 * reachable from the local network, which quietly contradicted the whole design:
 * the certificate exception is loopback-scoped, the credentials endpoint checks
 * its origin, and the app promises that nothing leaves the machine. A LAN peer
 * could reach all of it.
 *
 * Bound per family because macOS resolves `localhost` to `::1` first, and a
 * server on 127.0.0.1 alone would make every client pay a failed connection
 * first. A family that is unavailable (IPv6 disabled) is skipped, not fatal —
 * as long as one address answers.
 */
async function listenOnLoopback(
  make: () => Server,
  port: number
): Promise<{ close: () => void; addresses: string[]; port: number; primary: Server }> {
  const bound: Server[] = [];
  const addresses: string[] = [];
  let actualPort = port;

  const closeAll = (): void => {
    for (const server of bound) server.close();
  };

  try {
    for (const address of LOOPBACK_ADDRESSES) {
      const server = make();
      try {
        await new Promise<void>((resolve, reject) => {
          const onError = (err: Error): void => reject(err);
          server.once('error', onError);
          server.listen(actualPort, address, () => {
            server.off('error', onError);
            resolve();
          });
        });
      } catch (err) {
        server.close();

        /*
         * A family this machine does not have is skipped; a port that is already
         * taken is NOT. Tolerating EADDRINUSE let a second instance bind only ::1
         * and silently shadow the first for any client that prefers IPv6 — which
         * is exactly how a stale preview server ended up serving this app's
         * window a stub model. A half-bound server is worse than no server.
         */
        const code = (err as NodeJS.ErrnoException).code;
        const familyUnavailable = code === 'EAFNOSUPPORT' || code === 'EADDRNOTAVAIL' || code === 'EINVAL';
        if (!familyUnavailable) throw err;
        continue;
      }

      bound.push(server);
      addresses.push(address);
      // Port 0 means "pick one", and the second family has to land on the same
      // one — otherwise the two listeners would answer on different ports.
      const chosen = server.address();
      if (chosen && typeof chosen === 'object') actualPort = chosen.port;
    }
  } catch (err) {
    closeAll();
    throw err;
  }

  if (!bound.length) throw new Error(`could not listen on any loopback address for port ${port}`);

  return {
    addresses,
    port: actualPort,
    primary: bound[0] as Server,
    close: closeAll,
  };
}

/** Shared by the CLI and the Electron main process. */

/** Shared by the CLI and the Electron main process. */

export async function startServer({
  port = PORT,
  ai,
  search,
}: {
  port?: number;
  ai?: AiClient & { keySource?: string | null };
  search?: SearchClient | null;
} = {}): Promise<StartedServer> {
  /*
   * Injecting a client means the caller owns configuration, and this server must
   * not be able to rebuild one from a key on disk. That is what stops `--smoke`,
   * which passes a stub model, from ever constructing a real keyed client when
   * someone opens the settings dialog during a test run.
   */
  const injected = ai !== undefined || search !== undefined;
  const resolvedAi = ai ?? buildAi();
  const resolvedSearch = search ?? buildSearch();

  const credentials: Credentials | undefined = injected
    ? undefined
    : {
        read: describeConfig,
        apply(patch) {
          applyCredentialPatch(patch);
          return { ai: buildAi(), search: buildSearch() };
        },
      };

  const manifestTemplate = await readFile(join(WEB_DIR, 'manifest.template.xml'), 'utf8');

  const tls = resolveTls();
  const tlsOptions = tls ? { cert: await readFile(tls.cert), key: await readFile(tls.key) } : null;

  const app = createApp({
    ai: resolvedAi,
    search: resolvedSearch,
    credentials,
    // Reporting only: the app already did the certificate and the manifest.
    setupped: injected ? undefined : { manifestTemplate, manifestUrl: `https://localhost:${port}`, tlsSource: tls?.source ?? null },
  });
  const onRequest = (req: IncomingMessage, res: ServerResponse): void => {
    app(req, res).catch((err: unknown) => {
      console.error('[unhandled]', err);
      if (!res.headersSent) res.writeHead(500, { 'Content-Type': 'text/plain' });
      res.end('server error');
    });
  };

  const makeServer = (): Server =>
    tlsOptions ? createSecureServer(tlsOptions, onRequest) : createServer(onRequest);

  const listener = await listenOnLoopback(makeServer, port);

  const scheme = tls ? 'https' : 'http';
  return {
    server: listener,
    scheme,
    port: listener.port,
    url: `${scheme}://localhost:${listener.port}`,
    addresses: listener.addresses,
    ai: resolvedAi,
    search: resolvedSearch,
    tls: Boolean(tls),
    tlsSource: tls?.source ?? null,
  };
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1];

if (isMain) {
  const started = await startServer();
  console.log(`AI Spreadsheet Automation — ${started.url}`);
  console.log(`  model: ${started.ai.model}`);
  if (started.ai.keySource) {
    console.log(`  key:   loaded from ${started.ai.keySource}`);
  } else {
    console.log('  NOTE: no API key yet — the UI and diagnostics work, but /api/chat will refuse.');
    console.log('        Store one with:  node app/setkey.js');
  }
  console.log(
    started.search
      ? `  search: ${started.search.provider} (web_search enabled)`
      : '  search: not configured — the assistant cannot look anything up.\n' +
          '        Enable it with:  node app/setkey.js --search'
  );
  if (!started.tls) {
    console.log('  WARNING: no trusted cert, so this is HTTP only and Excel will not load the taskpane.');
    console.log('  Run once: npx office-addin-dev-certs install');
  }
  console.log('');
  console.log(`  desktop UI: ${started.url}/ui/`);
}
