import { useCallback, useEffect, useRef, useState } from 'react';

/*
 * All server communication for the window.
 *
 * The chat reply is streamed: POST /api/chat responds with NDJSON, one event per
 * line, so the assistant's text is printed as it is written and each tool call
 * appears the moment it runs. Approvals are polled separately, because the turn
 * parks on them and nothing else would surface the card.
 */

const STATUS_MS = 3000;
const APPROVAL_MS = 1500;
const REVIEW_KEY = 'ai-excel.review';
const SEARCH_KEY = 'ai-excel.search';

export interface ChatEntry {
  id: string;
  kind: 'user' | 'assistant' | 'tool' | 'search' | 'approval' | 'undo';
  at: number;
  text?: string;
  name?: string;
  detail?: string;
  ok?: boolean;
  query?: string;
  failed?: boolean;
}

export interface Approval {
  id: string;
  tool: string;
  /** What the change is aimed at: a range address, or a created object's name. */
  target: string;
  args: Record<string, unknown>;
  before: unknown;
}

export interface UsageSummary {
  calls: number;
  searches: number;
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  cacheHitTokens: number;
  cacheHitRate: number;
  tokensPerSecond: number;
  usd: number;
  estimated: boolean;
  peak: boolean;
  pricedAs: string;
  pricingAsOf: string;
}

export interface ServerStatus {
  excelConnected: boolean;
  model: string | null;
  search: string | null;
  undoAvailable: number;
  usage: UsageSummary | null;
  /** False when the server was handed its model, so it cannot manage keys at all. */
  configurable?: boolean;
  /** Which certificate the server is serving with — see the Details panel. */
  tlsSource?: string | null;
}

/**
 * A key as the window is allowed to know it. The server sends only a masked form —
 * the raw key exists in the config file and in the server's model client, and
 * nowhere else. Anything that would put it back on this side of the wire is a bug.
 */
export interface KeyState {
  set: boolean;
  masked: string | null;
  /** Which value is in charge: an environment variable beats the file. */
  source: 'env' | 'file' | null;
  envVar: string | null;
}

export interface ConfigState {
  file: string;
  deepseek: KeyState & { model: string | null };
  search: KeyState & { provider: string };
}

/** `null` clears a value; a field left out is left alone. */
export interface CredentialPatch {
  deepseekApiKey?: string | null;
  searchProvider?: string;
  searchApiKey?: string | null;
}

/**
 * What the app managed to do by itself about Excel, and what is left for a
 * person. `remaining` never mentions inserting the add-in — only the window knows
 * whether Excel is connected, so it adds that step (see `SetupPanel`).
 */
export interface SetupState {
  certificate: { dir: string; ready: boolean; detail: string; notAfter: string | null; trusted: boolean };
  manifest: {
    dir: string;
    path: string | null;
    installed: boolean;
    current: boolean;
    /** The folder could not be opened at all — a permission problem, not a missing file. */
    blocked: boolean;
    version: string | null;
    detail: string;
  };
  manifestDir: string;
  excelRunning: boolean;
  /** Which certificate the server is serving with, for the Details panel. */
  tlsSource: string | null;
}

/** Mirrors ChatEvent in app/server.ts — the wire contract. */
type ChatEvent =
  | { type: 'status'; text: string }
  | { type: 'assistant_start' }
  | { type: 'delta'; text: string }
  | { type: 'tool'; name: string; ok: boolean; detail?: string }
  | { type: 'search'; query: string }
  | { type: 'approval'; name: string; ok: boolean; detail?: string }
  | { type: 'done'; reply: string; stats: UsageSummary; elapsedMs: number }
  | { type: 'error'; error: string };

async function jsonFetch<T>(path: string, options?: RequestInit): Promise<T> {
  const res = await fetch(path, { cache: 'no-store', ...options });
  const text = await res.text();

  let body: unknown = {};
  if (text) {
    try {
      body = JSON.parse(text);
    } catch {
      throw new Error(`Unexpected response (HTTP ${res.status})`);
    }
  }
  if (!res.ok) {
    throw new Error((body as { error?: string }).error || `HTTP ${res.status}`);
  }
  return body as T;
}

const postJson = <T>(path: string, body: unknown): Promise<T> =>
  jsonFetch<T>(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });

let nextId = 0;
const entry = (kind: ChatEntry['kind'], props: Partial<ChatEntry> = {}): ChatEntry => ({
  id: `e${(nextId += 1)}`,
  kind,
  at: Date.now(),
  ...props,
});

export function useServer(sessionId: string) {
  const [reachable, setReachable] = useState(true);
  // `search === null` means "no provider"; before the first poll we simply do not
  // know, and saying "not configured" then would be a lie.
  const [statusLoaded, setStatusLoaded] = useState(false);
  const [status, setStatus] = useState<ServerStatus>({
    excelConnected: false,
    model: null,
    search: null,
    undoAvailable: 0,
    usage: null,
  });
  const [entries, setEntries] = useState<ChatEntry[]>([]);
  const [approvals, setApprovals] = useState<Approval[]>([]);
  const [busy, setBusy] = useState(false);
  const [review, setReviewState] = useState<boolean>(() => {
    const stored = localStorage.getItem(REVIEW_KEY);
    return stored === null ? true : stored === 'true'; // safest default for a first run
  });

  /*
   * Off unless asked for. Having a search key configured is not consent to send
   * the user's questions to a search provider, so the default is off and the
   * server treats a missing flag as off too.
   */
  const [webSearch, setWebSearchState] = useState<boolean>(() => localStorage.getItem(SEARCH_KEY) === 'true');

  const busyRef = useRef(false);
  busyRef.current = busy;

  const push = useCallback((item: ChatEntry) => setEntries((prev) => [...prev, item]), []);

  const setReview = useCallback((value: boolean) => {
    setReviewState(value);
    localStorage.setItem(REVIEW_KEY, String(value));
  }, []);

  const setWebSearch = useCallback((value: boolean) => {
    setWebSearchState(value);
    localStorage.setItem(SEARCH_KEY, String(value));
  }, []);

  /* ------------------------------------------------------------- polling */

  const pollStatus = useCallback(async () => {
    try {
      const s = await jsonFetch<ServerStatus & { undoAvailable: number }>(`/api/status?sessionId=${sessionId}`);
      setReachable(true);
      setStatusLoaded(true);
      setStatus({
        excelConnected: Boolean(s.excelConnected),
        model: s.model,
        search: s.search,
        undoAvailable: s.undoAvailable || 0,
        usage: s.usage || null,
      });
    } catch {
      setReachable(false);
    }
  }, [sessionId]);

  const pollApprovals = useCallback(async () => {
    try {
      const { approvals: list } = await jsonFetch<{ approvals: Approval[] }>(`/api/approvals?sessionId=${sessionId}`);
      setApprovals(list);
    } catch {
      // Transient.
    }
  }, [sessionId]);

  useEffect(() => {
    void pollStatus();
    const timer = setInterval(() => void pollStatus(), STATUS_MS);
    return () => clearInterval(timer);
  }, [pollStatus]);

  /*
   * Polled unconditionally, not only while this window is busy: a pending
   * approval must surface whenever it exists, or the turn it belongs to sits
   * blocked with nothing on screen to answer it.
   */
  useEffect(() => {
    void pollApprovals();
    const timer = setInterval(() => void pollApprovals(), APPROVAL_MS);
    return () => clearInterval(timer);
  }, [pollApprovals]);

  /* --------------------------------------------------------------- actions */

  const send = useCallback(
    async (text: string) => {
      if (busyRef.current || !text.trim()) return;
      setBusy(true);
      busyRef.current = true;

      const message = text.trim();
      push(entry('user', { text: message }));

      // The currently-open assistant bubble, if the stream has started one.
      let openId: string | null = null;
      let streamed = '';

      const openBubble = (): string => {
        const e = entry('assistant', { text: '' });
        openId = e.id;
        push(e);
        return e.id;
      };

      const handle = (event: ChatEvent): void => {
        switch (event.type) {
          case 'assistant_start':
            openBubble();
            break;
          case 'delta':
            if (!openId) openBubble();
            streamed += event.text;
            setEntries((prev) =>
              prev.map((e) => (e.id === openId ? { ...e, text: streamed } : e))
            );
            break;
          case 'tool':
            push(entry('tool', { name: event.name, ok: event.ok, detail: event.detail }));
            break;
          case 'search':
            push(entry('search', { query: event.query }));
            break;
          case 'approval':
            push(entry('approval', { name: event.name, ok: event.ok, detail: event.detail }));
            break;
          case 'done': {
            const reply = event.reply;
            if (openId) {
              const id = openId;
              setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, text: reply } : e)));
            } else {
              push(entry('assistant', { text: reply }));
            }
            setStatus((prev) => ({ ...prev, usage: event.stats }));
            break;
          }
          case 'error':
            push(entry('assistant', { text: event.error, failed: true }));
            break;
          default:
            break;
        }
      };

      try {
        const res = await fetch('/api/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ sessionId, message, review, search: webSearch }),
          cache: 'no-store',
        });

        if (!res.ok || !res.body) {
          const body = (await res.json().catch(() => ({}))) as { error?: string };
          throw new Error(body.error || `HTTP ${res.status}`);
        }

        const reader = res.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });
          const lines = buffer.split('\n');
          buffer = lines.pop() ?? '';
          for (const line of lines) {
            if (!line.trim()) continue;
            try {
              handle(JSON.parse(line) as ChatEvent);
            } catch {
              // A malformed frame should not abandon the rest of the reply.
            }
          }
        }
      } catch (err) {
        const failure = err as Error;
        if (openId) {
          const id = openId;
          setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, text: failure.message, failed: true } : e)));
        } else {
          push(entry('assistant', { text: failure.message, failed: true }));
        }
      } finally {
        setApprovals([]);
        setBusy(false);
        busyRef.current = false;
        void pollStatus();
      }
    },
    [sessionId, review, webSearch, push, pollStatus]
  );

  const decide = useCallback(
    async (id: string, decision: 'approve' | 'decline') => {
      setApprovals((prev) => prev.filter((a) => a.id !== id));
      await postJson('/api/approvals/decision', { sessionId, id, decision }).catch(() => {});
    },
    [sessionId]
  );

  const undo = useCallback(async () => {
    if (busyRef.current || status.undoAvailable === 0) return;
    setBusy(true);
    busyRef.current = true;
    const pending = entry('assistant', { text: 'Undoing the last task…' });
    push(pending);

    const settle = (props: Partial<ChatEntry>): void =>
      setEntries((prev) => prev.map((e) => (e.id === pending.id ? { ...e, ...props } : e)));

    try {
      const body = await postJson<{ message: string; failed?: string[] }>('/api/undo', { sessionId });
      settle({ text: body.message, failed: Boolean(body.failed?.length) });
    } catch (err) {
      settle({ text: `Undo failed: ${(err as Error).message}`, failed: true });
    } finally {
      await pollStatus();
      setBusy(false);
      busyRef.current = false;
    }
  }, [sessionId, status.undoAvailable, push, pollStatus]);

  const loadConfig = useCallback(
    () => jsonFetch<{ configurable: boolean; config: ConfigState | null }>('/api/config'),
    []
  );

  /*
   * Reporting only. The app writes the certificate and the manifest before the
   * server starts, so there is nothing to "run" from here — which is also why a
   * test server physically cannot touch the user's machine.
   */
  const loadSetup = useCallback(
    () => jsonFetch<{ available: boolean; state: SetupState | null }>('/api/setup'),
    []
  );

  const saveConfig = useCallback(
    async (patch: CredentialPatch) => {
      const body = await postJson<{ ok: boolean; config: ConfigState; warning?: string | null }>('/api/config', patch);
      // The Search toggle and the model name both come from status, and a key that
      // was just added must enable them without a reload.
      await pollStatus();
      return body;
    },
    [pollStatus]
  );

  return {
    reachable,
    statusLoaded,
    status,
    entries,
    approvals,
    busy,
    review,
    setReview,
    webSearch,
    setWebSearch,
    send,
    decide,
    undo,
    loadConfig,
    saveConfig,
    loadSetup,
  };
}
