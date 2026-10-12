/*
 * The taskpane is now a WORKER, not a UI.
 *
 * It long-polls the server for work, executes whatever it receives via Office.js,
 * and posts the result back. The chat lives in the desktop window. This pane just
 * has to stay open so Excel has somewhere to run code.
 *
 * SESSION_ID is a wire contract shared with ui/src/lib/api.ts: one user, one
 * workbook, one conversation. (Caveat: with two workbooks' taskpanes open,
 * whichever is currently polling wins each call — use one workbook at a time.)
 */

import { getContext, executeToolCalls, EXECUTOR_NAMES, CONTEXT_TOOL, type ExecutorCall } from './executor.js';
import { initDiagnostics } from './diagnostics.js';

const SESSION_ID = 'default';
const RETRY_MS = 2000;
const MAX_LOG = 25;

const panelConn = document.getElementById('panel-conn') as HTMLElement;
const panelDiag = document.getElementById('panel-diag') as HTMLElement;
const dot = document.getElementById('conn-dot') as HTMLElement;
const connText = document.getElementById('conn-text') as HTMLElement;
const connLog = document.getElementById('conn-log') as HTMLElement;

let running = false;
let excelReady = false;
let knownBuild: string | null = null;

interface LogEntry {
  ok: boolean;
  time: string;
  text: string;
}

const entries: LogEntry[] = [];

/* --------------------------------------------------------------- rendering */

function setStatus(state: '' | 'on' | 'off', text: string): void {
  dot.className = `dot ${state}`;
  connText.textContent = text;
}

function logLine(ok: boolean, text: string): void {
  const time = new Date().toLocaleTimeString();
  entries.unshift({ ok, time, text });
  entries.length = Math.min(entries.length, MAX_LOG);
  connLog.replaceChildren(
    ...entries.map((e) => {
      const li = document.createElement('li');
      if (!e.ok) li.className = 'err';
      const t = document.createElement('span');
      t.className = 't';
      t.textContent = e.time;
      li.append(t, document.createTextNode(e.text));
      return li;
    })
  );
}

function fmtError(err: unknown): string {
  if (!err) return 'unknown error';
  const e = err as { code?: string; message?: string; debugInfo?: unknown };
  return [e.code, e.message, e.debugInfo && JSON.stringify(e.debugInfo)].filter(Boolean).join(' — ');
}

/* -------------------------------------------------------------- the worker */

async function postResult(id: string, content: unknown, error?: string): Promise<void> {
  await fetch('api/executor/result', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ sessionId: SESSION_ID, id, content, error }),
    cache: 'no-store',
  });
}

async function handleCall(call: ExecutorCall): Promise<void> {
  // The server asks for the workbook context itself rather than having the UI
  // supply it — that's what lets the UI be anything that can POST.
  if (call.name === CONTEXT_TOOL) {
    try {
      await postResult(call.id, await getContext());
    } catch (err) {
      await postResult(call.id, null, fmtError(err));
    }
    return;
  }

  const [result] = await executeToolCalls([call]);
  if (!result) return;
  await postResult(call.id, result.content, result.error);
  logLine(
    !result.error,
    result.error
      ? `${call.name} failed: ${result.error}`
      : `${call.name} ${JSON.stringify(result.content).slice(0, 120)}`
  );
}

/**
 * The server stamps every response with a build id derived from its taskpane
 * files. If it moves while this pane is open, this pane is running stale code —
 * reload rather than keep serving old tools.
 */
function checkBuild(res: Response): boolean {
  const build = res.headers.get('X-App-Build');
  if (!build) return false;
  if (knownBuild === null) {
    knownBuild = build;
    const el = document.getElementById('conn-build');
    if (el) el.textContent = build;
    return false;
  }
  if (build !== knownBuild) {
    setStatus('', 'The server was updated — reloading this pane…');
    setTimeout(() => location.reload(), 400);
    return true;
  }
  return false;
}

async function fetchWork(): Promise<ExecutorCall | null> {
  const params = new URLSearchParams({
    sessionId: SESSION_ID,
    capabilities: EXECUTOR_NAMES.join(','),
  });
  const res = await fetch(`api/executor/next?${params}`, { cache: 'no-store' });
  if (checkBuild(res)) return null;
  if (res.status === 204) return null;
  if (!res.ok) throw new Error(`server returned HTTP ${res.status}`);
  const body = (await res.json()) as { call?: ExecutorCall };
  return body.call || null;
}

async function loop(): Promise<void> {
  running = true;
  while (running) {
    try {
      const call = await fetchWork();
      if (!excelReady) {
        setStatus('', 'Excel is still loading…');
        await new Promise((r) => setTimeout(r, 300));
        continue;
      }
      setStatus('on', 'Connected — the desktop window can control this workbook');
      if (call) await handleCall(call);
    } catch (err) {
      setStatus('off', `Not connected: ${(err as Error).message}`);
      await new Promise((r) => setTimeout(r, RETRY_MS));
    }
  }
}

/* ------------------------------------------------------------------ startup */

document.querySelectorAll<HTMLButtonElement>('.tabs button').forEach((btn) => {
  btn.addEventListener('click', () => {
    document
      .querySelectorAll<HTMLButtonElement>('.tabs button')
      .forEach((b) => b.classList.toggle('active', b === btn));
    panelConn.classList.toggle('active', btn.dataset.tab === 'conn');
    panelDiag.classList.toggle('active', btn.dataset.tab === 'diag');
  });
});

initDiagnostics(panelDiag, { available: false }); // re-initialised with the real value below

const fallback = setTimeout(() => {
  setStatus('off', 'Not running inside Excel — open this from the add-in taskpane.');
}, 4000);

if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => {
    clearTimeout(fallback);
    excelReady = info.host === Office.HostType.Excel;
    initDiagnostics(panelDiag, { available: excelReady });
    if (!excelReady) {
      setStatus('off', `Office.js loaded but the host is ${info.host}, not Excel.`);
      return;
    }
    setStatus('', 'Connected — waiting for work');
    void loop();
  });
} else {
  clearTimeout(fallback);
  setStatus('off', 'office.js failed to load.');
}
