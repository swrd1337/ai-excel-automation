/*
 * Electron main process — the desktop window.
 *
 * The window loads the built React app from /ui/ on the local server, which runs
 * in THIS process, so quitting the app releases port 3443.
 *
 * If a server is already listening (because `npm run serve` is running in a
 * terminal), we attach to it instead of failing.
 *
 *   npm run desktop -- --smoke
 *
 * `--smoke` loads the UI in a hidden window, fails on any renderer console error,
 * asserts the app actually rendered, writes a screenshot (SMOKE_SHOT, default
 * /tmp/ai-excel-ui.png) and exits 0/1. It is how the UI gets verified without a
 * window appearing on screen.
 */

import { readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { app, BrowserWindow, dialog, nativeTheme, shell, type WebContents } from 'electron';
import { startServer } from '../server.js';
import type { AiClient } from '../ai.js';
import { isOwnLoopbackOrigin as isOwnOrigin } from './cert-policy.js';
import { prepareExcel } from '../setup.js';

const SMOKE = process.argv.includes('--smoke');

/*
 * A smoke run must never touch the live setup.
 *
 * Two guards, both learned the hard way:
 *
 *  1. It binds a DIFFERENT port. The user's Excel taskpane polls 3443 because
 *     that is baked into the sideloaded manifest — so a test server on another
 *     port cannot be reached by their workbook at all.
 *  2. It never uses a real API key. If the smoke has to start a server itself it
 *     installs a stub model that refuses to answer, so a stray test can neither
 *     spend money nor write to a real workbook.
 */
const DEFAULT_PORT = Number(process.env.PORT || 3443);
const SMOKE_PORT = Number(process.env.SMOKE_PORT || 3499);
const FALLBACK_PORT = SMOKE ? SMOKE_PORT : DEFAULT_PORT;
const SMOKE_SHOT = process.env.SMOKE_SHOT || '/tmp/ai-excel-ui.png';

// Lets the smoke test verify both themes without touching system settings.
if (SMOKE && process.env.SMOKE_THEME) {
  nativeTheme.themeSource = process.env.SMOKE_THEME as 'light' | 'dark' | 'system';
}

let win: BrowserWindow | null = null;
let runningServer: { close: () => void } | null = null;
let ownPort = FALLBACK_PORT;

const isOwnLoopbackOrigin = (target: string): boolean => isOwnOrigin(target, ownPort);

/*
 * macOS has the Office dev CA only in the login keychain with no explicit trust
 * settings, so WebKit (which Excel uses) accepts it but Chromium does not — the
 * window would refuse to load our own UI with ERR_CERT_AUTHORITY_INVALID.
 *
 * Verification is meaningless for this particular connection: we started that
 * server ourselves, in this process, on the loopback interface. So we accept the
 * certificate ONLY for our own loopback origin and port, and say so in the log.
 * Every other origin is verified normally.
 */
app.on('certificate-error', (event, _contents, url, error, _certificate, callback) => {
  if (isOwnLoopbackOrigin(url)) {
    console.warn(`[desktop] accepting "${error}" for our own loopback origin (${url})`);
    event.preventDefault();
    callback(true);
    return;
  }
  callback(false); // fail closed for anything else
});

const stubAi: AiClient = {
  model: 'smoke-stub',
  async complete() {
    throw new Error('the smoke test does not call a model');
  },
};

async function resolveServerUrl(): Promise<string | null> {
  try {
    /*
     * Everything the app can do about Excel happens here, before the server
     * exists — so the server has nothing to do but report. Skipped in smoke mode
     * for the same reason the stub model is: a test must not generate a
     * certificate, trust it, or write into Excel's folder.
     */
    if (!SMOKE) {
      const template = await readFile(join(dirname(fileURLToPath(import.meta.url)), '../web/manifest.template.xml'), 'utf8');
      const prepared = prepareExcel({ port: FALLBACK_PORT, template });
      for (const note of prepared.notes) console.log(`[setup] ${note}`);
      if (prepared.manifest.current) console.log('[setup] Excel add-in is installed and current');
    }

    const started = SMOKE
      ? await startServer({ port: SMOKE_PORT, ai: stubAi, search: null })
      : await startServer({ port: FALLBACK_PORT });
    runningServer = started.server;
    ownPort = started.port;
    console.log(`[desktop] started server on ${started.url}`);
    return started.url;
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === 'EADDRINUSE') {
      ownPort = FALLBACK_PORT;
      const url = `https://localhost:${FALLBACK_PORT}`;
      console.log(`[desktop] port ${FALLBACK_PORT} already in use — attaching to ${url}`);
    if (SMOKE) console.log('[smoke] (attached to a preview server, not starting one)');
      return url;
    }
    dialog.showErrorBox('Could not start', String((err as Error).message || err));
    return null;
  }
}

function makeWindow(): BrowserWindow {
  return new BrowserWindow({
    width: 960,
    height: 780,
    minWidth: 640,
    minHeight: 520,
    show: !SMOKE,
    title: 'AI Spreadsheet Automation',
    backgroundColor: '#ffffff',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      // A hidden window throttles timers, which would stall the smoke test's
      // polling and leave the page half-rendered.
      backgroundThrottling: false,
    },
  });
}

/** Collects renderer console errors — an uncaught React error lands here. */
function watchConsole(contents: WebContents, errors: string[]): void {
  contents.on('console-message', (...args: unknown[]) => {
    // Electron <= 34 passes (event, level, message, line, sourceId);
    // newer versions pass a single details object.
    const first = args[0] as Record<string, unknown>;
    const modern = first && typeof first === 'object' && 'level' in first && args.length === 1;
    const level = modern ? first.level : (args[1] as number);
    const message = modern ? String(first.message) : String(args[2]);
    const source = modern ? String(first.sourceId ?? '') : String(args[4] ?? '');
    const line = modern ? Number(first.lineNumber ?? 0) : Number(args[3] ?? 0);

    const isError = level === 3 || level === 'error';
    if (isError) errors.push(`${message} (${source}:${line})`);
  });

  contents.on('did-fail-load', (_e, code, description, url) => {
    errors.push(`did-fail-load ${code} ${description} ${url}`);
  });
}

interface Probe {
  root?: boolean;
  title?: string | null;
  hasComposer?: boolean;
  approvalsFromPage?: number | string;
  buttons?: string[];
}

async function runSmoke(errors: string[]): Promise<void> {
  // The bundle is ~half a megabyte unminified; at 900ms the page was still
  // parsing and the first approval poll had not landed yet.
  const settle = Number(process.env.SMOKE_SETTLE_MS || 1500);
  await new Promise((r) => setTimeout(r, settle));

  // Click a control before capturing, so toggled states can be reviewed.
  if (process.env.SMOKE_CLICK && win) {
    // Waits for the control to exist AND be enabled: a fixed sleep races the
    // first status poll, and "clicked nothing" looks exactly like "the toggle
    // does not work".
    const outcome = await win.webContents.executeJavaScript(
      `(async () => {
         const sel = ${JSON.stringify(process.env.SMOKE_CLICK)};
         for (let i = 0; i < 50; i++) {
           const el = document.querySelector(sel);
           if (el && !el.disabled) {
             const before = el.getAttribute('data-state');
             el.click();
             await new Promise((r) => setTimeout(r, 250));
             const after = el.getAttribute('data-state');
             return (el.getAttribute('aria-label') || el.tagName) + ': ' + before + ' -> ' + after;
           }
           await new Promise((r) => setTimeout(r, 100));
         }
         const el = document.querySelector(sel);
         return el ? 'never became enabled' : 'never appeared: ' + sel;
       })()`
    );
    console.log(`[smoke] click: ${outcome}`);
    await new Promise((r) => setTimeout(r, 400));
  }

  /*
   * Optionally drive the composer, so a streamed reply can actually be observed.
   * React ignores a plain `el.value = …`, hence the native setter.
   */
  if (process.env.SMOKE_SEND && win) {
    const text = JSON.stringify(process.env.SMOKE_SEND);
    await win.webContents.executeJavaScript(
      `(() => {
         const ta = document.querySelector('textarea');
         if (!ta) return false;
         const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
         setter.call(ta, ${text});
         ta.dispatchEvent(new Event('input', { bubbles: true }));
         return true;
       })()`
    );
    await new Promise((r) => setTimeout(r, 150));
    await win.webContents.executeJavaScript(
      `(() => {
         const btn = [...document.querySelectorAll('button')].find((b) => b.textContent.trim().startsWith('Send'));
         btn?.click();
         return Boolean(btn);
       })()`
    );
    await new Promise((r) => setTimeout(r, Number(process.env.SMOKE_SEND_WAIT || 1200)));
  }

  // Optionally open a panel before capturing, so its contents can be reviewed.
  if ((process.env.SMOKE_OPEN === 'details' || process.env.SMOKE_OPEN === 'settings') && win) {
    const label = process.env.SMOKE_OPEN === 'settings' ? 'API keys' : 'Details';
    await win.webContents.executeJavaScript(
      `(() => { document.querySelector('[aria-label=${JSON.stringify(label)}]')?.click(); return true; })()`
    );
    await new Promise((r) => setTimeout(r, 400));
  }

  let probe: Probe = {};
  try {
    probe = (await win?.webContents.executeJavaScript(
      `(async () => {
         let approvalsFromPage = null;
         try {
           const r = await fetch('/api/approvals?sessionId=default', { cache: 'no-store' });
           approvalsFromPage = (await r.json()).approvals.length;
         } catch (e) {
           approvalsFromPage = 'ERR ' + e.message;
         }
         return {
           root: Boolean(document.getElementById('root')?.firstElementChild),
           title: document.querySelector('h1')?.textContent ?? null,
           hasComposer: Boolean(document.querySelector('textarea')),
           approvalsFromPage,
           buttons: [...document.querySelectorAll('button')].map((b) => b.textContent.trim()).filter(Boolean),
         };
       })()`
    )) as Probe;
  } catch (err) {
    errors.push(`could not inspect the page: ${(err as Error).message}`);
  }

  try {
    // A hidden window does not repaint, so capturePage would return the frame
    // from before the last state update. Force one, then give it a moment.
    win?.webContents.invalidate();
    await new Promise((r) => setTimeout(r, 300));

    const image = await win?.webContents.capturePage();
    if (image) {
      await writeFile(SMOKE_SHOT, image.toPNG());
      console.log(`[smoke] screenshot: ${SMOKE_SHOT}`);
    }
  } catch (err) {
    errors.push(`capture failed: ${(err as Error).message}`);
  }

  console.log(`[smoke] rendered: ${JSON.stringify(probe)}`);

  if (!probe.root) errors.push('the React app did not render into #root');
  if (!probe.hasComposer) errors.push('the composer is missing');

  if (errors.length) {
    console.error(`[smoke] FAILED (${errors.length} problem(s)):`);
    for (const e of errors) console.error(`  - ${e}`);
    runningServer?.close();
    app.exit(1);
    return;
  }

  console.log('[smoke] OK — UI rendered with no console errors');
  runningServer?.close();
  app.exit(0);
}

async function boot(): Promise<void> {
  const url = await resolveServerUrl();
  if (!url) {
    app.quit();
    return;
  }

  win = makeWindow();

  const consoleErrors: string[] = [];
  watchConsole(win.webContents, consoleErrors);

  // Links open in the real browser, never inside the app window.
  win.webContents.setWindowOpenHandler(({ url: target }) => {
    void shell.openExternal(target);
    return { action: 'deny' };
  });

  try {
    await win.loadURL(`${url}/ui/`);
  } catch (err) {
    if (!SMOKE) throw err;
    console.error(`[smoke] FAILED to load ${url}/ui/: ${(err as Error).message}`);
    runningServer?.close();
    app.exit(1);
    return;
  }

  if (SMOKE) {
    await runSmoke(consoleErrors);
    return;
  }

  win.on('closed', () => {
    win = null;
  });
}

app
  .whenReady()
  .then(boot)
  .catch((err: unknown) => {
    console.error('[desktop] boot failed:', err);
    app.exit(1);
  });

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) void boot();
});

app.on('window-all-closed', () => {
  app.quit();
});

app.on('before-quit', () => {
  runningServer?.close();
});
