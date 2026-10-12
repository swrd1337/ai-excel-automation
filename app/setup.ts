/*
 * Everything the app can do for itself so that Excel can reach it.
 *
 * The old first run was: install a certificate, find a hidden folder in
 * ~/Library/Containers, copy a manifest into it, restart Excel, insert the
 * add-in. Four steps, three of them undocumented anywhere a user would look.
 * All but the last two are the app's job, and they are all just files.
 *
 * What this module will NOT pretend to do:
 *
 *   - It cannot restart Excel. Excel reads the `wef` folder only at startup, so
 *     after a fresh install the user must quit and reopen it. We say so.
 *   - It cannot insert the add-in. Office exposes no way for an outside process
 *     to do that, and the first insertion is a manual step — after which the
 *     taskpane can ask to start itself (see `app/web/taskpane.ts`).
 *
 * Both are reported as `remaining` rather than left for the user to discover.
 */

import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

import { describeCertificate, ensureCertificate, type CertificateState } from './tls.js';

/** Where Office on macOS looks for sideloaded add-ins. Excel reads it at startup only. */
export function manifestDir(): string {
  if (process.env.AI_EXCEL_WEF_DIR) return process.env.AI_EXCEL_WEF_DIR;
  if (process.platform !== 'darwin') return '';
  return join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/Documents/wef');
}

export interface ManifestState {
  dir: string;
  path: string | null;
  /** A manifest is on disk. */
  installed: boolean;
  /** ...and it is exactly the one this build would serve. */
  current: boolean;
  /**
   * The folder could not be read at all — a permission or sandbox problem, not a
   * missing file. Reported separately because the remedy is completely different
   * and the old message blamed the wrong one.
   */
  blocked: boolean;
  version: string | null;
  detail: string;
}

export interface CertificateReport extends CertificateState {
  /** macOS trusts it. Untrusted means Excel refuses to load anything from it. */
  trusted: boolean;
}

export interface SetupState {
  certificate: CertificateReport;
  manifest: ManifestState;
  excelRunning: boolean;
  /** What a person still has to do. Ordered, and phrased as instructions. */
  remaining: string[];
}

/** The manifest is a template; the server substitutes its own base URL. */
export function renderManifest(template: string, baseUrl: string): string {
  return template.replaceAll('{{BASE_URL}}', baseUrl);
}

export function readManifestVersion(xml: string): string | null {
  return /<Version>([^<]+)<\/Version>/.exec(xml)?.[1] ?? null;
}

/**
 * Whether the manifest on disk is the one we would write.
 *
 * Compared as an exact string rather than parsed: the file is generated, so any
 * difference at all means it was written by an older build or by hand, and both
 * are reasons to replace it.
 */
export function describeManifest(rendered: string, dir = manifestDir()): ManifestState {
  const base = {
    dir,
    path: null as string | null,
    installed: false,
    current: false,
    blocked: false,
    version: null as string | null,
  };

  if (!dir) {
    return { ...base, detail: 'sideloading is only automated on macOS so far' };
  }

  const path = join(dir, 'manifest.xml');
  if (!existsSync(path)) {
    return { ...base, path, detail: 'no manifest is installed yet' };
  }

  let onDisk: string;
  try {
    onDisk = readFileSync(path, 'utf8');
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    const denied = code === 'EPERM' || code === 'EACCES';
    return {
      ...base,
      path,
      blocked: true,
      detail: denied
        ? `Excel's add-in folder could not be opened (${code}). Nothing is wrong with the file — ` +
          'something is stopping this app from reading it, usually a sandbox. Run the app from your own terminal.'
        : `the installed manifest could not be read (${(err as Error).message})`,
    };
  }

  const version = readManifestVersion(onDisk);
  if (onDisk === rendered) {
    return { ...base, path, installed: true, current: true, version, detail: `up to date (version ${version ?? 'unknown'})` };
  }
  return {
    ...base,
    path,
    installed: true,
    version,
    detail: `out of date (version ${version ?? 'unknown'} on disk)`,
  };
}

export function installManifest(rendered: string, dir = manifestDir()): ManifestState {
  if (!dir) return describeManifest(rendered, dir);

  try {
    mkdirSync(dir, { recursive: true });
    writeFileSync(join(dir, 'manifest.xml'), rendered, { mode: 0o644 });
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    return {
      dir,
      path: join(dir, 'manifest.xml'),
      installed: false,
      current: false,
      blocked: code === 'EPERM' || code === 'EACCES',
      version: null,
      detail:
        code === 'EPERM' || code === 'EACCES'
          ? `Excel's add-in folder could not be written (${code}). The app is probably running inside ` +
            'a sandbox that does not allow it — run it from your own terminal.'
          : `could not write it: ${(err as Error).message}`,
    };
  }
  return describeManifest(rendered, dir);
}

/** Excel reads the manifest folder once, at startup, so a restart is part of setup. */
export function isExcelRunning(): boolean {
  if (process.platform !== 'darwin') return false;
  return spawnSync('pgrep', ['-x', 'Microsoft Excel'], { encoding: 'utf8' }).status === 0;
}

/**
 * What is left after the app has done its part.
 *
 * Deliberately says nothing about inserting the add-in. This module has no way
 * to know whether that has happened, and guessing produces the worst kind of
 * instruction: one that is wrong on the second run. The window knows, because it
 * knows whether Excel is connected, so it adds that step (see `SetupPanel`).
 */
export function setupState(rendered: string, certificate: CertificateReport, dir = manifestDir()): SetupState {
  const manifest = describeManifest(rendered, dir);
  const excelRunning = isExcelRunning();
  const remaining: string[] = [];

  if (!certificate.ready) {
    remaining.push(`The local HTTPS certificate is not ready: ${certificate.detail}`);
  } else if (!certificate.trusted) {
    /*
     * The failure that looks like something else entirely. macOS will happily let
     * the app serve an untrusted certificate; Excel then refuses the page with
     * "the content is blocked because it isn't signed by a valid security
     * certificate", which blames the certificate and never the trust setting.
     */
    remaining.push(
      'The local HTTPS certificate is not trusted by macOS, so Excel will refuse to load the add-in. ' +
        'Press Retry, or run:  security add-trusted-cert -k ~/Library/Keychains/login.keychain-db ' +
        `"${join(certificate.dir, 'ca.crt')}"`
    );
  }

  if (manifest.blocked) {
    remaining.push(manifest.detail);
  } else if (!manifest.installed) {
    remaining.push(`The Excel add-in has not been installed yet: ${manifest.detail}`);
  } else if (!manifest.current && excelRunning) {
    // Excel reads the add-in folder once, at startup, so a changed manifest only
    // takes effect after a full quit.
    remaining.push('Quit Excel completely (Cmd-Q) and reopen it, so it notices the updated add-in.');
  }

  return { certificate, manifest, excelRunning, remaining };
}

/* ------------------------------------------------------------------ startup */

export interface PrepareResult {
  certificate: { ready: boolean; trusted: boolean; detail: string; trustDetail: string } | null;
  manifest: ManifestState;
  /** Anything worth printing at startup. The window reports the detail itself. */
  notes: string[];
}

/**
 * Everything that can be done for the user without asking them anything.
 *
 * Called by the app's entry point **before** the server exists, which is the
 * point: after this returns, the server only ever *reports* on Excel, so no
 * server — including one a test started with a stub model — can touch the
 * machine. Setup used to live inside the server and needed a guard to stay out
 * of tests; now there is nothing to guard.
 *
 * Never throws. A missing certificate or an unwritable folder must leave the app
 * starting and saying so, not failing to open.
 */
export function prepareExcel({ port, template }: { port: number; template: string }): PrepareResult {
  const notes: string[] = [];
  let certificate: PrepareResult['certificate'] = null;

  try {
    const ensured = ensureCertificate();
    certificate = {
      ready: ensured.ready,
      trusted: ensured.trusted,
      detail: ensured.detail,
      trustDetail: ensured.trustDetail,
    };
    if (ensured.action === 'generated') notes.push(`generated a local certificate (${ensured.trustDetail})`);
    if (!ensured.trusted) notes.push('the certificate is NOT trusted by macOS, so Excel will refuse to load the add-in');
  } catch (err) {
    certificate = { ready: false, trusted: false, detail: (err as Error).message, trustDetail: '' };
    notes.push(`certificate: ${(err as Error).message}`);
  }

  const manifest = installManifest(renderManifest(template, `https://localhost:${port}`));
  if (!manifest.current) notes.push(`manifest: ${manifest.detail}`);

  return { certificate, manifest, notes };
}
