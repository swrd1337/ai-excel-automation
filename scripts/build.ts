/*
 * One build entry point, run directly by Node (v24 runs TypeScript natively, so
 * the tooling needs no compile step of its own):
 *
 *   node scripts/build.ts              clean, compile, copy static, bundle the UI
 *   node scripts/build.ts --no-ui      skip the UI bundle, for a quick server loop
 *
 * Sources stay pure TypeScript in app/ and ui/. Everything generated lands in
 * dist/, so there is never a question about which file to edit.
 */

import { spawn } from 'node:child_process';
import { cp, mkdir, readdir, rm } from 'node:fs/promises';
import { join, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const PROJECT = fileURLToPath(new URL('../', import.meta.url));
const DIST = join(PROJECT, 'dist');

/* Files served as-is to Excel and the window; tsc only handles .ts. */
const STATIC_SRC = join(PROJECT, 'app', 'web');
const STATIC_DEST = join(DIST, 'app', 'web');
const STATIC_FILES = ['taskpane.html', 'manifest.template.xml'];
const STATIC_DIRS = ['assets'];

const skipUi = process.argv.includes('--no-ui');

function run(command: string, args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: PROJECT,
      stdio: 'inherit',
      // npx is a shell script on Unix and a .cmd on Windows.
      shell: process.platform === 'win32',
    });
    child.on('error', reject);
    child.on('exit', (code) =>
      code === 0 ? resolve() : reject(new Error(`${command} ${args.join(' ')} exited with ${code}`))
    );
  });
}

const UI_DIR = join(DIST, 'app', 'web', 'ui');

/**
 * Delete everything under `dir` except the listed paths.
 *
 * `--no-ui` must not wipe `dist/app/web/ui`: it sits inside the tree being
 * cleaned, and removing it leaves the server serving a window with no bundle.
 */
async function cleanExcept(dir: string, keep: string[]): Promise<void> {
  const entries = await readdir(dir, { withFileTypes: true }).catch(() => []);
  for (const entry of entries) {
    const full = join(dir, entry.name);
    if (keep.includes(full)) continue;
    // A kept path may live deeper, so recurse down that branch rather than
    // deleting the whole subtree. dist/app must be entered, not removed, to
    // reach dist/app/web/ui.
    if (keep.some((k) => k.startsWith(full + sep))) {
      await cleanExcept(full, keep);
      continue;
    }
    await rm(full, { recursive: true, force: true });
  }
}

// Cleaned first so a deleted source can never linger as a stale build artefact.
if (skipUi) {
  await cleanExcept(DIST, [UI_DIR]);
} else {
  await rm(DIST, { recursive: true, force: true });
}

console.log('compiling TypeScript');
await run('npx', ['tsc', '-p', 'tsconfig.json']);

console.log('copying static assets');
await mkdir(STATIC_DEST, { recursive: true });
for (const file of STATIC_FILES) {
  await cp(join(STATIC_SRC, file), join(STATIC_DEST, file));
}
for (const dir of STATIC_DIRS) {
  await cp(join(STATIC_SRC, dir), join(STATIC_DEST, dir), { recursive: true });
}

if (!skipUi) {
  console.log('typechecking the UI');
  await run('npx', ['tsc', '-p', 'tsconfig.ui.json']);

  console.log('bundling the UI');
  await run('npx', ['vite', 'build']);
}

console.log(`built -> ${DIST}`);
