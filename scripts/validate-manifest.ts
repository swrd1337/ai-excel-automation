/*
 * Render the manifest the way the app will, then hand it to Microsoft's own
 * validator.
 *
 * Worth running after any change to `app/web/manifest.template.xml`. It found two
 * errors in a manifest Excel had been loading happily for weeks: `<Requirements>`
 * in a position the schema does not allow, and a `<Version>` below the 1.0
 * minimum. Excel is lenient about both; the store, and other clients, are not.
 *
 * Not part of `npm test`: the validator pulls its own copy of the tool on first
 * use and its acceptance step talks to Microsoft, and the test suite is offline by
 * design. The offline checks in `app/selftest.ts` cover the specific mistakes this
 * catches that would otherwise ship — element order, the version floor, and a
 * runtime `resid` that resolves to nothing.
 *
 *   npm run validate:manifest
 *   BASE_URL=https://localhost:3443 npm run validate:manifest
 */

import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { renderManifest } from '../app/setup.js';

const BASE_URL = process.env.BASE_URL || 'https://localhost:3443';
const template = readFileSync('app/web/manifest.template.xml', 'utf8');
const rendered = renderManifest(template, BASE_URL);

if (rendered.includes('{{BASE_URL}}')) {
  console.error('A {{BASE_URL}} placeholder survived rendering — the manifest would be rejected.');
  process.exit(1);
}

const dir = mkdtempSync(join(tmpdir(), 'ai-excel-manifest-'));
const file = join(dir, 'manifest.xml');
writeFileSync(file, rendered);

console.log(`Validating the manifest as app/setup.ts would write it (BASE_URL=${BASE_URL})\n`);
try {
  execFileSync('npx', ['--yes', 'office-addin-manifest', 'validate', file], { stdio: 'inherit' });
} catch {
  // The validator has already said what was wrong; do not bury it in a stack trace.
  process.exit(1);
}
