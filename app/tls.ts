/*
 * The certificate this app serves itself with.
 *
 * Excel will not load a taskpane over plain HTTP — a `http://localhost` manifest
 * is rejected outright by the schema ("Insecure Desktop Source Location") — so
 * the app has to present a certificate macOS trusts. Asking the user to run
 * `npx office-addin-dev-certs install` is the setup step this exists to delete.
 *
 * It is ONE self-signed certificate for `localhost`, trusted directly.
 *
 * An earlier version generated a local CA and signed a leaf with it, mirroring
 * what the Office dev tooling does. That was worse in a way that is easy to miss:
 * trusting a CA grants it the power to vouch for **any** hostname, while trusting
 * a single leaf grants it power over `localhost` only. The narrower thing is also
 * the smaller thing — no chain, no second key, no `ca.srl`.
 *
 * macOS will store a certificate in the keychain without trusting it, and Office
 * refuses that with "the content is blocked because it isn't signed by a valid
 * security certificate" — an error naming neither the certificate nor the app.
 * Trust is therefore granted explicitly and then *checked*; see `isTrusted`.
 *
 * `openssl` does the X.509, because Node cannot create a certificate — it can only
 * parse one. That is an external binary, not an npm dependency, and its absence
 * is reported rather than guessed at.
 */

import { spawnSync } from 'node:child_process';
import { X509Certificate } from 'node:crypto';
import { chmodSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { CONFIG_DIR } from './config.js';

/** Ten years: a certificate that expires is a support ticket. */
const VALID_DAYS = 3650;

/** Regenerate this long before expiry, so nobody meets a broken morning. */
const RENEW_WITHIN_DAYS = 30;

/** The name the earlier CA-based design used. Removed if found, see `forgetOldCA`. */
const RETIRED_CA_NAME = 'AI Excel Automation Local CA';

export interface CertificateState {
  dir: string;
  /** True when this can serve HTTPS that macOS, and therefore Excel, will accept. */
  ready: boolean;
  /** One line a person can act on: why it is not ready, or when it expires. */
  detail: string;
  notAfter: string | null;
}

export interface EnsuredCertificate extends CertificateState {
  action: 'reused' | 'generated';
  /**
   * Whether macOS trusts it. The distinction that matters: a certificate in the
   * keychain is not a trusted one, and only the second works in Office.
   */
  trusted: boolean;
  /** What was done about it, for the startup line and the setup panel. */
  trustDetail: string;
}

export function certificateDir(): string {
  return process.env.AI_EXCEL_CERT_DIR || join(CONFIG_DIR, 'certs');
}

export function certificatePaths(dir = certificateDir()): { cert: string; key: string } {
  return { cert: join(dir, 'localhost.crt'), key: join(dir, 'localhost.key') };
}

export interface EnsureOptions {
  /**
   * Whether the certificate may be trusted on this machine. A test must pass
   * `false`: unlike the journal or the manifest, a mistake here leaves a trusted
   * certificate behind, and the suite runs on the user's machine.
   */
  trust?: boolean;
}

/* ------------------------------------------------------------------ openssl */

let opensslBinary: string | null | undefined;

/**
 * The first `openssl` that actually runs. Apple's LibreSSL at /usr/bin/openssl is
 * what a fresh Mac has; a newer one on PATH is preferred when present, but neither
 * is assumed.
 */
export function resolveOpenssl(): string | null {
  if (opensslBinary !== undefined) return opensslBinary;

  for (const candidate of [process.env.OPENSSL_BIN, 'openssl', '/usr/bin/openssl']) {
    if (!candidate) continue;
    const probe = spawnSync(candidate, ['version'], { encoding: 'utf8' });
    if (!probe.error && probe.status === 0) {
      opensslBinary = candidate;
      return candidate;
    }
  }
  opensslBinary = null;
  return null;
}

function openssl(args: string[], cwd: string): void {
  const binary = resolveOpenssl();
  if (!binary) throw new Error('openssl is not available, so a certificate cannot be generated');

  const res = spawnSync(binary, args, { cwd, encoding: 'utf8' });
  if (res.error) throw new Error(`${binary} could not run: ${res.error.message}`);
  if (res.status !== 0) {
    const last = (res.stderr || '').trim().split('\n').filter(Boolean).pop() || `exit ${res.status}`;
    throw new Error(`openssl ${args[0]} failed: ${last}`);
  }
}

/*
 * LibreSSL refuses to continue without a distinguished_name in [req]; OpenSSL 3
 * does not care. Writing the section keeps one recipe working on both, and -subj
 * overrides the placeholder.
 */
const CONFIG = `[req]
distinguished_name = dn
prompt = no

[dn]
CN = placeholder

[leaf_ext]
basicConstraints = critical,CA:FALSE
keyUsage = critical,digitalSignature,keyEncipherment
extendedKeyUsage = serverAuth
subjectAltName = DNS:localhost, IP:127.0.0.1, IP:::1
subjectKeyIdentifier = hash
`;

/* --------------------------------------------------------------- validation */

/**
 * Read and check the certificate, without trusting that it is what we think it is.
 *
 * `checkHost` is the only host check that can pass: Node implements the DNS name
 * case (`X509_check_host`) and not the IP one, so asking it about 127.0.0.1 always
 * returns undefined even when the SAN contains it. Do not "fix" that by asserting
 * on the IPs.
 */
export function describeCertificate(dir = certificateDir()): CertificateState {
  const { cert: certPath, key: keyPath } = certificatePaths(dir);
  const base = { dir, notAfter: null as string | null };

  if (!existsSync(certPath) || !existsSync(keyPath)) {
    return { ...base, ready: false, detail: 'no certificate yet' };
  }

  let cert: X509Certificate;
  try {
    cert = new X509Certificate(readFileSync(certPath));
  } catch (err) {
    return { ...base, ready: false, detail: `the stored certificate could not be read (${(err as Error).message})` };
  }

  /*
   * It must be self-signed, and that is not pedantry: a certificate macOS is
   * asked to *trust* has to be its own anchor, and a leaf from the retired CA
   * design is not — `add-trusted-cert` rejects it with a parameter error that
   * says nothing about why. Checking here turns that into a regeneration.
   */
  if (!cert.verify(cert.publicKey)) {
    return { ...base, ready: false, detail: 'the certificate is not self-signed (left over from the retired CA design)' };
  }

  if (!cert.checkHost('localhost')) {
    return { ...base, ready: false, detail: 'the certificate does not cover localhost' };
  }

  const expires = Date.parse(cert.validTo);
  if (!Number.isFinite(expires)) {
    return { ...base, ready: false, detail: 'the certificate has no readable expiry date' };
  }
  const daysLeft = (expires - Date.now()) / 86_400_000;
  if (daysLeft < RENEW_WITHIN_DAYS) {
    return {
      ...base,
      ready: false,
      detail: `the certificate expires in ${Math.max(0, Math.round(daysLeft))} days`,
      notAfter: cert.validTo,
    };
  }

  return { ...base, ready: true, detail: `valid until ${cert.validTo}`, notAfter: cert.validTo };
}

/* -------------------------------------------------------------- generation */

function generate(dir: string): void {
  const { cert, key } = certificatePaths(dir);
  mkdirSync(dir, { recursive: true, mode: 0o700 });
  chmodSync(dir, 0o700);

  writeFileSync(join(dir, 'openssl.cnf'), CONFIG, { mode: 0o600 });

  // Written fresh each time: a half-generated pair from an interrupted run must
  // never be mistaken for a usable one. The CA files are from the retired design.
  for (const stale of [cert, key, join(dir, 'ca.crt'), join(dir, 'ca.key'), join(dir, 'ca.srl')]) {
    rmSync(stale, { force: true });
  }

  openssl(['genrsa', '-out', 'localhost.key', '2048'], dir);
  openssl(
    ['req', '-x509', '-new', '-key', 'localhost.key', '-sha256', '-days', String(VALID_DAYS),
     '-subj', '/CN=localhost', '-config', 'openssl.cnf', '-extensions', 'leaf_ext', '-out', 'localhost.crt'],
    dir
  );

  chmodSync(key, 0o600);
  chmodSync(cert, 0o644);
}

/* ---------------------------------------------------------------- trusting */

function loginKeychain(): string {
  return join(process.env.HOME || '', 'Library/Keychains/login.keychain-db');
}

function inKeychain(commonName: string): boolean {
  const res = spawnSync('security', ['find-certificate', '-c', commonName, loginKeychain()], { encoding: 'utf8' });
  return res.status === 0;
}

/**
 * Does macOS actually trust this certificate?
 *
 * The one check that matters, and the one this project got wrong first time. Use
 * `-p basic`, not `-p ssl`: the ssl policy demands Certificate Transparency
 * timestamps that no locally-issued certificate will ever have, so it fails on
 * every certificate here and distinguishes nothing.
 */
export function isTrusted(certPath: string): boolean {
  if (process.platform !== 'darwin') return true; // no equivalent check elsewhere yet
  if (!existsSync(certPath)) return false;
  return spawnSync('security', ['verify-cert', '-p', 'basic', '-c', certPath], { encoding: 'utf8' }).status === 0;
}

/**
 * Remove the trusted CA the earlier design installed.
 *
 * Left behind it would be a trust anchor able to vouch for any hostname, for no
 * purpose — the current certificate is self-signed and does not use it. Best
 * effort: an older machine that never had one is the normal case.
 */
function forgetOldCA(): void {
  if (process.platform !== 'darwin') return;
  if (!inKeychain(RETIRED_CA_NAME)) return;
  const { cert } = certificatePaths();
  spawnSync('security', ['remove-trusted-cert', join(cert, '..', 'ca.crt')], { encoding: 'utf8' });
  spawnSync('security', ['delete-certificate', '-c', RETIRED_CA_NAME, loginKeychain()], { encoding: 'utf8' });
  if (!inKeychain(RETIRED_CA_NAME)) {
    console.log(`[tls] removed the retired local CA from the login keychain`);
  }
}

/**
 * Put the certificate in the keychain, and report whether macOS trusts it.
 *
 * **It does not grant trust, deliberately.** `add-trusted-cert` raises an
 * authorization prompt — measured, not assumed: it blocks indefinitely without
 * one — and macOS offers no non-interactive way to change a user trust setting.
 * Running it during startup therefore either hangs the app behind a password
 * dialog nobody was expecting, or silently succeeds on a machine that happens not
 * to ask. Neither is acceptable, so trust is the one thing the user does, with
 * the exact command printed for them (`npm run trust-cert`).
 *
 * Storing the certificate is silent, and is what makes that one command short.
 *
 * Replacement matters: a regenerated certificate has the same common name, so a
 * stale entry would leave the OLD one trusted while serving the new one — a
 * mismatch that presents as exactly the same broken-add-in error.
 */
function ensureStored(certPath: string): Pick<EnsuredCertificate, 'trusted' | 'trustDetail'> {
  if (process.platform !== 'darwin') {
    return { trusted: true, trustDetail: 'not macOS, so there is no keychain to trust it in' };
  }

  const trusted = isTrusted(certPath);
  if (trusted) return { trusted: true, trustDetail: 'trusted by macOS' };

  // Only remove an entry that is not this certificate, so the common case of a
  // second run touches nothing.
  if (inKeychain('localhost')) {
    const existing = spawnSync('security', ['find-certificate', '-c', 'localhost', '-p', loginKeychain()], {
      encoding: 'utf8',
    });
    const same = (existing.stdout || '').includes(readFileSync(certPath, 'utf8').trim().split('\n').slice(1, -1).join(''));
    if (!same) spawnSync('security', ['delete-certificate', '-c', 'localhost', loginKeychain()], { encoding: 'utf8' });
  }

  if (!inKeychain('localhost')) {
    const res = spawnSync('security', ['add-certificates', '-k', loginKeychain(), certPath], { encoding: 'utf8' });
    if (res.status !== 0) {
      const detail = (res.stderr || res.stdout || '').trim() || `exit ${res.status}`;
      return { trusted: false, trustDetail: `could not store it in the login keychain: ${detail}` };
    }
  }

  return { trusted: false, trustDetail: 'stored, but macOS does not trust it yet — see the setup panel' };
}

/** The one command that turns a stored certificate into a trusted one. */
export function trustCommand(dir = certificateDir()): string {
  return `security add-trusted-cert -k ~/Library/Keychains/login.keychain-db ${join(dir, 'localhost.crt')}`;
}

/* ------------------------------------------------------------------- entry */

/**
 * Make sure a usable, trusted certificate exists.
 *
 * Throws only when one is needed and cannot be produced at all — no openssl, or a
 * generation failure. Trust trouble is reported in the result rather than thrown,
 * because the app should still start and say what is wrong.
 */
export function ensureCertificate(dir = certificateDir(), options: EnsureOptions = {}): EnsuredCertificate {
  /*
   * AI_EXCEL_NO_KEYCHAIN exists so this can be exercised end to end without
   * putting a trusted certificate in the running user's keychain — the one side
   * effect of this module that is not a file in our own directory.
   */
  const allowTrust = options.trust !== false && process.env.AI_EXCEL_NO_KEYCHAIN !== '1';
  const { cert } = certificatePaths(dir);

  forgetOldCA();

  const existing = describeCertificate(dir);
  if (existing.ready) {
    if (!allowTrust) return { ...existing, action: 'reused', trusted: false, trustDetail: 'not requested' };
    if (isTrusted(cert)) {
      return { ...existing, action: 'reused', trusted: true, trustDetail: 'already trusted' };
    }
    // Usable but untrusted: macOS lets the app serve it, Office refuses anything
    // from it. Repair the trust rather than regenerate.
    return { ...existing, action: 'reused', ...ensureStored(cert) };
  }

  generate(dir);

  const after = describeCertificate(dir);
  if (!after.ready) throw new Error(`the generated certificate is not usable: ${after.detail}`);

  return {
    ...after,
    action: 'generated',
    ...(allowTrust ? ensureStored(cert) : { trusted: false, trustDetail: 'not requested' }),
  };
}
