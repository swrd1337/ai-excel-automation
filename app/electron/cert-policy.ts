/*
 * The one place this app accepts a certificate it cannot verify.
 *
 * Kept as a pure function in its own module because it is a security boundary:
 * the failure mode is silently trusting *any* certificate, so the scoping has to
 * be testable rather than eyeballed. See selftest.
 *
 * Why any exception is needed at all: macOS has the Office dev CA only in the
 * login keychain with no explicit trust settings, so WebKit (which Excel uses)
 * accepts it but Chromium does not, and Electron would refuse to load our own UI.
 * Verification is meaningless for this specific connection — we started that
 * server ourselves, in this process, on the loopback interface — but it is only
 * waived for exactly that origin. Everything else fails closed.
 */

// 0.0.0.0 is deliberately absent: it is a wildcard bind address, not loopback.
const LOOPBACK_HOSTS = new Set(['localhost', '127.0.0.1', '::1', '[::1]']);

export function isOwnLoopbackOrigin(target: string, ownPort: number): boolean {
  let url: URL;
  try {
    url = new URL(target);
  } catch {
    return false;
  }

  if (url.protocol !== 'https:' && url.protocol !== 'http:') return false;
  if (!LOOPBACK_HOSTS.has(url.hostname)) return false;

  const port = url.port ? Number(url.port) : url.protocol === 'https:' ? 443 : 80;
  return port === Number(ownPort);
}
