/*
 * Trust the local certificate, once.
 *
 * This is the one step of setup that cannot be silent, and it is not a gap in the
 * implementation: changing a user trust setting needs authorization, and macOS
 * raises a password prompt for it. The app deliberately does not run this by
 * itself, because doing so can block startup behind a dialog nobody asked for.
 *
 * The app generates and stores the certificate; this turns "stored" into
 * "trusted", which is what Excel actually requires — a certificate in the
 * keychain that macOS does not trust is refused with "the content is blocked
 * because it isn't signed by a valid security certificate".
 *
 *   npm run trust-cert
 *
 * To undo:  security remove-trusted-cert ~/.ai-excel/certs/localhost.crt
 */

import { execFileSync } from 'node:child_process';

import { certificateDir, isTrusted, trustCommand } from '../app/tls.js';

const cert = `${certificateDir()}/localhost.crt`;

if (isTrusted(cert)) {
  console.log('Already trusted — nothing to do.');
  process.exit(0);
}

console.log('macOS will ask for your password: trusting a certificate is an authorization change.\n');
try {
  execFileSync('security', ['add-trusted-cert', '-k', `${process.env.HOME}/Library/Keychains/login.keychain-db`, cert], {
    stdio: 'inherit',
  });
} catch {
  console.error(`\nThat did not work. Run it by hand to see the error:\n  ${trustCommand()}`);
  process.exit(1);
}

console.log(
  isTrusted(cert)
    ? '\nDone — Excel will now accept the add-in. Restart Excel if it is open.'
    : '\nThe command reported success but macOS still does not verify it.'
);
process.exit(isTrusted(cert) ? 0 : 1);
