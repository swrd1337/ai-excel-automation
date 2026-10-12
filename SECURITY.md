# Security policy

## Reporting a vulnerability

Please use GitHub's private reporting: **Security → Report a vulnerability** on this repository.
That opens a draft advisory only you and the maintainer can see. Please do not open a public issue
for anything exploitable.

Useful to include: what you did, what happened, and what you expected. A proof of concept helps
enormously for this project, because most of its interesting surface is local.

## What this app's security model actually is

Worth stating plainly, since it shapes what counts as a vulnerability.

**It is a local application.** The server binds `127.0.0.1` and `::1` and nothing else, and a busy
port is fatal rather than tolerated — a half-bound server would be worse than none. It is not
intended to be reachable from your network, and a way to reach it from another host is a
vulnerability.

**It writes to your Excel workbooks.** That is its job, which makes two classes of bug
security-relevant rather than cosmetic:

- **A way to make it write somewhere you did not ask for** — a workbook you did not name, a range
  outside what it showed you, or a change that its own undo cannot reverse.
- **Anything that turns a normal conversation into code execution or file access** on your machine.
  Tool arguments come from a language model and are therefore untrusted input.

**It holds an API key.** The key lives in `~/.ai-excel/config.json` with mode `0600`, and is
deliberately never sent back to the window — `GET /api/config` returns only whether a key is set, a
masked form, and its source. A way to read the key value back through the API is a vulnerability.

**It installs a certificate.** The app generates a self-signed certificate for `localhost` and stores
it in your login keychain. **Trusting** it is a separate, deliberate step you perform
(`npm run trust-cert`) because macOS requires a password for any trust change — the app never grants
that trust unattended. The certificate covers `localhost` only; it is not a CA and cannot vouch for
any other hostname. If you find a way to make the app trust, or serve, a certificate it should not,
that is a vulnerability, and `app/electron/cert-policy.ts` is the code to look at.

**Cell data leaves your machine when you ask it a question.** The task, the conversation, and what
the tools read from the workbook are sent to your model provider to be answered. That is inherent to
using an LLM, it is why the provider and the key are yours, and it is documented in the README rather
than buried. It is not a vulnerability — but a way to make the app send data somewhere you did not
configure *would* be.

## Supported versions

This is pre-1.0 software under active development. Fixes go to the latest commit on `main`; there are
no maintained release branches yet.
