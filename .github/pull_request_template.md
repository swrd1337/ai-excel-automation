<!--
Thanks for the pull request. The short version of what matters here is in CONTRIBUTING.md and the
long version — with the reasons — is in AGENTS.md. The most common reasons a PR here needs changes
are listed below; deleting this comment is fine.
-->

## What this changes

<!-- One or two sentences. If it fixes an issue, "Fixes #123". -->

## Why

<!-- The reason it is worth doing. If you chose between options, say what the other one was. -->

## Checklist

- [ ] `npm test` passes (67 checks, offline)
- [ ] `npm run typecheck:ui` passes
- [ ] `npm run validate:manifest` passes — **if** I changed `app/web/manifest.template.xml`

- [ ] **If this adds or changes an Excel operation:** it has an undo, and a test proving the
      snapshot is taken *before* the write and that undo restores the state
- [ ] **If this changes Office.js behaviour:** I have **not** claimed it works. It needs the
      Diagnostics tab run in real Excel, and I have said what to look for
- [ ] **If this changes `app/tools.ts`:** the schema and the executor agree, and the schema is the
      only validator (no second source of truth)
- [ ] **If this touches certificates, the keychain, or Excel's add-in folder:** the test suite still
      cannot reach any of them (`AI_EXCEL_CERT_DIR`, `AI_EXCEL_WEF_DIR`, `AI_EXCEL_NO_KEYCHAIN=1`)
- [ ] I have not pasted an API key, a real keychain dump, or private workbook data into this PR

## Anything you want a reviewer to know

<!--
For example: "this is the part I am unsure about", or "this needed a decision, and I put one in
docs/DECISIONS.md because the next person will hit the same trap".
-->
