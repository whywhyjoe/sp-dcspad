# State — reconcile work-prod drift into main

Last touched: 2026-09-17
Mode: Joe
Branch: main (PR #19 merged)
State: all code is on main; the only work left is the prod redeploy from the work machine

## What this is

DCSPad was live on the work (bmo) prod tenant from an uncommitted work-machine tree, and
that clone also held an unresolved `git stash pop`. Everything worth keeping has been
brought into git from the home machine and merged (PR #19); what remains is redeploying
prod from a fresh clone so prod and `main` are the same thing again.

## Done

- Work stash and work working tree contained nothing new (the stash was the
  `additionalTypes` feature, already on main as e7c9c0d). Discard both.
- The one real prod-only change — the export check-out → upload → metadata → check-in
  lifecycle — is merged with the home consent-gate branch in 63e89ad (+ bundle rebuild).
  The design is recorded in HANDOFF.md ("ForceCheckout" paragraphs) and README.md.
- An independent review of the PR found no blockers and seven real gaps (UPN-vs-claim
  identity, the real refusal sentence, blind new files, the stale Workbench listing, the
  dialog trap, the implicit check-in of `bNewDocumentUpdate`); all fixed in 2f7b580.
- Suites: files 63, workbench-edit 26, and every other suite green — except `monaco.mjs`,
  see Landmines.

## Next

- [ ] Work machine: delete the old clone (nothing in it is needed — check `git stash list`
      is only the known stale stash first), clone fresh, run
      `cd tools && npm install` with the machine's own Node, create
      `deploy/deploy.settings.local.json` with the prod `livePath`, then
      `deploy\Sync-Live.ps1 -Environment prod`. No `?v=` bump is needed: boot.js and
      boot-workbench.js are unchanged.
- [ ] Work tenant, README validation step 8: overwrite a file in a Require-Check-Out
      library from the pad and from Workbench Files; confirm it ends checked in, and that a
      brand-new file is not left checked out.
- [ ] Then delete this file, and the home machine's local-only branch
      `recovered/work-checkout-prod` (the literal prod source, kept as a reference until
      the tenant check passes).

## Open questions

- Joe, on the work tenant: the consent box now starts UNticked (Overwrite disabled until
  ticked). The prod build pre-ticked it. One-line change in `applySpCheckoutState`
  (`src/main.js`) if the pre-ticked behaviour is preferred.

## Landmines

- On 2026-09-17 `tests/monaco.mjs` began timing out on the home machine at its first hover
  check (`.monaco-hover-content` present but empty/hidden; 14 of 44 pass) — on untouched
  `origin/main` as well, an hour after passing 44/44. Environmental (headless Edge 153),
  not this change. Unverified guess: hover widgets not painting while the display is
  off/locked. Re-run with the screen awake before treating it as a regression.

- Never copy files back from the deployed SharePoint folder: the Document ID service
  injects `mso:CustomDocumentProperties` blocks into every .html it stores.
- `checkInFile()` deliberately reads `CheckOutType` before posting `CheckIn()`. Do not
  "simplify" it to an unconditional call: whether an overwrite leaves the check-out
  standing is unverified on either tenant, and CheckIn() on a checked-in file is an error.
- tests/files.mjs picker waits must treat "Connecting to SharePoint site…" as busy; the
  older "Loading SharePoint folder" only predicate raced deterministically on this machine.
