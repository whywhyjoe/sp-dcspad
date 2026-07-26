# DCSPad review log

Triage record for external code reviews: what was reported, what we agreed with,
what we rejected (and why), and the accepted-but-deferred backlog. Append a new
section per review round; move backlog items out when they land.

---

## 2026-07-25 — GHCP GPT 5.6 Sol review (devpad/ only, security excluded)

Two inputs were triaged: the reviewer's **final review** and an earlier
**chain-of-thought segment** from the same session. Every claim in both was
verified against the code before acting — all were factually accurate; the
disagreements below are about remedies and severity, not facts.

### Landed (this round)

| # | Issue | Fix |
|---|-------|-----|
| C1 | Network ids (`f1`, `x1`) restarted per iframe while the panel keeps cross-run history → new runs silently overwrote old `requests` map entries; old rows showed new data; multi-select glitch | Ids namespaced by run token in `harness.js`; `markRun()` now also settles still-pending rows as **cancelled** (they could never complete — the parent drops stale-token messages) |
| C2 | SP context + `__REQUESTDIGEST` captured once at bootstrap and reused forever, contradicting the README's freshness claim; digests expire (~30 min) | `getSpContext({ refresh: true })` re-captures per run in `run()`; README reworded to state exactly what is guaranteed on classic vs modern pages |
| S1 | Rerun called `evalCallbacks.clear()` without settling — a REPL command awaiting a slow promise hung forever after rerun | Abandoned callbacks are resolved with an explicit cancellation result before clearing |
| S3 | Harness fetch ignored `res.ok` — a 404 injected an HTML error page as the preview's first `<script>` | `initRunner` throws an actionable error; `run()` surfaces it in the status bar |
| S5 | Edits made inside the 600 ms autosave debounce window were lost on tab close/reload | `pagehide` listener flushes a pending save |
| S8 | No coverage of the rerun boundary (exactly where C1/S1 lived) | 5 new smoke checks: cross-run network history, per-run id integrity via detail pane, single-selection, cancelled pending request, cancelled REPL eval — all using `waitForFunction`/`page.route`, no fixed sleeps |

### Rejected (with justification)

- **S1's added eval timeout.** Awaiting a slow `sp.web.get()` on a sluggish
  tenant is the intended use case; DevTools doesn't time out evals either. The
  only true leak path was rerun, which the cancellation fix closes.
- **S4's streaming-prefix body read.** Accurate observation (`clone.text()`
  materializes whole bodies), but a `getReader()` loop in a classic no-modules
  script adds real complexity, loses accurate `size` reporting, and the target
  workload is modest SP JSON. The cheap half (content-type gating) is kept in
  the backlog below.
- **CoT issue: `Map`-based key lookup in `sp-shapes.js` `entity()`.** The linear
  search is real but bounded by construction: serialized nodes cap at
  `MAX_KEYS = 100` and the largest `ENTITY_FIELDS` list is 5 → worst case ~500
  string compares per entity, dwarfed by the DOM building around it. There is
  no scaling regime where it matters; revisit only if the future Site Inspector
  profiles hot here.

### Accepted but deferred — low-priority backlog

- [ ] **Run-scoped load timeout** (`main.js`): each run's 15 s "still loading…"
  timeout checks the shared `spinnerTimer`, so run A's timer can stomp run B's
  status if two slow loads overlap. Cosmetic; keep one `loadTimeout`, clear it
  in `run()` and in the `loaded` handler.
- [ ] **History caps**: console entries and network rows/`requests` map grow for
  the whole session. Cap at ~500, dropping the oldest DOM row + map entry
  together. Only matters for long autorun/polling sessions; both panels already
  have clear buttons.
- [ ] **Body-capture content-type gate** (`harness.js` fetch wrapper): skip
  `clone.text()` for non-text/JSON/XML content types so binary downloads aren't
  buffered just to measure their length. (Streaming prefix read: rejected above.)
- [x] **Old vendor-builder cleanup**: resolved by the 2026-07-26 Monaco
  migration; `tools/build-vendor.mjs` and its temporary entry were removed.
- [ ] **`tests/README.md` env-var formats**: document that `DCSPAD_URL` is a
  full URL including `/index.html` while `DCSPAD_FIXTURES` is a bare origin
  (the suite appends `/fixtures/…`), with an example invocation.
- [ ] **Opportunistic sleep→`waitForFunction` migration** in the pre-existing
  smoke checks, as those sections get touched. New rerun-lifecycle checks
  already follow this pattern.

---

## 2026-07-25 — Gemini Pro 3.1 Extended review (post-fix pass)

Reviewed the code *after* the round-1 fixes landed and did not re-report any of
them — a useful independent signal that those closed cleanly. All four findings
verified accurate; none rejected, a first across these rounds. The disagreements
were about severity and about the shape of one fix.

### Landed (this round)

| # | Issue | Fix |
|---|-------|-----|
| 1 | **In-page anchors destroyed the preview.** A plain `<a href="#foo">` navigated the iframe away, replacing the run | `harness.js` intercepts fragment clicks and re-creates the same-document jump (`location.hash` assignment so `hashchange` fires, then explicit `scrollIntoView`) |
| 2 | Console text filter matched the injected timestamp, so searching `10` hit every line logged at `:10` seconds | Filter now reads `.entry-body` rather than the whole entry |
| 3 | Filtering re-queried `.lvl-filter.active` and the filter input **per entry**, and ran synchronously on every keystroke | Filter state cached and refreshed once per change; input debounced at 150 ms |
| 4 | REPL's `eval()` semantics (`let`/`const` vanish, no top-level `await`) were undocumented and surprising vs DevTools | Placeholder + `title` hint on the REPL input |

Five new smoke checks (34 total): fragment link scrolls rather than leaves,
`hashchange` still fires, user `preventDefault()` still wins, filter ignores
timestamps, filter still matches real content.

### Where the review was wrong — the anchor bug is worse than diagnosed

The review attributed the navigation to the injected `<base href>` and filed it
as a 🟡 suggestion. Both parts are off, and it matters:

- **Cause.** An `about:srcdoc` document resolves relative *and fragment* URLs
  against the **parent document's** base URL. `<base href>` changes the
  destination but is not the cause — the bug reproduces with no `<base>` at all,
  i.e. in the local mock, where `baseHref` is deliberately `null`.
- **Severity.** Measured directly: pre-fix, clicking `<a href="#target">` in the
  preview navigated the iframe to `http://localhost:8642/index.html#target` —
  **DCSPad recursively loaded inside its own preview pane**, run state gone. On a
  tenant it lands on the SP web instead. The trigger is also far more common than
  the review implied: `<a href="#">` as a JS-hook button is idiomatic in
  SharePoint code, and any such link whose handler omits `preventDefault()` hits
  it on the first click.

This was the most consequential finding in the review and was ranked below two
issues that are cosmetic and minor-correctness respectively.

### Where the proposed fixes were changed

- **Fix 1** as written (`preventDefault()` + "optionally scroll") would have
  silently dropped `hashchange`, trading one fidelity break for a subtler one
  that breaks hash-routing user code. We assign `location.hash` so the browser
  performs a real same-document navigation, then scroll explicitly as a fallback
  in case a future engine declines fragment nav on `about:srcdoc`. The listener
  is deliberately **bubble phase** and honours `defaultPrevented`: capture phase
  would have overridden the very common `<a href="#">` + `preventDefault()`
  button pattern and scrolled the preview to the top against the user's code —
  a direct violation of "pad chrome must lose to user code".
- **Fix 3's debounce treats the symptom.** The larger cost was the per-entry
  document queries inside `applyFilterTo`, which also runs on the hot path in
  `addEntry` — a run logging 1,000 lines paid 1,000 redundant `querySelectorAll`
  calls *during the run*. Hoisting the state fixes both paths; the debounce is
  kept as well.

### Severity disagreement (no items rejected)

Neither 🔴 was critical: one is a responsiveness nit already bounded by the
history-cap backlog item, the other a minor correctness bug in a convenience
filter. Ranked here as: anchors (medium) > timestamp match (medium) > filter
perf (low-medium) > REPL hint (low).

### Test limitation, recorded honestly

The three fragment checks are regression guards on the interceptor, run against
the local mock. The `<base href>` variant of the destination can only be
exercised on a real tenant — same class of local limitation as the CDN-egress
note in `tests/README.md`. The underlying navigation *is* reproduced locally,
which is what makes the guards meaningful.

---

## 2026-07-25 — Round 3: GHCP Sonnet re-review + Gemini Pro 3.1 review (save/load build)

Two independent reviews of the same commit (`b000a9c`). All findings verified
against the code; every accepted fix landed together. Notably, the two reviews
**converged from different angles on the same real gap** — storage failures not
surfacing reliably (Sonnet S3, Gemini C1) — which earned that item the most
thorough treatment.

### Landed

| Source | Issue | Fix |
|---|---|---|
| Sonnet S2 | `replaceCatalog` left dead ids in workspace `enabled`/`pinned` forever | Prune to the new catalog's ids inside `replaceCatalog`; regression-tested via the catalog-file round-trip |
| Sonnet S3 + Gemini C1 | Storage failures surfaced unreliably: workspace quota failure emitted no event (status stuck at "saving…"), catalog/snippet errors wrote directly to `#status-save` where the next autosave tick overwrote them | Unified: `persist()` emits `'error'`; status bar renders a three-state (`saving…` / `✓ saved` / `save failed`, styled); catalog + snippets report through an `onStorageError` callback that also writes a persistent console-panel warning |
| Sonnet S1 | `chkModule` referenced before declaration (TDZ trap, not a runtime bug) | Element resolved at call time inside the import handler |
| Sonnet S4 | `newId` counter resets per page load — theoretical same-millisecond cross-session collision | Session-random seed in the id |
| Sonnet S5 | ↑/↓/✕ closures captured `idx` at render time | `liveIdx()` lookup at event time for move **and** delete (the reviewer flagged only move; delete had the identical pattern) |
| Sonnet S6 | No catalog file export/import coverage | Round-trip test through a real download, including the S2 prune assertion |
| Gemini S1 | Synchronous `revokeObjectURL` can abort downloads (Safari-family) | Revocation deferred 1 s |
| Gemini S2 | No size bound on imports — a huge mis-picked file locks the main thread in `JSON.parse` | 5 MB cap (principled: a legitimate DCSPad file can't exceed the localStorage quota it came from) |
| Gemini S3 | Parse errors indistinguishable from wrong-schema files | `wireJsonImport` rejects invalid JSON with its own message; handlers keep shape alerts |

Smoke suite 44 → 49: catalog export shape, import-replaces, import-prunes-dead-ids,
round-trip restore, and quota-failure surfacing (deterministic via a stubbed
`Storage.prototype.setItem` — the only way to hit quota on demand).

### Declined

- **Sonnet S2's second half** (warn "re-enable your frameworks" after catalog
  import). The framing conflates layers: the catalog is *global* ("what's
  installable"), enablement is *per-project* ("what this project uses").
  Everything-unchecked after importing someone else's catalog is the designed
  semantics, not a bug to apologize for; after the prune, state is consistent.

### Corrections to reviewer reasoning (fixes taken anyway)

- **Sonnet S3 / Gemini C1 mechanism.** Sonnet's "the next autosave tick
  overwrites the error with `✓ saved` — a lie" and Gemini's "listener defaults
  to `✓ saved`" are both slightly wrong: on a truly full store the workspace
  write *also* fails, no `'saved'` event fires, and the status sticks at
  "saving…". The lying-checkmark only occurs in the mixed case (small workspace
  write fits, large catalog write doesn't). The fix is the same either way.
- **Sonnet S5 mechanism.** A second physical click is hit-tested against the
  re-rendered DOM at dispatch time, so real users get a *fresh* closure on
  whatever row now sits under the cursor — the stale-closure path is
  essentially automation-only. Taken regardless: the live lookup costs nothing
  and is strictly more correct.
- **Gemini C1 severity.** "Critical — must fix before merge" overstates a
  quota-exhaustion path that leaves a visible (if weak) "saving…" signal;
  ranked here as medium. Best finding in that review all the same.

### Still open (unchanged low-priority backlog from round 1)

Run-scoped load timeout · console/network history caps · fetch body
content-type gate · `tests/README.md` env-var format docs · opportunistic
sleep→`waitForFunction` migration.
