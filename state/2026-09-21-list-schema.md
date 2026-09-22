# State — SP Workbench List Schema

Last touched: 2026-09-21
Mode: Joe
Branch: claude/workbench-list-schema
State: stage 1a code and tests are done and green in mock mode; nothing has run against a live tenant yet

## What this is

The SP Workbench could inspect a list but not reproduce one. Stage 1a ports the schema half
of the console-first SPUtils list-copy primitives (`utilities/dcspad-sp-utilities.js`,
`getListSchema`/`createListFromSchema`, verified live on the dev tenant) into the Workbench UI
as an assisted flow: a Schema tab that captures and exports a list's settings/fields/views/
content types, and a Copy to… / New from schema… dialog that applies that document to a target
list on the same or another same-tenant web. Plan record:
`~/.claude/plans/plan-to-add-a-idempotent-rossum.md`. Full design detail (capture reads, the
apply step plan, the dialog, mock fixtures, test layout) lives there — this file tracks status,
not design.

Stages: **1a** schema capture/export/import/create/copy (this work) · **1b** item data on the
same engine · **2** document libraries (schema only; no file transfer).

## Done

- `list-schema.js` (pure core: doc normalize/scrub, `buildApplyPlan`, `buildApplyReport`,
  `retryPlan`), `list-schema-capture.js`, `list-schema-script.js` (PnP.PowerShell / PnPjs 2 /
  SPUtils-call emitters from the same plan), `list-schema-apply.js` (executor), and
  `list-schema-dialog.js` (Copy to… / New from schema… UI) — all under `src/workbench/`.
- Plumbing: `grid.createMenuButton` lifted out for reuse, `sp-write.js` gained header/mergeJson
  support and a 429/503 retry, `main.js` grew a `createClient` shell dep so the dialog can hold
  a second (target-web) REST connection without disturbing the shell client, `mock-data.js`
  grew two new mock webs (`/sites/schema` source, `/sites/target` apply target) plus a stateful
  mock writer.
- `views/lists.js`: Schema tab (after Content types), `New from schema…` on the grid toolbar,
  route write-back of the active tab.
- New suite `tests/workbench-schema.mjs`: 78 checks (pure planner/scrub/emitters, Schema tab
  mock UI, executor mock + stubbed live, dialog mock UI, live 401). `workbench.mjs` stays 133
  (only its tab-order pin gained Schema). All suites together: 513 checks, green.
- Docs updated: `CLAUDE.md` (file map, test paragraph, roadmap), `tests/README.md` (suite line,
  two-web note), `HANDOFF.md` (`## SP Workbench: List schema (stage 1a)` section — the doc
  contract, the SPUtils gaps closed, the second-client pattern, hooks for 1b/2, the live-tenant
  checklist below repeated there), `utilities/README.md` (Copying a list section names the
  assisted vs. console path and the two SPUtils follow-ups).
- `dcspad.workbench.js` rebuilt from clean HEAD (commit 2d96bc3).

## Next

- [ ] Live dev-tenant validation — see the checklist below (also recorded in HANDOFF.md).
      `deploy/Sync-Live.ps1 -Environment dev`, then walk the Workbench against a real list.
- [ ] Stage 1b: item data on the same engine (`list-data.js`, Items fieldset currently
      `disabled` in the dialog, `AddValidateUpdateItemUsingPath`, lookup re-resolution, date
      calibration — see the plan's Stage 1b section for the detailed design).
- [ ] Stage 2: document libraries, schema only (`baseType === 1` detection, `BaseTemplate:101`
      create, library-specific settings groups, `Copy to…` currently disabled on libraries with
      a stage-2 title).

## Open questions

- None outstanding from Joe as of 2026-09-21; the two SPUtils gaps below are follow-ups against
  `utilities/dcspad-sp-utilities.js`, not open design questions for the Workbench port itself.

## Landmines

- The mock writer must return the ids the executor binds to (`web/lists` → `{Id, Title,
  RootFolder}`, `createfieldasxml` → `{Id, InternalName}`, views add → `{Id}`) — the probe step
  after create resolves through the same registry, so a writer that doesn't register a write
  breaks the next read in the same run.
- `sp-write.post` merges caller headers over its base set (digest, Accept, content-type) rather
  than the other way round — a step that needs `X-HTTP-Method: MERGE` or `IF-MATCH: *` must pass
  them explicitly; the base set alone never supplies them.
- `tests/workbench.mjs:140` pins the Lists drill-down tab order; Schema must stay in the pinned
  position (after Content types) or that check breaks for reasons unrelated to Schema itself.
- `buildApplyPlan`'s `dependsOn`: views and the validation-formula step depend on the `list`
  step only, not on every field. One field that can't be recreated (e.g. managed metadata) must
  not block every view or the validation formula from being applied.
- Plan-time refusals the probe makes (type clash, taken title, missing content type) carry
  `final: true` and are never retried by "Retry failed steps" — only steps that failed during
  execution are re-run.
- The build must run from `tools/` under the system Node (`C:\Program Files\nodejs\node.exe`,
  ARM64) — see the CLAUDE.md Gotchas entry on the two-Node-installs machine; an x64 esbuild
  binary throws "installed esbuild for another platform" partway through `Sync-Live.ps1`.

## Live-tenant checklist (copied from the plan's Verification section; open, unfilled)

- nometadata `POST web/lists` accepted; `createfieldasxml` `{parameters:{SchemaXml, Options}}`
  in nometadata and the `Options 8|4` behaviour on a content-types-enabled list.
- `views?$expand=ViewFields`; `ViewTypeKind` in the views POST; `RowLimit`/`Paged` in one MERGE.
- Field MERGE of `Indexed`/`EnforceUniqueValues`/`CustomFormatter` on lookup/choice types.
- `addAvailableContentType` posted before field creation.
- A User field round-trips with `List="UserInfo"`; a FullHtml Note round-trips.
- Property availability of `EnableRequestSignOff`/`ListExperienceOptions`/`DisableGridEditing`
  on this tenant.
- A cross-web copy into a subsite.
- The 429 retry path under real throttling.
- Whether SPUtils' string-overload `createFieldAsXml` regenerates internal names.
- Export a SPUtils v1 doc and import it in the Workbench, and the reverse through
  `SPUtils.createListFromSchema`.
