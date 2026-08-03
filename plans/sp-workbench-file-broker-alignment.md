# SP Workbench file-broker alignment plan

Status: proposed; no broker integration is implemented by this review.

## Recommendation

Adopt the file broker selectively as the shared SharePoint write and metadata
engine, while keeping the SP Workbench Files view as its current compact
browser. This is worth doing only as a coordinated consolidation of the
duplicate upload/metadata paths in SP Workbench and DCSPad. Replacing just the
Workbench upload call would add a dependency without retiring enough code.

The broker already provides the material improvements that justify this work:

- one provider-owned implementation of digest handling, path validation,
  conflict behavior, binary upload, and post-upload metadata updates;
- discovery-mode metadata forms instead of a second SharePoint-field editor;
- a two-step upload contract whose metadata retry never uploads the bytes again;
- the document-library-root `GetList(@listUrl)` fallback now needed by the
  Workbench; and
- a neutral UI/provider boundary that can also serve DCSPad's file dialogs.

## Scope and non-goals

- Keep the Files rail entry, grid, breadcrumbs, details pane, and routing.
- Do not embed the broker's full picker dialog inside the Files view.
- Do not import source directly from a sibling checkout at runtime. Consume a
  pinned release or vendor a reviewed build artifact.
- Do not change list/folder browsing until the write-path migration is proven.
- Do not raise upload size limits as part of this work; chunked uploads are a
  separate feature.

## Phased implementation

### 0. Package and compatibility contract

1. Publish or vendor an immutable broker build with its version recorded.
2. Define the adapter inputs available from Workbench: target `webUrl`, fetch,
   current folder, file bytes, overwrite consent, and status callbacks.
3. Confirm classic SharePoint compatibility and that the build adds no module
   or CSP requirement the hosted page cannot satisfy.

Exit: the broker can be instantiated for the currently selected Workbench site
without reaching into broker internals.

### 1. Workbench upload engine

1. Route binary upload, conflict handling, digest refresh, and metadata update
   through the broker's headless SharePoint provider.
2. Preserve the current Workbench consent UI and capture the destination folder
   when an upload begins.
3. Map broker progress/errors into the existing notice and error washes.
4. Remove the superseded Workbench write helpers only after parity tests pass.

Exit: upload and overwrite behavior is unchanged to users, metadata failures can
be retried without re-uploading, and only the broker owns SharePoint write rules.

### 2. Shared metadata form

1. Replace the Workbench-specific file metadata editor with
   `createMetadataForm({ mode: 'discover' })`.
2. Preserve Workbench detail-pane placement, Close behavior, and compact visual
   treatment through the broker's neutral rendering hooks/tokens.
3. Verify text, multiline, number, date/time, boolean, choice, multi-choice,
   URL, user, lookup, and taxonomy support before deleting the old editor.

Exit: one field-discovery and serialization implementation serves both repos.

### 3. DCSPad consolidation

1. Migrate the duplicate open/import/export picker and upload paths to the same
   broker release.
2. Retain DCSPad-specific orchestration and document semantics outside the
   broker adapter.
3. Delete duplicate picker/write code only after old and new flows pass the
   same hosted-SharePoint checks.

This phase supplies most of the maintenance payoff; without it, stop after
phase 0 and keep the current Workbench implementation.

## Acceptance gates

- Broker unit/integration suite passes unchanged.
- `tests/workbench.mjs` and `tests/workbench-edit.mjs` pass.
- Live tenant: root-library and nested-folder upload, overwrite decline/accept,
  digest expiry, metadata save failure/retry, unusual names containing `#`,
  `%`, quotes, and Unicode, and a non-admin contributor account.
- No bytes are posted during a metadata-only retry.
- Site switching cannot send an upload or metadata update to the previous site.
- The hosted Workbench bundle has no unresolved cross-repo imports.

## Rollback

Keep the existing Workbench adapter behind one internal boundary for the first
release. A single build-time flag may select it during rollout; remove that
fallback after one release and successful live-tenant verification.
