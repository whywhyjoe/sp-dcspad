# State — SP Workbench page status + EEEU audit

Last touched: 2026-09-23
Mode: Joe
Branch: claude/dcspad-sp-utilities-availability-ylr25q
State: PLANNED — decisions below are Joe's answers (2026-09-23); build in progress

## What this is

Two SP Workbench features. The first is a REBUILD: a page-status feature was built once on the
work (bmo) machine and lost — the only surviving record is the original prompt and the plan that
session wrote (no code, no diffs; the answers to that session's questions were also lost). The
second is new: an EEEU / broad-access audit ported (loosely) from Joe's standalone
`DCS.SecurityGroups` script, whose content-audit engine is the reference — its group-membership
viewer is NOT in scope (the Workbench's Permissions view already covers groups/members).

## Decisions — page status (Pages view)

- Scope: modern Site Pages (119 with PromotedState + CanvasContent1) and the classic publishing
  Pages library (850) only. Other libraries don't get the scan/chips/Permissions tab changes.
- "Published" = the page has EVER had a published major version, even with a newer draft on top
  (File/MajorVersion > 0, or UIVersion >= 512). NOT `_ModerationStatus` — that is 0 on every item
  when content approval is off. Two states only: Published / Unpublished.
- Grid: a **Scan Page Status** toolbar button. On click (one extra library query, no per-page
  calls) the grid gains three columns: **Published** (Published/Unpublished), **Checked out**
  (the holder's name), **Inheritance** (marker when broken). The grid shows no date.
  Scan columns ride along in CSV/JSON/markdown exports. Cleared on library switch.
- Detail header: status chips after the kind chip — Published (accent) / Unpublished (neutral);
  Inheritance broken (accent) only when true; Checked out (accent, holder on tooltip) only when
  true. Status register (`design/INFO-CHIP.md`), never the info chip. Chips only when the field
  was readable.
- Detail header buttons (Export MD, Export HTML, Export raw, Open page) move to the tab row,
  right-aligned.
- New detail tab **Permissions**, between Metadata and Structure: a status chip (Inherited /
  Broken inheritance) + the page's own role assignments (the item endpoint returns the inherited
  set when inheritance is intact — never the web's). Principal / Login / Type / Roles, same
  grammar as the Permissions view. 401/403 via denied.js.
- Metadata tab: ONLY these rows, in this order; any field missing from the list is skipped:
  ID · Name (FileLeafRef) · Title · Publish Status (Published/Unpublished) · Publish Date (LAST
  published — date of the current published major version; one versions read, detail only) ·
  First Published Date (FirstPublishedDate) · Permissions (Inherited / Broken inheritance) ·
  Checked Out To (user name) · Promoted State ("Promoted" when 2, "False" otherwise) · Content
  Category (`bmocContentCategory`, managed metadata multi — read-only) · Modified · Modified By ·
  Created · Created By · Item Type (`FolderType`, single choice — display name differs) ·
  Contact (`Contact`, person single — read-only: Workbench can't edit people yet) · Pillar ·
  Org (`Org`, single choice) · Compliance Asset ID (ComplianceAssetId) · Wiki Content
  (WikiField). Scheduled start date (`_PublishStartDate`) is dropped.
  Editable: Name, Title, Item Type, Pillar, Org (plus anything else simple); hard ones
  (managed metadata, person, computed/system) read-only.
- Per-page `.md` export metadata block gains the publish status.

## Decisions — EEEU audit (Permissions view, new tab "EEEU audit")

- Detection, both: (a) a principal list — defaults "Everyone except external users",
  "Everyone", "SharePoint EEEU Visitors" (BMO prod convention), editable in the form, matched on
  title OR login, with EEEU/Everyone also matched by claim (`spo-grid-all-users`, `c:0(.s|true`);
  (b) at scan start, every site group whose members include EEEU/Everyone is added as a target.
- Also flag "People in your organization" sharing links (`SharingLinks.*.OrganizationView*`
  groups).
- Scope: four checkboxes — Site (the web's own role assignments only), Pages libraries, Lists,
  Document libraries. Hidden and system lists skipped.
- Depth option: "Include items and folders" (default on) — off = list/library level only.
- Subsites: the site's direct subsites listed as checkboxes; a checked subsite is scanned with
  its whole tree (sub-subsites are never listed individually).
- Limited Access–only matches hidden (noise from item-level shares).
- Output: one results grid (Type, Site, Location, Name, Shared with, Permission, Unique, Link)
  + a Problems grid; grid exports. No "folder containing matching files" summary rows.
- Speed: bounded concurrency (4) + sp-rest.js's 429/503 retry; Cancel keeps partial results
  (no 10-minute timeout).

## Next

- [ ] Page status: pure helpers + mock fixtures + pages.js (grid scan, chips, tab-row actions,
      Permissions tab, Metadata spec) + CSS + page-export metadata + tests
- [ ] EEEU audit: pure engine module + security.js tab + mock fixtures + tests
- [ ] CLAUDE.md file map + test counts; rebuild dcspad.workbench.js from a clean tree
- [ ] Live-tenant check on bmo: internal names (bmocContentCategory, FolderType, Org, Pillar,
      Contact), publish state on a page with a draft over a published version, EEEU audit on a
      site with a known EEEU grant
