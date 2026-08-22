// Pages view: the page inspector. Master grid over the web's pages library —
// modern Site Pages (BaseTemplate 119), classic publishing Pages (850), or
// any library titled "Pages" — with the resolved library named in the header
// so it is always clear which one is on screen.
//
// The drilldown adapts to the page, not the library, because one 119 library
// can hold both modern and wiki pages (see classic-page.js):
//   canvas page     CanvasContent1 -> structure tree, web-part inventory,
//                   extracted content
//   classic page    PublishingPageContent/WikiField PLUS the Content Editor
//                   and Script Editor web parts, merged into one extract;
//                   no Structure tab, since sections/columns do not exist
// Both kinds also get the editable metadata sheet and the raw entity.

import { createGrid, bindNewTab } from '../grid.js?v=2';
import { copyText } from '../export.js';
import {
  parseCanvasContent, buildSectionTree, textOfControl, webPartName, sanitizeHtml,
} from '../canvas.js';
import { createSpWriteClient } from '../sp-write.js';
import { createFieldEditorForm } from '../field-editor.js';
import {
  buildContentExport, buildRawExport, exportFileStem, contentParts,
} from '../page-export.js';
import { downloadText } from '../../io.js?v=2';
import { enhance } from '../../inspect/sp-shapes.js';
import { renderValue } from '../../inspect/tree-view.js';
import { odataPathLiteral } from '../../sp-odata.js';
import { toNode } from '../../inspect/to-node.js';
import { showFailure } from '../denied.js';
import {
  libraryKindOf, libraryKindLabel, pageContentKindOf, pageContentKindLabel,
  classicWebParts, classicContentParts,
} from '../classic-page.js';

// PromotedState is modern-only — selecting it against a library that lacks it
// 400s with "The field or property 'PromotedState' does not exist".
//
// Title is NOT universal either, which cost a live degrade to learn. The
// PointPublishing hub/community webs and a Project Web App site all carry a
// BaseTemplate 119 "Site Pages" library with no Title field at all, so
// `$select=Id,Title` 400s on them exactly like PromotedState does. Every rung
// that named Title was refused, the ladder fell to the bare bottom rung, and
// the view reported the loss as "the Editor column" — true, but only because
// the bottom rung drops the expand; Editor itself was never the problem. So
// Title is probed like PromotedState rather than assumed.
//
// Id stays unconditional: it is an intrinsic OData property, not a list
// field, and does NOT appear in /fields — probing for it would drop the one
// column the drilldown routes on.
const PAGE_SELECT_BASE = [
  'Id', 'Title', 'FileLeafRef', 'FileRef', 'FileDirRef',
  'Modified', 'UniqueId', 'Editor/Title',
];
const PAGE_SELECT_MODERN = [...PAGE_SELECT_BASE, 'PromotedState'];

// BaseTemplate is NOT a schema guarantee: 119 is the Wiki Page Library
// template, and a classic team site that never got the modern Site Pages
// feature has a 119 library WITHOUT PromotedState/CanvasContent1. So the
// query shapes are decided by probing the library's actual fields (one
// request, shared with the metadata pane), and the library kind is only the
// fallback when that probe fails.
export function pageQueryPlan(fieldInternalNames, kind) {
  const names = fieldInternalNames ? new Set(fieldInternalNames) : null;
  const hasField = (f) => (names ? names.has(f) : kind === 'modern');
  const showPromoted = hasField('PromotedState');
  // Unprobed (the fields request itself failed) keeps Title: it is present on
  // most libraries, and the ladder still covers the ones it is not.
  const showTitle = names ? names.has('Title') : true;
  // PromotedState + CanvasContent1 together mark the modern Site Pages
  // feature, which also provisions the rest of DETAIL_SELECT.
  const modern = showPromoted && hasField('CanvasContent1');
  const gridSelect = (showPromoted ? PAGE_SELECT_MODERN : PAGE_SELECT_BASE)
    .filter((f) => f !== 'Title' || showTitle);
  return {
    showPromoted,
    showTitle,
    gridSelect,
    // Ladders, not single shapes — see queryLadder(). Rung 0 is the query we
    // want; every rung below it gives something up to stay answerable.
    gridShapes: [
      { options: { select: gridSelect, expand: 'Editor' } },
      // No lookup projection, so no expand to satisfy: the Editor column
      // goes blank and everything else still lists.
      {
        options: { select: gridSelect.filter((f) => !f.includes('/')) },
        lost: 'the Editor column',
      },
      // Nothing named at all. SPO returns the item's own fields, which is
      // every column this grid reads except the expanded Editor.
      { options: {}, lost: 'the Editor column' },
    ],
    detailShapes: [
      {
        options: modern
          ? { select: DETAIL_SELECT, expand: DETAIL_EXPAND }
          : { select: CLASSIC_DETAIL_SELECT, expand: DETAIL_EXPAND },
      },
      // '*' still carries every content field the drilldown reads
      // (CanvasContent1, PublishingPageContent, WikiField); only the two
      // expanded people fields are out of reach, leaving their raw ids.
      { options: { select: ['*'] }, lost: 'the author and editor names' },
      { options: {}, lost: 'the author and editor names' },
    ],
  };
}

// SharePoint answers a query it cannot SHAPE with 400 — a field this library
// does not have ("The field or property 'X' does not exist"), or an expand
// whose target the select fails to name ("The query to field 'Author' is not
// valid…"). None of that says the page is unreachable, so a rejected shape
// steps down to a simpler one instead of dead-ending the view: better a grid
// with a blank column, or a page without its author, than a red bar over
// content the operator can plainly see exists.
//
// 400 only. A 403 (no rights), 404 (gone) or 429 (throttled) is about the
// resource, not the query, and no rung of the ladder would fix it — those
// stay loud, and stay fast.
//
// Two things the ladder owes the caller, because a degrade that explains
// nothing is indistinguishable from a first-party bug:
//
//   `reason`  SharePoint's sentence from the rung that was REJECTED, not the
//             one that answered. When the step-down was caused by a typo in
//             our own select, "The field or property 'Titel' does not exist"
//             is the whole diagnosis — and it is the sentence the chip shows.
//   `index`   which rung answered, so a caller looping over many items can
//             start there instead of re-earning the same 400 per item. A
//             shape rejection is a property of the LIST's schema; paying for
//             it once per page turned one bad shape into one failed request
//             per click.
//
// If every rung fails, the FIRST error is thrown, not the last: rung 0 names
// the field or projection SharePoint actually objected to, while the bare
// fallback at the bottom of the ladder can only ever say "Bad Request".
export async function queryLadder(shapes, attempt, startAt = 0) {
  let firstError;
  const from = Math.min(Math.max(startAt, 0), shapes.length - 1);
  for (let index = from; index < shapes.length; index += 1) {
    const shape = shapes[index];
    try {
      return {
        value: await attempt(shape.options),
        lost: shape.lost || '',
        index,
        reason: firstError?.message || '',
      };
    } catch (err) {
      if (err?.status !== 400) throw err;
      firstError ||= err;
    }
  }
  throw firstError;
}

// $expand of a User field is only legal alongside a $select that NAMES the
// expanded target field. Expanding Author/Editor with no $select at all 400s
// on a classic library with "The query to field 'Author' is not valid. The
// $select query string must specify the target fields and the $expand query
// string must contains Author." (Modern Site Pages happened to escape it
// only because DETAIL_SELECT already spells Author/Title out.) Both shapes
// therefore ship their own select; the grid query above pairs
// Editor/Title with $expand=Editor for the same reason.
const DETAIL_EXPAND = ['Author', 'Editor'];

// Modern-only: every field here exists on a modern Site Pages item and
// nowhere else.
const DETAIL_SELECT = [
  'Id', 'Title', 'FileLeafRef', 'FileRef', 'FileDirRef', 'Description',
  'BannerImageUrl', 'PromotedState', 'Created', 'Modified',
  'Author/Title', 'Editor/Title', 'CanvasContent1', 'LayoutWebpartsContent',
];

// Every other library: the field set varies per site, and asking for a field
// a publishing/wiki schema lacks is a 400, so the whole item is taken with
// '*' — the only way to reach PublishingPageContent/WikiField without
// probing — plus the two expand projections. Same '*'-with-expand shape the
// Items tab pays for in lists.js.
const CLASSIC_DETAIL_SELECT = ['*', 'Author/Title', 'Editor/Title'];

const FIELD_SELECT = [
  'Id', 'Title', 'InternalName', 'TypeAsString', 'FieldTypeKind', 'Required',
  'Hidden', 'ReadOnlyField', 'Group', 'DefaultValue', 'Choices', 'Description',
  'FillInChoice',
];

const SITE_PAGES_BASE_TEMPLATE = 119;
const PUBLISHING_PAGES_BASE_TEMPLATE = 850;

// Locate the web's pages libraries: modern Site Pages (119) first, then the
// classic publishing "Pages" library (850), then any visible library that
// is simply titled Pages — older sites use all three shapes.
//
// A web can hold MORE THAN ONE of these at once, and routinely does: any
// classic publishing site that has ever had a modern page added carries both
// an 850 "Pages" library (where the real content is) and a 119 "Site Pages"
// one (often near-empty). Resolving to a single winner and discarding the
// rest made the other library unreachable from this view and gave no hint it
// existed — the Pages tab just looked empty. So the ranking still decides
// what opens by default, but every candidate is kept and offered in the
// picker.
// Visibility outranks template. The ladder used to read
// 119-visible, 119-hidden, 850-visible, 850-hidden, which preferred a HIDDEN
// Site Pages library over a VISIBLE publishing one — backwards, since a
// hidden library is the more likely vestigial of the two. Within each
// visibility tier the modern template still wins.
// (A hidden library merely *titled* "Pages" is still not a candidate at all;
// at that point it is almost certainly something internal.)
const PAGES_LIBRARY_RANKS = [
  (l) => !l.Hidden && l.BaseTemplate === SITE_PAGES_BASE_TEMPLATE,
  (l) => !l.Hidden && l.BaseTemplate === PUBLISHING_PAGES_BASE_TEMPLATE,
  (l) => !l.Hidden && String(l.Title).toLowerCase() === 'pages',
  (l) => l.BaseTemplate === SITE_PAGES_BASE_TEMPLATE,
  (l) => l.BaseTemplate === PUBLISHING_PAGES_BASE_TEMPLATE,
];

// Every pages library in the web, best-first by the ranking above. Deduped by
// Id, so a library matching two ranks (a visible 119 matches both the first
// and second) appears once, at its best rank.
export function pagesLibraryCandidates(items) {
  const seen = new Set();
  const found = [];
  for (const matches of PAGES_LIBRARY_RANKS) {
    for (const list of items || []) {
      if (seen.has(list.Id) || !matches(list)) continue;
      seen.add(list.Id);
      found.push(list);
    }
  }
  return found;
}

// The default: the best-ranked candidate. Unchanged semantics — this is still
// exactly what the old ladder returned.
export function pickPagesLibrary(items) {
  return pagesLibraryCandidates(items)[0] || null;
}

const promotedLabel = (v) => ({ 0: '', 1: 'News (pending)', 2: 'News' }[v] ?? String(v ?? ''));
const fmtDate = (v) => (v ? String(v).slice(0, 10) : '');

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

const encodedServerPath = (path) => String(path || '').split('/').map((segment) => {
  try { return encodeURIComponent(decodeURIComponent(segment)); }
  catch { return encodeURIComponent(segment); }
}).join('/');

const guidPath = (listId, sub = '') => `web/lists(guid'${listId}')${sub}`;

// What a degraded query cost, said quietly. This is the info register, not
// the status one (design/INFO-CHIP.md): nothing here is a state to act on —
// it classifies what the view is able to show — so it composes .wb-info-chip
// like the library-kind and page-kind chips, carries a phrase in sentence
// case, and puts the sentence on the tooltip.
//
// `because` is the server's sentence from the shape it turned down. It is the
// difference between a chip that hides a bug and one that reports it: the
// rejection is usually SharePoint declining a field the schema lacks, but it
// reads identically to our own malformed select, and only the server's words
// tell the two apart. The chip used to end "Everything else on this page is
// complete" — a promise it cannot keep, since the bottom rung asks for no
// $select at all and SPO may then withhold expensive fields of its own accord.
function reducedChip(lost, where, because = '') {
  const chip = el('span', 'wb-info-chip wb-reduced-chip', 'some fields unavailable');
  chip.title = `SharePoint rejected part of this query${where ? ` for ${where}` : ''}, `
    + `so ${lost} could not be read.`
    + (because ? `\n\nSharePoint said: ${because}` : '');
  return chip;
}

export function createPagesView({ client, navigate, updateRoute }) {
  const root = el('section', 'wb-view wb-view-pages');
  const spWrite = createSpWriteClient({ client });

  // ---- master pane ----
  const gridPane = el('div', 'wb-pane');
  const head = el('div', 'wb-view-head');
  // The hint stays generic and static, matching every other view's head —
  // the resolved-library facts live on the status strip, not in prose.
  head.innerHTML = '<h2>Pages</h2>'
    + '<p class="wb-view-hint">Every page in this web’s pages library, '
    + 'subfolders included. Click a row to inspect content, metadata, and structure.</p>';
  // Resolved-library status strip. Modern and classic libraries are both
  // supported and the view must say which one is on screen — as tool tokens
  // (name · kind badge · open link), never prose; sentences ride on the
  // tooltips. Born in the head as a loading placeholder, then adopted into
  // the grid toolbar (toolbarExtras) so status shares the row with the
  // filter and actions.
  const strip = el('div', 'wb-lib-strip');
  strip.append(el('span', 'wb-lib-wait', 'locating library…'));
  head.append(strip);
  const libraryLink = el('a', 'btn btn-xs wb-head-link', 'Open ↗');
  bindNewTab(libraryLink);
  libraryLink.hidden = true;
  strip.append(libraryLink);

  // The chip carries libraryKindLabel()'s full phrase — the same idiom the
  // page-kind chip takes from pageContentKindLabel(), and one source of
  // truth for these words. The BaseTemplate number lives on the tooltip and,
  // when there is a choice to make, in the picker's options. It is an info
  // chip, not a status chip: which library shape this is classifies the view,
  // it does not report a condition, so no kind is coloured.

  // One library: a click-to-copy name token. More than one: a picker in its
  // place, so the libraries the ranking did not choose are reachable instead
  // of invisible. The kind chip and Open link always describe the selection.
  function renderLibraryStrip() {
    if (!current) { strip.hidden = true; return; }
    strip.hidden = false;
    strip.textContent = '';

    if (libraries.length > 1) {
      const select = el('select', 'wb-lib-select wb-lib-picker');
      select.setAttribute('aria-label', 'Pages library to inspect');
      select.title = `This web has ${libraries.length} pages libraries — pick which one to inspect.`;
      for (const lib of libraries) {
        // Titles collide across shapes often enough (a doc library literally
        // named "Pages" beside Site Pages), so the template disambiguates.
        const opt = el('option', '', `${lib.title} · ${lib.baseTemplate}${lib.hidden ? ' · hidden' : ''}`);
        opt.value = lib.listId;
        if (lib.listId === current.listId) opt.selected = true;
        select.append(opt);
      }
      select.addEventListener('change', () => {
        switchLibrary(libraries.find((l) => l.listId === select.value));
      });
      strip.append(select);
    } else {
      const name = el('span', 'wb-lib-name sp-copy', current.title);
      if (current.rootPath) {
        name.title = `Click to copy the library path\n${current.rootPath}`;
        name.addEventListener('click', () => copyText(current.rootPath, name));
      }
      strip.append(name);
    }

    const kind = el('span', `wb-info-chip wb-lib-kind wb-lib-${current.kind}`,
      libraryKindLabel(current.kind));
    kind.title = `BaseTemplate ${current.baseTemplate}`
      + (current.hidden ? ' · hidden library' : '')
      + (current.rootPath ? `\n${current.rootPath}` : '');
    strip.append(kind);

    if (current.viewUrl) {
      libraryLink.href = current.viewUrl;
      libraryLink.title = `Open ${current.title} in a new tab`;
      libraryLink.hidden = false;
    } else {
      libraryLink.hidden = true;
    }
    strip.append(libraryLink);
  }
  const masterStatus = el('div', 'wb-grid-status');
  masterStatus.hidden = true;
  gridPane.append(head, masterStatus);

  const detailPane = el('div', 'wb-pane');
  detailPane.hidden = true;
  root.append(gridPane, detailPane);

  let librariesPromise = null;   // -> library[] (best-first, possibly empty)
  let libraries = [];
  let current = null;            // the library on screen
  let grid = null;
  let pagesLoaded = false;
  // Keyed `${listId}:${pageId}`, never the bare page id: two libraries in one
  // web both start at item 1, and a rejection handler firing after a switch
  // would otherwise evict the OTHER library's entry.
  const detailCache = new Map();   // "listId:pageId" -> Promise<item>
  // The rung the detail ladder settled on for THIS library. A rejected shape
  // is a fact about the list's schema, not about item 7, so the first page
  // that steps down spares every page after it the same failing request.
  // Reset with the rest of the per-library state in adoptLibrary().
  let detailRung = 0;
  let fieldsPromise = null;        // list fields shared by every page
  let detailRun = 0;
  let loadRun = 0;                 // generation guard for loadPages

  const toLibrary = (list) => ({
    listId: list.Id,
    title: list.Title,
    kind: libraryKindOf(list),
    baseTemplate: list.BaseTemplate,
    hidden: Boolean(list.Hidden),
    rootPath: list.RootFolder?.ServerRelativeUrl || '',
    viewUrl: list.DefaultViewUrl || list.RootFolder?.ServerRelativeUrl || '',
  });

  function pagesLibraries() {
    if (!librariesPromise) {
      librariesPromise = client.getAll('web/lists', {
        select: ['Id', 'Title', 'BaseTemplate', 'Hidden', 'DefaultViewUrl', 'RootFolder/ServerRelativeUrl'],
        expand: 'RootFolder',
        top: 5000,
      }).then(({ items }) => {
        // Client-side filter: the mock resolver ignores $filter, and the
        // libraries are cheap to find in the full list either way.
        libraries = pagesLibraryCandidates(items).map(toLibrary);
        if (!current) current = libraries[0] || null;
        return libraries;
      }).catch((err) => {
        librariesPromise = null;
        throw err;
      });
    }
    return librariesPromise;
  }

  // Everything above is scoped to ONE library and must go when it changes.
  function adoptLibrary(next) {
    current = next;
    detailCache.clear();
    webPartCache.clear();
    fieldsPromise = null;
    // The query plan is PROBED from the old library's fields. Carrying it over
    // would compose a request naming fields the new library does not have —
    // PromotedState against a publishing library is a 400 on live SPO, the
    // exact failure the probe exists to prevent.
    planPromise = null;
    detailRung = 0;
    if (grid) { grid.el.remove(); grid = null; }
    pagesLoaded = false;
  }

  // The picker. Records the choice in the route as well, so a reload lands on
  // the library the user was actually looking at rather than the ranked
  // default (which would resolve a saved page id against the wrong list).
  function switchLibrary(next) {
    if (!next || next.listId === current?.listId) return;
    adoptLibrary(next);
    // navigate() lands back on the list view, which reloads it. Detaching the
    // grid took the strip with it (the toolbar owns it once adopted);
    // loadPages re-adopts the same node into the rebuilt toolbar.
    navigate({ view: 'pages', libId: next.listId });
  }

  // Web identity for exports (site display name + absolute URL base).
  let webInfoPromise = null;
  function webIdentity() {
    if (!webInfoPromise) {
      webInfoPromise = client.get('web', { select: ['Title', 'Url'] })
        .catch(() => ({ Title: '', Url: client.webUrl() }));
    }
    return webInfoPromise;
  }

  // Folder of a page relative to the library root ('' at the root).
  function folderOf(fileDirRef, rootPath) {
    const dir = String(fileDirRef || '');
    const root = String(rootPath || '').replace(/\/+$/, '');
    if (!root || !dir.toLowerCase().startsWith(root.toLowerCase())) return '';
    return dir.slice(root.length).replace(/^\/+/, '');
  }

  async function loadPages() {
    if (pagesLoaded) return;
    const run = ++loadRun;
    masterStatus.hidden = true;
    try {
      await pagesLibraries();
      if (!current) {
        strip.hidden = true;
        masterStatus.textContent = 'This web has no pages library — looked for modern '
          + 'Site Pages (BaseTemplate 119), classic publishing Pages (850), and any '
          + 'library titled “Pages”.';
        masterStatus.hidden = false;
        return;
      }
      const sitePages = current;
      renderLibraryStrip();
      if (!grid) {
        const plan = await queryPlan(sitePages);
        // Ordering and the cap are not part of what SPO can reject on schema
        // grounds, so they ride on every rung.
        const paging = { orderby: 'FileLeafRef', top: 5000 };
        const query = { path: guidPath(sitePages.listId, '/items') };
        const descriptor = {
          ...query,
          options: { ...plan.gridShapes[0].options, ...paging },
          webUrl: client.webUrl(),
        };
        grid = createGrid({
          columns: [
            { key: 'FileLeafRef', label: 'Name', mono: true },
            // Dropped with the field itself: a library without Title would
            // otherwise carry a column that can only ever be blank.
            ...(plan.showTitle ? [{ key: 'Title', label: 'Title' }] : []),
            {
              key: 'Folder',
              label: 'Folder',
              value: (row) => folderOf(row.FileDirRef, sitePages.rootPath),
              format: (v) => (v ? `/${v}` : ''),
            },
            ...(plan.showPromoted
              ? [{ key: 'PromotedState', label: 'Promoted', format: promotedLabel }]
              : []),
            { key: 'Modified', label: 'Modified', format: fmtDate },
            { key: 'Editor', label: 'Editor', value: (row) => row.Editor?.Title || '' },
            {
              key: 'FileRef',
              label: '',
              format: () => '',
              render: (fileRef) => {
                if (!fileRef) return null;
                const a = document.createElement('a');
                a.className = 'wb-cell-link';
                a.href = fileRef;
                a.title = 'Open the page in a new tab';
                a.textContent = '↗';
                bindNewTab(a);
                return a;
              },
            },
          ],
          onOpen: (row) => navigate({
            view: 'pages',
            pageId: row.Id,
            pageName: row.FileLeafRef || row.Title,
            // Without this a reload resolves the saved id against the ranked
            // default library — same id, different page, and the Metadata tab
            // would then write to the wrong item.
            libId: sitePages.listId,
          }),
          emptyText: `No pages in ${sitePages.title}.`,
          subject: `the pages in ${sitePages.title}`,
          filterPlaceholder: 'Filter pages…',
          toolbarExtras: strip,
          exportName: 'sp-pages',
          // The same object the ladder rewrites below, on purpose: the
          // "Copy as…" menu reads it at click time, so a script copied out of
          // a degraded grid reproduces the query that actually worked rather
          // than the one SharePoint rejected.
          descriptor,
        });
        gridPane.append(grid.el);
        grid.setLoading('Loading pages…');
        const { value, lost, reason } = await queryLadder(plan.gridShapes, (options) => {
          descriptor.options = { ...options, ...paging };
          return client.getAll(query.path, descriptor.options);
        });
        const { items, partial } = value;
        // A switch during the request replaced `grid`; without this the old
        // library's rows land in the new library's table under the new chip.
        if (run !== loadRun) return;
        if (lost) strip.insertBefore(reducedChip(lost, sitePages.title, reason), libraryLink);
        grid.setRows(items, { partial });
        pagesLoaded = true;
      }
    } catch (err) {
      // Same guard on the failure path: a stale rejection must not paint an
      // error over the grid that replaced it.
      if (run !== loadRun) return;
      // Resolution itself failed: don't leave the "locating…" token up
      // next to the error. A resolved strip stays — it is still true.
      if (strip.querySelector('.wb-lib-wait')) strip.hidden = true;
      if (grid) grid.setError(err);
      else showFailure(masterStatus, err, 'this web’s pages');
    }
  }

  // The probed query plan for the resolved library (see pageQueryPlan). A
  // failed fields probe falls back to the BaseTemplate heuristic rather than
  // blocking the view.
  let planPromise = null;
  function queryPlan(sitePages) {
    if (!planPromise) {
      planPromise = listFields(sitePages.listId)
        .then((fields) => pageQueryPlan(fields.map((f) => f.InternalName), sitePages.kind))
        .catch(() => pageQueryPlan(null, sitePages.kind));
    }
    return planPromise;
  }

  // ---- drilldown ----

  // Resolves to { item, lost, reason } — `lost` naming what a stepped-down
  // query gave up ('' when the wanted shape worked) and `reason` carrying
  // SharePoint's own words for why, straight onto the chip. The Metadata tab fetches its own
  // copy with FieldValuesAsText, so a lost people field still has a readable
  // value there; only this pane's header and the content export go without.
  function pageItem(listId, pageId, shapes) {
    // Library-qualified: both libraries of a web start at item 1, and a
    // rejection arriving after a switch used to evict the other library's
    // entry by bare id.
    const key = `${listId}:${pageId}`;
    const path = guidPath(listId, `/items(${pageId})`);
    if (!detailCache.has(key)) {
      detailCache.set(
        key,
        queryLadder(shapes, (options) => client.get(path, options), detailRung)
          .then(({ value, lost, index, reason }) => {
            detailRung = index;
            return { item: value, lost, reason };
          })
          .catch((err) => {
            detailCache.delete(key);
            throw err;
          }),
      );
    }
    return detailCache.get(key);
  }

  // Classic pages keep most of their real content in Content Editor / Script
  // Editor web parts rather than in an item field, so the reading model has
  // to go through the page's shared web-part manager. Best effort by design:
  // the endpoint 403s without ManageWebParts on some webs and is meaningless
  // on modern pages, and neither case should break the drilldown.
  const webPartCache = new Map();   // fileRef -> Promise<{ parts, error }>
  function classicWebPartsOf(fileRef) {
    const key = String(fileRef || '');
    if (!key) return Promise.resolve({ parts: [], error: null });
    if (!webPartCache.has(key)) {
      const path = `web/getfilebyserverrelativepath(decodedurl='${odataPathLiteral(key)}')`
        + '/getlimitedwebpartmanager(scope=1)/webparts';
      webPartCache.set(key, client.getAll(path, { expand: 'WebPart/Properties' })
        .then(({ items }) => ({ parts: classicWebParts(items), error: null }))
        .catch((err) => {
          // Evict so a transient failure (429, network) retries on the next
          // visit; the current caller still gets the error to surface.
          webPartCache.delete(key);
          return { parts: [], error: err };
        }));
    }
    return webPartCache.get(key);
  }

  function listFields(listId) {
    if (!fieldsPromise) {
      fieldsPromise = client.getAll(guidPath(listId, '/fields'), { select: FIELD_SELECT })
        .then(({ items }) => items)
        .catch((err) => { fieldsPromise = null; throw err; });
    }
    return fieldsPromise;
  }

  function structurePane(parsed) {
    const wrap = el('div', 'wb-tab-pane');
    const tree = el('div', 'wb-canvas-tree');
    const { sections, unplaced } = buildSectionTree(parsed.controls);
    if (!sections.length && !unplaced.length) {
      tree.append(el('div', 'wb-grid-status', 'No canvas sections on this page.'));
    }
    sections.forEach((section, i) => {
      const bits = [`${section.columns.length} column${section.columns.length === 1 ? '' : 's'}`];
      if (section.emphasis) bits.push(`emphasis ${section.emphasis}`);
      if (section.vertical) bits.push('vertical');
      if (section.collapsible) bits.push('collapsible');
      tree.append(el('div', 'wb-canvas-section', `Section ${i + 1} — ${bits.join(', ')}`));
      for (const column of section.columns) {
        const row = el('div', 'wb-canvas-column');
        const width = typeof column.sectionFactor === 'number'
          ? `${column.sectionFactor}/12` : 'auto';
        row.append(el('span', 'wb-canvas-width', width));
        if (!column.controls.length) row.append(el('span', 'wb-canvas-chip wb-canvas-empty', 'empty'));
        for (const control of column.controls) {
          const chipLabel = control.kind === 'text' ? 'Text'
            : control.kind === 'webpart'
              ? (control.webPartData.title || webPartName(control.webPartId))
              : control.kind;
          const chip = el('span', 'wb-canvas-chip', chipLabel);
          chip.title = control.kind === 'webpart'
            ? `${webPartName(control.webPartId)} · ${control.webPartId}`
            : textOfControl(control).slice(0, 200);
          row.append(chip);
        }
        tree.append(row);
      }
    });
    if (unplaced.length) {
      tree.append(el('div', 'wb-canvas-section', `Unplaced entries (${unplaced.length})`));
      for (const control of unplaced) {
        const row = el('div', 'wb-canvas-column');
        row.append(el('span', 'wb-canvas-chip', control.kind));
        tree.append(row);
      }
    }
    wrap.append(tree);
    return wrap;
  }

  function webPartsPane(parsed) {
    const wrap = el('div', 'wb-tab-pane');
    const rows = parsed.controls
      .filter((c) => c.kind === 'webpart')
      .map((c, i) => ({
        Id: c.id || String(i),
        Title: c.webPartData.title,
        Type: webPartName(c.webPartId),
        WebPartId: c.webPartId,
        ControlId: c.id,
        Text: textOfControl(c).slice(0, 160),
      }));
    const partsGrid = createGrid({
      columns: [
        { key: 'Title', label: 'Title' },
        { key: 'Type', label: 'Type' },
        { key: 'WebPartId', label: 'Web part id', mono: true, copyable: true },
        { key: 'ControlId', label: 'Control id', mono: true, copyable: true },
        { key: 'Text', label: 'Text' },
      ],
      emptyText: 'No client-side web parts on this page.',
      filterPlaceholder: 'Filter web parts…',
      exportName: 'sp-page-webparts',
    });
    wrap.append(partsGrid.el);
    partsGrid.setRows(rows);
    return wrap;
  }

  function classicWebPartsPane(webParts, error) {
    const wrap = el('div', 'wb-tab-pane');
    if (error) {
      const notice = el('div', 'wb-grid-notice',
        '⚠ The page’s web parts could not be read — ' + (error.message || String(error)));
      wrap.append(notice);
    }
    const rows = webParts.map((wp) => ({
      Title: wp.title,
      Zone: wp.zoneIndex,
      Content: wp.content ? 'inline' : wp.contentLink ? 'linked' : '',
      ContentLink: wp.contentLink,
      Hidden: wp.hidden ? 'yes' : '',
      Closed: wp.closed ? 'yes' : '',
      Id: wp.id,
    }));
    const partsGrid = createGrid({
      columns: [
        { key: 'Title', label: 'Title' },
        { key: 'Zone', label: 'Zone index' },
        { key: 'Content', label: 'Content' },
        { key: 'ContentLink', label: 'Content link', mono: true, copyable: true },
        { key: 'Hidden', label: 'Hidden' },
        { key: 'Closed', label: 'Closed' },
        { key: 'Id', label: 'Web part id', mono: true, copyable: true },
      ],
      emptyText: 'No web parts on this page.',
      filterPlaceholder: 'Filter web parts…',
      exportName: 'sp-page-webparts',
    });
    wrap.append(partsGrid.el);
    partsGrid.setRows(rows);
    return wrap;
  }

  // Extract is the reading tab: one box with the whole page's content in
  // document order under a heading per part, and one box with the underlying
  // HTML. Empty parts are skipped, and nothing here names ids or control
  // types — that lives on Web parts, Structure and Raw.
  function textPane(parts, notice) {
    const wrap = el('div', 'wb-tab-pane wb-text-pane');
    if (notice) wrap.append(notice);
    if (!parts.length) {
      wrap.append(el('div', 'wb-grid-status', 'No readable content on this page.'));
      return wrap;
    }

    const contentBlock = el('div', 'wb-text-block');
    contentBlock.append(el('div', 'wb-subpanel-title', 'Content'));
    const rendered = el('div', 'wb-text-rendered');
    for (const part of parts) {
      rendered.append(el('h3', 'wb-text-part', part.label));
      if (part.kind === 'text') {
        const body = el('div', 'wb-text-body');
        body.innerHTML = sanitizeHtml(part.html);
        rendered.append(body);
      } else {
        const list = el('ul', 'wb-text-lines');
        for (const line of part.lines) list.append(el('li', '', line));
        rendered.append(list);
      }
    }
    contentBlock.append(rendered);
    wrap.append(contentBlock);

    const withHtml = parts.filter((p) => p.kind === 'text');
    if (withHtml.length) {
      const htmlBlock = el('div', 'wb-text-block');
      htmlBlock.append(el('div', 'wb-subpanel-title', 'HTML'));
      htmlBlock.append(el('pre', 'wb-text-raw',
        withHtml.map((p) => `<!-- ${p.label} -->\n${p.html}`).join('\n\n')));
      wrap.append(htmlBlock);
    }
    return wrap;
  }

  function metadataPane(listId, pageId) {
    const wrap = el('div', 'wb-tab-pane');
    const status = el('div', 'wb-grid-status', 'Loading metadata…');
    wrap.append(status);

    (async () => {
      const fields = await listFields(listId);
      // Values + display text. FieldValuesAsText covers complex types; if the
      // combined expand misbehaves on a tenant, fall back to two requests.
      let item;
      let itemAsText = {};
      try {
        item = await client.get(guidPath(listId, `/items(${pageId})`), {
          expand: 'FieldValuesAsText',
        });
        itemAsText = item.FieldValuesAsText || {};
      } catch {
        item = await client.get(guidPath(listId, `/items(${pageId})`));
        try {
          itemAsText = await client.get(guidPath(listId, `/items(${pageId})/FieldValuesAsText`));
        } catch { itemAsText = {}; }
      }
      status.remove();
      const form = createFieldEditorForm({
        fields,
        item,
        itemAsText,
        onSave: (formValues) =>
          spWrite.validateUpdateListItem({ listId, itemId: pageId }, formValues),
      });
      wrap.append(form.el);
    })().catch((err) => {
      showFailure(status, err, 'this page’s metadata');
    });
    return wrap;
  }

  function rawPane(item, parsed, webParts = []) {
    const wrap = el('div', 'wb-tab-pane');
    const payload = { item, parsedCanvas: parsed.controls };
    if (webParts.length) payload.webParts = webParts;
    const node = toNode(payload, 0, { maxDepth: 10, maxItems: 400 });
    const inspector = el('div', 'wb-raw');
    inspector.append(enhance(node) ?? renderValue(node));
    wrap.append(inspector);
    return wrap;
  }

  async function showDetail(route) {
    const run = ++detailRun;
    gridPane.hidden = true;
    detailPane.hidden = false;
    detailPane.textContent = '';

    const back = el('button', 'btn btn-xs wb-back', '← All pages');
    back.type = 'button';
    back.addEventListener('click', () => navigate({ view: 'pages', libId: current?.listId }));
    const title = el('h2', '', route.pageName || `Page ${route.pageId}`);
    const headRow = el('div', 'wb-detail-head');
    headRow.append(back, title);
    detailPane.append(headRow);

    const status = el('div', 'wb-grid-status', 'Loading page…');
    detailPane.append(status);

    let sitePages;
    let item;
    let lostFields = '';
  let lostReason = '';
    try {
      await pagesLibraries();
      sitePages = current;
      if (!sitePages) throw new Error('This web has no pages library.');
      const plan = await queryPlan(sitePages);
      ({ item, lost: lostFields, reason: lostReason } = await pageItem(
        sitePages.listId, route.pageId, plan.detailShapes,
      ));
    } catch (err) {
      if (run !== detailRun) return;
      showFailure(status, err, 'this page');
      return;
    }
    if (run !== detailRun) return;
    status.remove();

    // The server-relative URL fragment, click-to-copy the FULL absolute URL.
    // (The item id lives in the Metadata tab.)
    if (item.FileRef) {
      const origin = (() => {
        try { return new URL(client.webUrl()).origin; } catch { return ''; }
      })();
      const fullUrl = `${origin}${encodedServerPath(item.FileRef)}`;
      const frag = el('span', 'wb-detail-id sp-copy', item.FileRef);
      frag.title = `Click to copy the full URL\n${fullUrl}`;
      frag.addEventListener('click', () => copyText(fullUrl, frag));
      headRow.append(frag);
    }

    const parsed = parseCanvasContent(item.CanvasContent1);
    const contentKind = pageContentKindOf(item);
    const isCanvas = contentKind === 'canvas';

    // Classic pages: most of the content is in web parts, not item fields.
    let classicParts = [];
    let webParts = [];
    let webPartError = null;
    if (!isCanvas) {
      const fetched = await classicWebPartsOf(item.FileRef);
      if (run !== detailRun) return;
      webParts = fetched.parts;
      webPartError = fetched.error;
      classicParts = classicContentParts({ item, webParts, contentKind }).parts;
    }
    const readingParts = isCanvas ? contentParts(parsed.controls).parts : classicParts;

    // The chip must describe what Extract actually shows: a page whose body
    // field is empty but whose web parts carry readable content is a
    // web-part page, not "no readable body".
    const displayKind = (!isCanvas && contentKind === 'empty' && readingParts.length)
      ? 'webparts' : contentKind;
    const kindChip = el('span', 'wb-info-chip wb-detail-kind', pageContentKindLabel(displayKind));
    kindChip.title = isCanvas
      ? 'Modern canvas page — Structure shows its sections and columns.'
      : `${pageContentKindLabel(displayKind)} — no canvas sections or columns, so the `
        + 'Structure tab does not apply. Content Editor and Script Editor web-part '
        + 'content is merged into Extract.';
    headRow.append(kindChip);
    // Sits beside the kind chip, same quiet register: the page is fully
    // readable, one or two fields just aren't in it.
    if (lostFields) headRow.append(reducedChip(lostFields, 'this page', lostReason));

    const actions = el('span', 'wb-detail-actions');
    const exportContent = el('button', 'btn btn-xs', 'Export content');
    exportContent.type = 'button';
    exportContent.title = 'One human-readable file: metadata, merged web-part content, full metadata';
    const exportRaw = el('button', 'btn btn-xs', 'Export raw');
    exportRaw.type = 'button';
    exportRaw.title = 'Item + parsed canvas controls as JSON, for scripts';
    actions.append(exportContent, exportRaw);
    if (item.FileRef) {
      const open = el('a', 'btn btn-xs', 'Open page ↗');
      open.href = item.FileRef;
      bindNewTab(open);
      actions.append(open);
    }
    headRow.append(actions);

    exportContent.addEventListener('click', async () => {
      const web = await webIdentity();
      downloadText(`${exportFileStem(item)}-content.md`, buildContentExport({
        item,
        controls: parsed.controls,
        parts: readingParts,
        siteTitle: web.Title || '',
        webUrl: web.Url || client.webUrl(),
        libraryTitle: sitePages.title,
        libraryRootPath: sitePages.rootPath,
      }), 'text/markdown;charset=utf-8');
    });
    exportRaw.addEventListener('click', () => {
      downloadText(`${exportFileStem(item)}-raw.json`,
        buildRawExport({ item, controls: parsed.controls, webParts }),
        'application/json');
    });
    if (isCanvas && parsed.errors.length) {
      const notice = el('div', 'wb-grid-notice',
        `⚠ ${parsed.errors.length} canvas entr${parsed.errors.length === 1 ? 'y' : 'ies'} `
        + 'could not be fully parsed — shown raw where possible.');
      notice.title = parsed.errors.join('\n');
      detailPane.append(notice);
    }

    const tabsBar = el('div', 'wb-tabs');
    tabsBar.setAttribute('role', 'tablist');
    const body = el('div', 'wb-tab-body');
    const panes = new Map();

    // Order is Joe's spec: content first (Extract), then metadata, then the
    // structural/diagnostic tabs.
    const TABS = [
      {
        id: 'text',
        label: 'Extract',
        build: () => textPane(readingParts, webPartError
          ? el('div', 'wb-grid-notice',
            '⚠ This page’s web parts could not be read, so embedded content may be '
            + `missing — ${webPartError.message || String(webPartError)}`)
          : null),
      },
      { id: 'metadata', label: 'Metadata', build: () => metadataPane(sitePages.listId, route.pageId) },
      ...(isCanvas
        ? [{ id: 'structure', label: 'Structure', build: () => structurePane(parsed) }]
        : []),
      {
        id: 'webparts',
        label: 'Web parts',
        build: () => (isCanvas
          ? webPartsPane(parsed)
          : classicWebPartsPane(webParts, webPartError)),
      },
      { id: 'raw', label: 'Raw', build: () => rawPane(item, parsed, webParts) },
    ];

    function activate(tab) {
      for (const btn of tabsBar.children) {
        btn.classList.toggle('active', btn.dataset.tab === tab.id);
        btn.setAttribute('aria-selected', btn.dataset.tab === tab.id ? 'true' : 'false');
      }
      if (!panes.has(tab.id)) panes.set(tab.id, tab.build());
      body.textContent = '';
      body.append(panes.get(tab.id));
    }

    for (const tab of TABS) {
      const btn = el('button', 'wb-tab', tab.label);
      btn.type = 'button';
      btn.dataset.tab = tab.id;
      btn.setAttribute('role', 'tab');
      btn.addEventListener('click', () => activate(tab));
      tabsBar.append(btn);
    }
    detailPane.append(tabsBar, body);
    activate(TABS.find((t) => t.id === route.tab) || TABS[0]);
  }

  // Adopt the library named by the route before anything resolves against
  // `current`. A saved route outlives the closure, so on a reload `current`
  // would otherwise be the ranked default.
  //
  // Returns false when the route names a library this web no longer offers —
  // deleted, renamed, or access lost since the route was saved. That must FAIL
  // CLOSED: falling through would resolve the saved page id against the
  // default library, which is the wrong-page/wrong-write bug this routing
  // exists to prevent, just reached by a different door.
  async function applyRouteLibrary(route) {
    if (!route?.libId) return true;
    await pagesLibraries();
    if (route.libId === current?.listId) return true;
    const wanted = libraries.find((l) => l.listId === route.libId);
    if (!wanted) return false;
    adoptLibrary(wanted);
    return true;
  }

  // Persist the library actually on screen back into the stored route. The
  // rail navigates with `{ view }` alone, so re-entering Pages that way drops
  // libId even though the view still holds the selection — a later reload
  // would then land on the default library.
  function rememberLibrary() {
    if (current && libraries.length > 1) updateRoute?.({ libId: current.listId });
  }

  function showMissingLibrary() {
    detailPane.hidden = true;
    gridPane.hidden = false;
    strip.hidden = true;
    if (grid) { grid.el.remove(); grid = null; }
    pagesLoaded = false;
    masterStatus.textContent = 'The pages library this page was opened from is no '
      + 'longer available on this web — it may have been deleted or renamed, or your '
      + 'access may have changed. Pick a library to continue.';
    masterStatus.classList.add('wb-error');
    masterStatus.hidden = false;
  }

  function load(route) {
    if (route?.pageId) {
      detailRun += 1;
      const run = detailRun;
      applyRouteLibrary(route)
        .catch(() => false)
        .then((ok) => {
          if (run !== detailRun) return;
          if (!ok) { showMissingLibrary(); return; }
          rememberLibrary();
          showDetail(route);
        });
    } else {
      detailRun += 1;
      detailPane.hidden = true;
      gridPane.hidden = false;
      applyRouteLibrary(route)
        .catch(() => false)
        .then((ok) => {
          masterStatus.classList.remove('wb-error');
          if (!ok) { showMissingLibrary(); return; }
          rememberLibrary();
          loadPages();
        });
    }
  }

  return { el: root, load };
}
