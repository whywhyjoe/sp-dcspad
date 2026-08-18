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
import {
  libraryKindOf, libraryKindLabel, pageContentKindOf, pageContentKindLabel,
  classicWebParts, classicContentParts,
} from '../classic-page.js';

// Fields every pages library has, whatever its template. PromotedState is
// modern-only — selecting it against a classic publishing library 400s with
// "The field or property 'PromotedState' does not exist", which is why the
// grid select is assembled per library kind rather than hardcoded.
const PAGE_SELECT_BASE = [
  'Id', 'Title', 'FileLeafRef', 'FileRef', 'FileDirRef',
  'Modified', 'UniqueId', 'Editor/Title',
];
const PAGE_SELECT_MODERN = [...PAGE_SELECT_BASE, 'PromotedState'];

const pageSelectFor = (kind) => (kind === 'modern' ? PAGE_SELECT_MODERN : PAGE_SELECT_BASE);

// Modern-only: every field here exists on a 119 Site Pages item and nowhere
// else. Classic libraries are fetched WITHOUT a $select instead (see
// pageItem) — asking for a field a publishing/wiki schema lacks is a 400, and
// the field set varies per site, so taking the whole item is both safer and
// the only way to reach PublishingPageContent/WikiField without probing.
const DETAIL_SELECT = [
  'Id', 'Title', 'FileLeafRef', 'FileRef', 'FileDirRef', 'Description',
  'BannerImageUrl', 'PromotedState', 'Created', 'Modified',
  'Author/Title', 'Editor/Title', 'CanvasContent1', 'LayoutWebpartsContent',
];

const FIELD_SELECT = [
  'Id', 'Title', 'InternalName', 'TypeAsString', 'FieldTypeKind', 'Required',
  'Hidden', 'ReadOnlyField', 'Group', 'DefaultValue', 'Choices', 'Description',
  'FillInChoice',
];

const SITE_PAGES_BASE_TEMPLATE = 119;
const PUBLISHING_PAGES_BASE_TEMPLATE = 850;

// Locate the web's pages library: modern Site Pages (119) first, then the
// classic publishing "Pages" library (850), then any visible library that
// is simply titled Pages — older sites use all three shapes.
export function pickPagesLibrary(items) {
  const lists = items || [];
  return lists.find((l) => l.BaseTemplate === SITE_PAGES_BASE_TEMPLATE && !l.Hidden)
    || lists.find((l) => l.BaseTemplate === SITE_PAGES_BASE_TEMPLATE)
    || lists.find((l) => l.BaseTemplate === PUBLISHING_PAGES_BASE_TEMPLATE && !l.Hidden)
    || lists.find((l) => l.BaseTemplate === PUBLISHING_PAGES_BASE_TEMPLATE)
    || lists.find((l) => String(l.Title).toLowerCase() === 'pages' && !l.Hidden)
    || null;
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

export function createPagesView({ client, navigate }) {
  const root = el('section', 'wb-view wb-view-pages');
  const spWrite = createSpWriteClient({ client });

  // ---- master pane ----
  const gridPane = el('div', 'wb-pane');
  const head = el('div', 'wb-view-head');
  head.innerHTML = '<h2>Pages</h2>';
  // Filled in once the library is resolved — modern and classic libraries are
  // both supported and the view must say which one it is looking at.
  const hint = el('p', 'wb-view-hint', 'Locating this web’s pages library…');
  head.append(hint);
  const libraryLink = el('a', 'btn btn-xs wb-head-link', 'Open library ↗');
  bindNewTab(libraryLink);
  libraryLink.hidden = true;
  head.append(libraryLink);
  const masterStatus = el('div', 'wb-grid-status');
  masterStatus.hidden = true;
  gridPane.append(head, masterStatus);

  const detailPane = el('div', 'wb-pane');
  detailPane.hidden = true;
  root.append(gridPane, detailPane);

  let sitePagesPromise = null;   // -> { listId, title } | null
  let grid = null;
  let pagesLoaded = false;
  const detailCache = new Map();   // pageId -> Promise<item>
  let fieldsPromise = null;        // list fields shared by every page
  let detailRun = 0;

  function sitePagesList() {
    if (!sitePagesPromise) {
      sitePagesPromise = client.getAll('web/lists', {
        select: ['Id', 'Title', 'BaseTemplate', 'Hidden', 'DefaultViewUrl', 'RootFolder/ServerRelativeUrl'],
        expand: 'RootFolder',
        top: 5000,
      }).then(({ items }) => {
        // Client-side filter: the mock resolver ignores $filter, and the
        // library is cheap to find in the full list either way.
        const found = pickPagesLibrary(items);
        return found ? {
          listId: found.Id,
          title: found.Title,
          kind: libraryKindOf(found),
          baseTemplate: found.BaseTemplate,
          rootPath: found.RootFolder?.ServerRelativeUrl || '',
          viewUrl: found.DefaultViewUrl || found.RootFolder?.ServerRelativeUrl || '',
        } : null;
      }).catch((err) => {
        sitePagesPromise = null;
        throw err;
      });
    }
    return sitePagesPromise;
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
    masterStatus.hidden = true;
    try {
      const sitePages = await sitePagesList();
      if (!sitePages) {
        hint.textContent = '';
        masterStatus.textContent = 'This web has no pages library — looked for modern '
          + 'Site Pages (BaseTemplate 119), classic publishing Pages (850), and any '
          + 'library titled “Pages”.';
        masterStatus.hidden = false;
        return;
      }
      hint.textContent = `${sitePages.title} — ${libraryKindLabel(sitePages.kind)} `
        + `(BaseTemplate ${sitePages.baseTemplate}), subfolders included. `
        + 'Click a row to inspect content, metadata, and structure.';
      if (sitePages.viewUrl) {
        libraryLink.href = sitePages.viewUrl;
        libraryLink.textContent = `Open ${sitePages.title} ↗`;
        libraryLink.hidden = false;
      }
      if (!grid) {
        const query = {
          path: guidPath(sitePages.listId, '/items'),
          options: {
            select: pageSelectFor(sitePages.kind),
            expand: 'Editor',
            orderby: 'FileLeafRef',
            top: 5000,
          },
        };
        grid = createGrid({
          columns: [
            { key: 'FileLeafRef', label: 'Name', mono: true },
            { key: 'Title', label: 'Title' },
            {
              key: 'Folder',
              label: 'Folder',
              value: (row) => folderOf(row.FileDirRef, sitePages.rootPath),
              format: (v) => (v ? `/${v}` : ''),
            },
            ...(sitePages.kind === 'modern'
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
            view: 'pages', pageId: row.Id, pageName: row.FileLeafRef || row.Title,
          }),
          emptyText: `No pages in ${sitePages.title}.`,
          filterPlaceholder: 'Filter pages…',
          exportName: 'sp-pages',
          descriptor: { ...query, webUrl: client.webUrl() },
        });
        gridPane.append(grid.el);
        grid.setLoading('Loading pages…');
        const { items, partial } = await client.getAll(query.path, query.options);
        grid.setRows(items, { partial });
        pagesLoaded = true;
      }
    } catch (err) {
      if (grid) grid.setError(err);
      else {
        masterStatus.textContent = err?.message || String(err);
        masterStatus.classList.add('wb-error');
        masterStatus.hidden = false;
      }
    }
  }

  // ---- drilldown ----

  function pageItem(listId, pageId, kind) {
    if (!detailCache.has(pageId)) {
      // No $select off the modern path: the classic body fields differ per
      // site and naming one that is absent fails the whole request.
      const options = kind === 'modern'
        ? { select: DETAIL_SELECT, expand: ['Author', 'Editor'] }
        : { expand: ['Author', 'Editor'] };
      detailCache.set(pageId, client.get(guidPath(listId, `/items(${pageId})`), options)
        .catch((err) => {
          detailCache.delete(pageId);
          throw err;
        }));
    }
    return detailCache.get(pageId);
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
        .catch((err) => ({ parts: [], error: err })));
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
      status.textContent = err?.message || String(err);
      status.classList.add('wb-error');
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
    back.addEventListener('click', () => navigate({ view: 'pages' }));
    const title = el('h2', '', route.pageName || `Page ${route.pageId}`);
    const headRow = el('div', 'wb-detail-head');
    headRow.append(back, title);
    detailPane.append(headRow);

    const status = el('div', 'wb-grid-status', 'Loading page…');
    detailPane.append(status);

    let sitePages;
    let item;
    try {
      sitePages = await sitePagesList();
      if (!sitePages) throw new Error('This web has no pages library.');
      item = await pageItem(sitePages.listId, route.pageId, sitePages.kind);
    } catch (err) {
      if (run !== detailRun) return;
      status.textContent = err?.message || String(err);
      status.classList.add('wb-error');
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

    const kindChip = el('span', 'wb-detail-kind', pageContentKindLabel(contentKind));
    kindChip.title = isCanvas
      ? 'Modern canvas page — Structure shows its sections and columns.'
      : `${pageContentKindLabel(contentKind)} — no canvas sections or columns, so the `
        + 'Structure tab does not apply. Content Editor and Script Editor web-part '
        + 'content is merged into Extract.';
    headRow.append(kindChip);

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

  function load(route) {
    if (route?.pageId) {
      showDetail(route);
    } else {
      detailRun += 1;
      detailPane.hidden = true;
      gridPane.hidden = false;
      loadPages();
    }
  }

  return { el: root, load };
}
