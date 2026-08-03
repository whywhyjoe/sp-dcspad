// Pages view: the modern-page inspector. Master grid over the Site Pages
// library (BaseTemplate 119), drilldown parses CanvasContent1 into a
// section/column structure tree, a web-part inventory, extracted text
// content, an editable metadata sheet, and the raw entity.

import { createGrid } from '../grid.js?v=2';
import { copyText } from '../export.js';
import {
  parseCanvasContent, buildSectionTree, textOfControl, webPartName, sanitizeHtml,
} from '../canvas.js';
import { createSpWriteClient } from '../sp-write.js';
import { createFieldEditorForm } from '../field-editor.js';
import { buildContentExport, buildRawExport, exportFileStem } from '../page-export.js';
import { downloadText } from '../../io.js?v=2';
import { enhance } from '../../inspect/sp-shapes.js';
import { renderValue } from '../../inspect/tree-view.js';
import { toNode } from '../../inspect/to-node.js';

const PAGE_SELECT = [
  'Id', 'Title', 'FileLeafRef', 'FileRef', 'FileDirRef', 'PromotedState',
  'Modified', 'UniqueId', 'Editor/Title',
];

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
  head.innerHTML = '<h2>Pages</h2>'
    + '<p class="wb-view-hint">Modern pages in this web’s Site Pages library, '
    + 'subfolders included. Click a row to inspect content, metadata, and structure.</p>';
  const libraryLink = el('a', 'btn btn-xs wb-head-link', 'Open Site Pages library ↗');
  libraryLink.target = '_blank';
  libraryLink.rel = 'noopener';
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
        const found = items.find((l) => l.BaseTemplate === SITE_PAGES_BASE_TEMPLATE && !l.Hidden)
          || items.find((l) => l.BaseTemplate === SITE_PAGES_BASE_TEMPLATE);
        return found ? {
          listId: found.Id,
          title: found.Title,
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
        masterStatus.textContent = 'This web has no Site Pages library (BaseTemplate 119).';
        masterStatus.hidden = false;
        return;
      }
      if (sitePages.viewUrl) {
        libraryLink.href = sitePages.viewUrl;
        libraryLink.hidden = false;
      }
      if (!grid) {
        const query = {
          path: guidPath(sitePages.listId, '/items'),
          options: {
            select: PAGE_SELECT, expand: 'Editor', orderby: 'FileLeafRef', top: 5000,
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
            { key: 'PromotedState', label: 'Promoted', format: promotedLabel },
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
                a.target = '_blank';
                a.rel = 'noopener';
                a.title = 'Open the page in a new tab';
                a.textContent = '↗';
                a.addEventListener('click', (e) => e.stopPropagation());
                return a;
              },
            },
          ],
          onOpen: (row) => navigate({
            view: 'pages', pageId: row.Id, pageName: row.FileLeafRef || row.Title,
          }),
          emptyText: 'No pages in this library.',
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

  function pageItem(listId, pageId) {
    if (!detailCache.has(pageId)) {
      detailCache.set(pageId, client.get(guidPath(listId, `/items(${pageId})`), {
        select: DETAIL_SELECT,
        expand: ['Author', 'Editor'],
      }).catch((err) => {
        detailCache.delete(pageId);
        throw err;
      }));
    }
    return detailCache.get(pageId);
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

  function textPane(parsed) {
    const wrap = el('div', 'wb-tab-pane wb-text-pane');
    const texts = parsed.controls.filter((c) => c.kind === 'text');
    if (!texts.length) {
      wrap.append(el('div', 'wb-grid-status', 'No text web parts on this page.'));
      return wrap;
    }
    texts.forEach((control, i) => {
      const block = el('div', 'wb-text-block');
      block.append(el('div', 'wb-subpanel-title', `Text web part ${i + 1}`));
      const rendered = el('div', 'wb-text-rendered');
      rendered.innerHTML = sanitizeHtml(control.innerHTML);
      block.append(rendered);
      const details = document.createElement('details');
      details.append(el('summary', '', 'Raw HTML'));
      const pre = el('pre', 'wb-text-raw', control.innerHTML);
      details.append(pre);
      block.append(details);
      wrap.append(block);
    });
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

  function rawPane(item, parsed) {
    const wrap = el('div', 'wb-tab-pane');
    const node = toNode({ item, parsedCanvas: parsed.controls }, 0, { maxDepth: 10, maxItems: 400 });
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
      if (!sitePages) throw new Error('This web has no Site Pages library.');
      item = await pageItem(sitePages.listId, route.pageId);
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
      open.target = '_blank';
      open.rel = 'noopener';
      actions.append(open);
    }
    headRow.append(actions);

    const parsed = parseCanvasContent(item.CanvasContent1);

    exportContent.addEventListener('click', async () => {
      const web = await webIdentity();
      downloadText(`${exportFileStem(item)}-content.md`, buildContentExport({
        item,
        controls: parsed.controls,
        siteTitle: web.Title || '',
        webUrl: web.Url || client.webUrl(),
        libraryTitle: sitePages.title,
        libraryRootPath: sitePages.rootPath,
      }), 'text/markdown;charset=utf-8');
    });
    exportRaw.addEventListener('click', () => {
      downloadText(`${exportFileStem(item)}-raw.json`,
        buildRawExport({ item, controls: parsed.controls }),
        'application/json');
    });
    if (parsed.errors.length) {
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
      { id: 'text', label: 'Extract', build: () => textPane(parsed) },
      { id: 'metadata', label: 'Metadata', build: () => metadataPane(sitePages.listId, route.pageId) },
      { id: 'structure', label: 'Structure', build: () => structurePane(parsed) },
      { id: 'webparts', label: 'Web parts', build: () => webPartsPane(parsed) },
      { id: 'raw', label: 'Raw', build: () => rawPane(item, parsed) },
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
