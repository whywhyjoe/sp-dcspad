// Files view: a featureful document-library browser for the inspected web.
// Unlike the pad's picker (sp-files.js listFolder — purpose-filtered to code
// and text types, unpaged), this browser lists EVERY file through the paged
// workbench REST client, downloads through SharePoint's download endpoint,
// uploads any file type as binary, and reads/edits the full metadata column
// set through the shared field-editor + sp-write plumbing.

import { showFailure } from '../denied.js';
import { createGrid } from '../grid.js?v=2';
import { copyText } from '../export.js';
import { odataPathLiteral } from '../../sp-odata.js';
import { createSpWriteClient, MAX_UPLOAD_BYTES } from '../sp-write.js';
import { createFieldEditorForm } from '../field-editor.js';
// One rule for "checked out", and for "checked out to me", shared with the
// pad's SharePoint export so the two cannot drift.
import { isCheckedOut, isCheckedOutByCurrentUser } from '../../sp-files.js';
import {
  metadataFieldStates, anyMetadataAvailable, openUploadMetadataDialog,
} from '../upload-metadata.js';

const FIELD_SELECT = [
  'Id', 'Title', 'InternalName', 'EntityPropertyName', 'TypeAsString',
  'FieldTypeKind', 'Required', 'Hidden', 'ReadOnlyField', 'Group',
  'DefaultValue', 'Choices', 'Description', 'FillInChoice',
];

const DOCUMENT_LIBRARY_BASE_TYPE = 1;
// Recorded on the version when an upload here ends a check-out.
const CHECK_IN_COMMENT = 'Uploaded from SP Workbench';
const GUID = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i;

const FOLDER_SELECT = ['Name', 'ServerRelativeUrl', 'ItemCount', 'TimeLastModified'];
const FILE_SELECT = [
  'Name', 'ServerRelativeUrl', 'Length', 'TimeLastModified', 'UIVersionLabel', 'CheckOutType',
];

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

const icon = (name, size = 15) => {
  const paths = {
    folder: [
      '<path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>',
    ],
    file: [
      '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/>',
      '<path d="M14 2v4a2 2 0 0 0 2 2h4"/>',
    ],
    download: [
      '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>',
      '<path d="m7 10 5 5 5-5"/>', '<path d="M12 15V3"/>',
    ],
    link: [
      '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>',
      '<path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>',
    ],
  };
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('width', String(size));
  svg.setAttribute('height', String(size));
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '2');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.setAttribute('aria-hidden', 'true');
  svg.innerHTML = (paths[name] || []).join('');
  return svg;
};

const fileRole = (row) => {
  if (row.kind === 'folder') return /^forms$|^_/i.test(row.Name || '') ? 'sys' : 'user';
  const ext = extOf(row.Name);
  if (['js', 'mjs', 'cjs', 'ts', 'tsx'].includes(ext)) return 'js';
  if (['html', 'htm', 'svg'].includes(ext)) return 'html';
  if (['css', 'scss', 'less'].includes(ext)) return 'css';
  if (['json', 'csv', 'tsv', 'xml', 'xlsx', 'xls'].includes(ext)) return 'json';
  if (['doc', 'docx', 'pdf', 'ppt', 'pptx', 'rtf'].includes(ext)) return 'doc';
  return 'file';
};

const encodedServerPath = (path) => String(path || '').split('/').map((segment) => {
  try { return encodeURIComponent(decodeURIComponent(segment)); }
  catch { return encodeURIComponent(segment); }
}).join('/');

const fmtDate = (v) => (v ? String(v).slice(0, 10) : '');

export function formatBytes(n) {
  const bytes = Number(n);
  if (!Number.isFinite(bytes) || bytes < 0) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}

const extOf = (name) => {
  const dot = String(name || '').lastIndexOf('.');
  return dot > 0 ? String(name).slice(dot + 1).toLowerCase() : '';
};

function normalizedPath(value) {
  let path = String(value || '').trim().replaceAll('\\', '/');
  if (!path.startsWith('/')) path = `/${path}`;
  path = path.replace(/\/{2,}/g, '/');
  if (path.length > 1) path = path.replace(/\/+$/, '');
  return path;
}

export function createBrowserView({ client, navigate }) {
  const root = el('section', 'wb-view wb-view-files');
  const spWrite = createSpWriteClient({ client });

  const head = el('div', 'wb-view-head');
  head.innerHTML = '<h2>Files</h2>'
    + '<p class="wb-view-hint">Browse any library or folder of this web — every '
    + 'file type, with download, binary upload, folder creation, and full '
    + 'metadata editing.</p>';

  // Library picker + breadcrumbs. Rendered here until the grid exists, then
  // adopted into the grid toolbar (toolbarExtras) so location, filter, and
  // actions share one row instead of stacking three.
  const bar = el('div', 'wb-crumbs-bar');
  const librarySelect = el('select', 'wb-lib-select');
  librarySelect.setAttribute('aria-label', 'Jump to a document library');
  const crumbs = el('div', 'wb-crumbs');
  bar.append(librarySelect, crumbs);

  const consent = el('div', 'wb-consent');
  consent.hidden = true;

  const gridWrap = el('div', 'wb-files-grid');
  const metaPanel = el('div', 'wb-subpanel wb-file-meta');
  metaPanel.hidden = true;

  root.append(head, bar, consent, gridWrap, metaPanel);

  // ---- state ----
  let libraries = [];
  let currentPath = '';
  let currentListing = { folders: [], files: [] };
  let grid = null;
  let librariesLoaded = false;
  let listingRun = 0;
  const parentListCache = new Map();   // folder path -> Promise<listId>
  const fieldsCache = new Map();       // listId -> Promise<fields>

  function webRootPath() {
    try {
      return normalizedPath(decodeURIComponent(new URL(client.webUrl()).pathname)) || '/';
    } catch { return '/'; }
  }

  function checkedPath(path) {
    const rootPath = webRootPath();
    const normalized = normalizedPath(path || rootPath);
    if (rootPath !== '/' && normalized !== rootPath && !normalized.startsWith(`${rootPath}/`)) {
      throw new Error('That path is outside the inspected web.');
    }
    return normalized;
  }

  const folderApi = (path, sub) =>
    `web/GetFolderByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')${sub}`;
  const fileApi = (path, sub) =>
    `web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')${sub}`;

  function downloadHref(serverRelativeUrl) {
    if (spWrite.isMock()) return serverRelativeUrl;
    return `${client.webUrl()}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;
  }

  // ---- breadcrumbs + library picker ----
  function renderCrumbs() {
    crumbs.textContent = '';
    const rootPath = webRootPath();
    const segments = currentPath === '/' ? [] : currentPath.slice(1).split('/');
    let acc = '';
    const rootBtn = el('button', 'wb-crumb', rootPath === '/' ? '/' : rootPath);
    rootBtn.type = 'button';
    rootBtn.addEventListener('click', () => navigate({ view: 'files', path: rootPath }));
    // Only show the web-root crumb when the path descends from it visibly.
    let started = rootPath === '/';
    if (started) crumbs.append(rootBtn);
    for (const segment of segments) {
      acc += `/${segment}`;
      if (!started) {
        if (normalizedPath(acc) === rootPath) {
          started = true;
          const btn = el('button', 'wb-crumb', rootPath);
          btn.type = 'button';
          btn.addEventListener('click', () => navigate({ view: 'files', path: rootPath }));
          crumbs.append(btn);
        }
        continue;
      }
      crumbs.append(el('span', 'wb-crumb-sep', '/'));
      const target = acc;
      const btn = el('button', 'wb-crumb', segment);
      btn.type = 'button';
      btn.addEventListener('click', () => navigate({ view: 'files', path: target }));
      crumbs.append(btn);
    }
    syncCrumbOverflow();
  }

  // Keep the deepest segment — where you actually are — pinned in view, let
  // the ancestors scroll off to the left, and mark that edge so the clip reads
  // as "there's more".
  //
  // This has to run on resize, not just on render: the toolbar row is elastic,
  // so narrowing the window after navigating shrinks the trail without
  // re-rendering it. Measured at 720px the crumbs box reached clientWidth 0
  // with scrollLeft still 0 — the whole trail unreachable, and the scrollbar
  // is hidden so there was no affordance to get it back.
  function syncCrumbOverflow() {
    crumbs.scrollLeft = crumbs.scrollWidth;
    crumbs.classList.toggle('is-clipped', crumbs.scrollWidth > crumbs.clientWidth + 1);
  }

  // Retained so it can be disconnected: the shell drops and recreates every
  // view instance when the inspected web changes, which would otherwise strand
  // one observer per site switch on a detached crumbs node.
  const crumbObserver = typeof ResizeObserver === 'function'
    ? new ResizeObserver(() => syncCrumbOverflow())
    : null;
  crumbObserver?.observe(crumbs);

  async function loadLibraries() {
    if (librariesLoaded) return;
    try {
      const { items } = await client.getAll('web/lists', {
        select: ['Id', 'Title', 'BaseType', 'Hidden', 'RootFolder/ServerRelativeUrl'],
        expand: 'RootFolder',
        orderby: 'Title',
        top: 5000,
      });
      libraries = items.filter((l) => l.BaseType === DOCUMENT_LIBRARY_BASE_TYPE && !l.Hidden);
      librariesLoaded = true;
    } catch { libraries = []; }
    librarySelect.textContent = '';
    const blank = el('option', '', 'Libraries…');
    blank.value = '';
    librarySelect.append(blank);
    for (const lib of libraries) {
      const url = lib.RootFolder?.ServerRelativeUrl;
      if (!url) continue;
      const opt = el('option', '', lib.Title);
      opt.value = url;
      librarySelect.append(opt);
    }
  }

  librarySelect.addEventListener('change', () => {
    if (librarySelect.value) navigate({ view: 'files', path: librarySelect.value });
  });

  // ---- listing ----
  function makeGrid() {
    grid = createGrid({
      columns: [
        {
          key: 'Name',
          label: 'Name',
          value: (row) => row.Name,
          render: (name, row) => {
            const wrap = el('span', `wb-file-name wb-node-${fileRole(row)}`);
            const glyph = icon(row.kind === 'folder' ? 'folder' : 'file');
            glyph.classList.add('wb-node');
            wrap.append(glyph, el('span', 'wb-file-name-text', name));
            return wrap;
          },
        },
        { key: 'Type', label: 'Type', value: (row) => (row.kind === 'folder' ? 'Folder' : extOf(row.Name)) },
        { key: 'Length', label: 'Size', num: true, value: (row) => (row.kind === 'folder' ? null : Number(row.Length) || 0), format: (v, row) => (row.kind === 'folder' ? '' : formatBytes(v)) },
        { key: 'TimeLastModified', label: 'Modified', format: fmtDate },
        { key: 'UIVersionLabel', label: 'Version', value: (row) => (row.kind === 'folder' ? '' : row.UIVersionLabel || '') },
        {
          key: 'Actions',
          label: '',
          value: (row) => row.ServerRelativeUrl,
          format: () => '',
          render: (serverRelativeUrl, row) => {
            if (row.kind === 'folder') return null;
            const span = document.createElement('span');
            span.className = 'wb-file-actions';
            const dl = document.createElement('a');
            dl.className = 'wb-cell-link';
            dl.href = downloadHref(serverRelativeUrl);
            dl.title = 'Download';
            dl.setAttribute('aria-label', `Download ${row.Name}`);
            dl.append(icon('download', 13));
            if (!spWrite.isMock()) dl.setAttribute('download', row.Name);
            dl.addEventListener('click', (e) => e.stopPropagation());
            span.append(dl);
            const link = document.createElement('button');
            link.type = 'button';
            link.className = 'wb-cell-link wb-cell-copylink';
            link.title = 'Copy the direct URL';
            link.setAttribute('aria-label', `Copy the direct URL for ${row.Name}`);
            link.append(icon('link', 13));
            link.addEventListener('click', (e) => {
              e.stopPropagation();
              // Direct (non-sharing) URL: web origin + encoded server path.
              const origin = new URL(client.webUrl()).origin;
              copyText(`${origin}${encodedServerPath(serverRelativeUrl)}`, link);
            });
            span.append(link);
            return span;
          },
        },
      ],
      rowKey: 'ServerRelativeUrl',
      onOpen: (row) => {
        if (row.kind === 'folder') navigate({ view: 'files', path: row.ServerRelativeUrl });
        else openMetadata(row);
      },
      emptyText: 'This folder is empty.',
      subject: 'this folder',
      filterPlaceholder: 'Filter files…',
      exportName: 'sp-files',
      toolbarExtras: bar,
    });

    // Toolbar: upload button + hidden input, refresh.
    const uploadBtn = el('button', 'btn btn-xs wb-primary', 'Upload…');
    uploadBtn.type = 'button';
    const fileInput = el('input');
    fileInput.type = 'file';
    fileInput.hidden = true;
    fileInput.setAttribute('aria-label', 'Choose a file to upload');
    uploadBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
      if (fileInput.files?.length) startUpload(fileInput.files[0]);
      fileInput.value = '';
    });
    const newFolderBtn = el('button', 'btn btn-xs wb-newfolder', 'New folder…');
    newFolderBtn.type = 'button';
    newFolderBtn.addEventListener('click', promptNewFolder);
    const refreshBtn = el('button', 'btn btn-xs', 'Refresh');
    refreshBtn.type = 'button';
    refreshBtn.addEventListener('click', () => listFolder(currentPath, { force: true }));
    grid.actionsEl.prepend(uploadBtn, fileInput, newFolderBtn, refreshBtn);
    gridWrap.append(grid.el);
  }

  async function listFolder(path, { force = false } = {}) {
    void force;
    const run = ++listingRun;
    currentPath = checkedPath(path);
    renderCrumbs();
    metaPanel.hidden = true;
    consent.hidden = true;
    if (!grid) makeGrid();
    grid.setLoading('Loading folder…');
    try {
      const [folders, files] = await Promise.all([
        client.getAll(folderApi(currentPath, '/Folders'), { select: FOLDER_SELECT, top: 5000 }),
        client.getAll(folderApi(currentPath, '/Files'), { select: FILE_SELECT, top: 5000 }),
      ]);
      const sortByName = (a, b) =>
        String(a.Name).localeCompare(String(b.Name), undefined, { sensitivity: 'base' });
      if (run !== listingRun) return;
      currentListing = {
        folders: folders.items.map((f) => ({ ...f, kind: 'folder' })).sort(sortByName),
        files: files.items.map((f) => ({ ...f, kind: 'file' })).sort(sortByName),
      };
      grid.setRows([...currentListing.folders, ...currentListing.files], {
        partial: folders.partial || files.partial,
      });
      const matching = [...librarySelect.options].find((o) => o.value
        && (currentPath === o.value || currentPath.startsWith(`${o.value}/`)));
      librarySelect.value = matching ? matching.value : '';
    } catch (err) {
      if (run !== listingRun) return;
      grid.setError(err);
    }
  }

  // ---- upload ----
  // `gate`, when given, is a sentence the operator must tick before Replace
  // becomes clickable — the forced-check-out consent. It rides the pad's
  // .sp-metadata-consent classes (app.css loads before workbench.css), so
  // the two consents look like one component.
  function showConsent(message, onConfirm, { gate = '' } = {}) {
    consent.textContent = '';
    consent.hidden = false;
    consent.append(el('span', 'wb-consent-text', message));
    const replace = el('button', 'btn btn-xs', 'Replace');
    replace.type = 'button';
    replace.addEventListener('click', () => { consent.hidden = true; onConfirm(); });
    const cancel = el('button', 'btn btn-xs', 'Cancel');
    cancel.type = 'button';
    cancel.addEventListener('click', () => { consent.hidden = true; });
    if (gate) {
      replace.disabled = true;
      const row = el('label', 'sp-metadata-consent__row wb-consent-gate');
      const box = el('input');
      box.type = 'checkbox';
      box.className = 'wb-consent-checkout';
      box.addEventListener('change', () => { replace.disabled = !box.checked; });
      row.append(box, el('span', 'sp-metadata-consent__label', gate));
      consent.append(row);
    }
    consent.append(replace, cancel);
  }

  function uploadNotice(message, isError = false) {
    consent.textContent = '';
    consent.hidden = false;
    consent.classList.toggle('wb-consent-error', isError);
    consent.append(el('span', 'wb-consent-text', message));
    const dismiss = el('button', 'btn btn-xs', 'Dismiss');
    dismiss.type = 'button';
    dismiss.addEventListener('click', () => { consent.hidden = true; });
    consent.append(dismiss);
  }

  async function startUpload(file) {
    const folderPath = currentPath;
    consent.classList.remove('wb-consent-error');
    if (file.size > MAX_UPLOAD_BYTES) {
      uploadNotice(
        `“${file.name}” is ${formatBytes(file.size)} — above the ${formatBytes(MAX_UPLOAD_BYTES)} upload limit.`,
        true,
      );
      return;
    }
    const existing = currentListing.files.find(
      (f) => String(f.Name).toLowerCase() === file.name.toLowerCase(),
    );
    if (existing) {
      await confirmReplace(file, folderPath, { existing });
      return;
    }
    await runUpload(file, { overwrite: false, folderPath });
  }

  // ---- check-out -----------------------------------------------------------
  // A library with ForceCheckout refuses a write to a file that is not
  // checked out, so replacing one there is three acts: the check-out, the
  // upload, and the check-in that makes it visible. Checking a file out
  // changes what everyone else in the library sees, so it is consented to,
  // never implied by the Replace button — the same contract as the pad's
  // export dialog.
  // Deliberately uncached: it costs one GET per replace-consent (a rare,
  // already-interactive moment), and a policy someone flipped mid-session
  // must not be answered from a stale yes/no.
  async function libraryForcesCheckout(folderPath) {
    const listId = await parentListId(folderPath);
    const list = await client.get(`web/lists(guid'${listId}')`, { select: 'ForceCheckout' });
    return Boolean(list?.ForceCheckout);
  }

  // The listing already carries CheckOutType per file, so a checked-in file
  // costs no request; the holder (and, on the 409-race path, the state
  // itself) is read from the server only when it is actually needed.
  async function readCheckOut(filePath, existing) {
    if (existing && !isCheckedOut(existing.CheckOutType)) return { checkedOut: false };
    const file = await client.get(fileApi(filePath, ''), {
      select: 'CheckOutType,CheckedOutByUser/Id,CheckedOutByUser/Title,'
        + 'CheckedOutByUser/LoginName,CheckedOutByUser/Email',
      expand: 'CheckedOutByUser',
    });
    return {
      checkedOut: isCheckedOut(file?.CheckOutType),
      user: file?.CheckedOutByUser || null,
    };
  }

  // Never throws: an unreadable policy leaves SharePoint the authority and
  // the upload behaves as it did before this gate existed.
  async function checkOutState(folderPath, fileName, existing) {
    const state = {
      required: false, checkedOut: false, checkedOutByCurrentUser: false, checkedOutBy: '',
    };
    try {
      state.required = await libraryForcesCheckout(folderPath);
      // Read in any library: a file someone else holds refuses the write
      // whatever the policy, and one the operator holds needs checking in.
      const { checkedOut, user } = await readCheckOut(`${folderPath}/${fileName}`, existing);
      state.checkedOut = checkedOut;
      if (!checkedOut) return state;
      state.checkedOutBy = String(user?.Title || user?.LoginName || '').trim();
      state.checkedOutByCurrentUser = isCheckedOutByCurrentUser(
        user,
        client.context()?.pageContext,
        { sameWeb: client.webUrl() === client.hostWebUrl() },
      );
    } catch { /* policy unknown — leave the gate off */ }
    return state;
  }

  async function confirmReplace(file, folderPath, { existing = null, carriedValues = null } = {}) {
    const checkout = await checkOutState(folderPath, file.name, existing);
    if (checkout.checkedOut && !checkout.checkedOutByCurrentUser) {
      uploadNotice(
        `“${file.name}” is checked out to ${checkout.checkedOutBy || 'another user'}. `
        + 'It cannot be replaced until they check it back in.',
        true,
      );
      return;
    }
    let gate = '';
    if (checkout.checkedOutByCurrentUser) {
      gate = `${checkout.required ? 'This library requires check-out — ' : ''}`
        + `“${file.name}” is already checked out to you. Replace it and check it back in.`;
    } else if (checkout.required) {
      gate = `This library requires check-out — check “${file.name}” out before replacing it, `
        + 'then back in.';
    }
    showConsent(
      `“${file.name}” already exists in this folder. Replace it?`,
      () => runUpload(file, { overwrite: true, folderPath, carriedValues, checkout }),
      { gate },
    );
  }

  // Probe the destination library for the pad's three curated metadata
  // columns (Title, _ExtendedDescription, DocVersion); null skips the
  // dialog and uploads bare.
  async function uploadMetadataStates(folderPath) {
    try {
      const listId = await parentListId(folderPath);
      const fields = await listFields(listId);
      const states = metadataFieldStates(fields);
      return anyMetadataAvailable(states) ? states : null;
    } catch { return null; }
  }

  // Prefill from the file being replaced — best-effort, blanks otherwise.
  async function prefillUploadValues(states, folderPath, fileName) {
    const values = { title: '', description: '', docVersion: '' };
    const available = Object.values(states).filter((s) => s.available);
    try {
      const item = await client.get(
        fileApi(`${folderPath}/${fileName}`, '/ListItemAllFields'),
        { select: available.map((s) => s.entityPropertyName) },
      );
      for (const s of available) {
        values[s.key] = String(item?.[s.entityPropertyName] ?? item?.[s.internalName] ?? '');
      }
    } catch { /* prefill only */ }
    return values;
  }

  async function runUpload(file, {
    overwrite, folderPath, carriedValues = null, checkout = null,
  }) {
    let data;
    try {
      data = await file.arrayBuffer();
    } catch (err) {
      uploadNotice(`Could not read the file: ${err?.message || err}`, true);
      return;
    }
    let uploaded = null;
    let checkedOutHere = false;
    const bareUpload = async () => {
      if (overwrite && checkout?.required && !checkout.checkedOutByCurrentUser) {
        await spWrite.checkOutFile(`${folderPath}/${file.name}`);
        // The file is ours now: a retry must not check it out a second time,
        // which SharePoint rejects.
        checkout.checkedOutByCurrentUser = true;
        checkedOutHere = true;
      }
      uploaded = await spWrite.uploadFile(folderPath, file.name, data, { overwrite });
      return uploaded;
    };
    // Runs once the file (and any metadata) has landed. A check-out made or
    // already held here is released, and so is the one SharePoint puts on a
    // new file in a ForceCheckout library — the upload response says so.
    const settle = async (message) => {
      const held = checkout?.checkedOutByCurrentUser || isCheckedOut(uploaded?.checkOutType);
      if (!held) return finishUpload(file, folderPath, message);
      try {
        // The metadata write may already have checked the file in; either
        // way a file that was held has ended checked in.
        await spWrite.checkInFile(`${folderPath}/${file.name}`, { comment: CHECK_IN_COMMENT });
        return finishUpload(file, folderPath, `${message} Checked in.`);
      } catch (err) {
        return finishUpload(
          file,
          folderPath,
          `${message} It could not be checked in and is still checked out to you: `
            + `${err?.message || err}`,
          true,
        );
      }
    };
    const fail = async (err, carried) => {
      if (checkedOutHere && !uploaded) {
        // The listing still shows the file checked in. Refresh it, or the
        // next attempt reads that stale row and posts a second CheckOut(),
        // which SharePoint rejects.
        if (currentPath === folderPath) await listFolder(folderPath, { force: true });
        uploadNotice(
          `Upload failed: ${err?.message || err} “${file.name}” is still checked out to you.`,
          true,
        );
        return;
      }
      handleUploadError(err, file, folderPath, overwrite, carried);
    };
    const states = await uploadMetadataStates(folderPath);

    if (!states) {
      uploadNotice(`Uploading “${file.name}”…`);
      try {
        await bareUpload();
        await settle(`Uploaded “${file.name}” ✓`);
      } catch (err) {
        await fail(err, null);
      }
      return;
    }

    // The changed-only diff must compare against the FILE's baseline, never
    // against carried user input — otherwise a 409-race retry would treat
    // the user's typed values as "unchanged" and silently drop them.
    const baseline = overwrite
      ? await prefillUploadValues(states, folderPath, file.name)
      : { title: '', description: '', docVersion: '' };
    const values = carriedValues || baseline;
    const initial = { ...baseline };
    const filePath = `${folderPath}/${file.name}`;
    try {
      const outcome = await openUploadMetadataDialog({
        fileName: file.name,
        overwrite,
        states,
        values,
        doUpload: bareUpload,
        // Write only what the user changed against the prefill (a cleared
        // prefill still writes ''); untouched values cost no request.
        doMetadata: async (entered) => {
          const formValues = Object.values(states)
            .filter((s) => s.available
              && String(entered[s.key] ?? '') !== String(initial[s.key] ?? ''))
            .map((s) => ({ FieldName: s.internalName, FieldValue: String(entered[s.key] ?? '') }));
          if (!formValues.length) return;
          await spWrite.validateUpdateListItem(
            { fileServerRelativeUrl: filePath },
            formValues,
            { newDocumentUpdate: true, checkInComment: CHECK_IN_COMMENT },
          );
        },
      });
      if (outcome === 'cancelled') return;
      await settle(outcome === 'saved'
        ? `Uploaded “${file.name}” ✓`
        : `Uploaded “${file.name}” ✓ (kept without metadata)`);
    } catch (err) {
      await fail(err, err?.uploadMetadataValues || null);
    }
  }

  async function finishUpload(file, folderPath, message, isError = false) {
    if (currentPath === folderPath) await listFolder(folderPath, { force: true });
    uploadNotice(message, isError);
  }

  function handleUploadError(err, file, folderPath, overwrite, carriedValues) {
    if (err?.code === 'conflict' && !overwrite) {
      // Race: the file appeared between listing and upload. The retry carries
      // the values the user already typed into the dialog, and the file is
      // not in our listing, so its check-out state is read from the server.
      void confirmReplace(file, folderPath, { carriedValues, existing: null });
      return;
    }
    uploadNotice(`Upload failed: ${err?.message || err}`, true);
  }

  // ---- new folder ----
  function promptNewFolder() {
    const folderPath = currentPath;
    consent.classList.remove('wb-consent-error');
    consent.textContent = '';
    consent.hidden = false;
    consent.append(el('span', 'wb-consent-text', `New folder in ${folderPath}:`));
    const nameIn = el('input', 'wb-folder-name');
    nameIn.type = 'text';
    nameIn.placeholder = 'Folder name';
    nameIn.setAttribute('aria-label', 'New folder name');
    const create = el('button', 'btn btn-xs wb-primary', 'Create');
    create.type = 'button';
    const cancel = el('button', 'btn btn-xs', 'Cancel');
    cancel.type = 'button';
    cancel.addEventListener('click', () => { consent.hidden = true; });
    const submit = async () => {
      const name = nameIn.value.trim();
      if (!name) { nameIn.focus(); return; }
      create.disabled = true;
      try {
        await spWrite.createFolder(folderPath, name);
        await listFolder(folderPath, { force: true });
        uploadNotice(`Created folder “${name}”.`);
      } catch (err) {
        uploadNotice(`Could not create the folder: ${err?.message || err}`, true);
      }
    };
    create.addEventListener('click', submit);
    nameIn.addEventListener('keydown', (e) => { if (e.key === 'Enter') submit(); });
    consent.append(nameIn, create, cancel);
    nameIn.focus();
  }

  // ---- metadata ----
  function parentListId(folderPath) {
    const key = folderPath.toLowerCase();
    if (!parentListCache.has(key)) {
      parentListCache.set(key, client.get(folderApi(folderPath, ''), {
        select: 'ListItemAllFields/ParentList/Id',
        expand: 'ListItemAllFields,ListItemAllFields/ParentList',
      }).catch(() => ({})).then(async (data) => {
        let id = String(
          data?.ListItemAllFields?.ParentList?.Id
          || data?.ListItemAllFields?.ParentList?.ID || '',
        ).replace(/[{}]/g, '').trim();

        // A document-library root has no list item, so ParentList is empty.
        // Resolve that root by its server-relative list URL instead.
        if (!GUID.test(id)) {
          // apiUrl cannot express OData aliases; append the alias to the path.
          const aliasPath = `web/GetList(@listUrl)?@listUrl='${odataPathLiteral(folderPath)}'&$select=Id`;
          const viaUrl = await client.get(aliasPath);
          id = String(viaUrl?.Id || viaUrl?.ID || '').replace(/[{}]/g, '').trim();
        }
        if (!GUID.test(id)) {
          throw new Error('SharePoint did not identify this folder’s document library.');
        }
        return id;
      }).catch((err) => {
        parentListCache.delete(key);
        throw err;
      }));
    }
    return parentListCache.get(key);
  }

  function listFields(listId) {
    if (!fieldsCache.has(listId)) {
      fieldsCache.set(listId, client.getAll(`web/lists(guid'${listId}')/fields`, {
        select: FIELD_SELECT,
      }).then(({ items }) => items).catch((err) => {
        fieldsCache.delete(listId);
        throw err;
      }));
    }
    return fieldsCache.get(listId);
  }

  async function openMetadata(row) {
    metaPanel.hidden = false;
    metaPanel.textContent = '';
    const titleRow = el('div', 'wb-file-meta-head');
    titleRow.append(el('h3', 'wb-subpanel-title', `Metadata for ${row.Name}`));
    const close = el('button', 'btn btn-xs', 'Close');
    close.type = 'button';
    close.addEventListener('click', () => { metaPanel.hidden = true; });
    titleRow.append(close);
    metaPanel.append(titleRow);
    const body = el('div', 'wb-subpanel-body');
    metaPanel.append(body);
    const status = el('div', 'wb-grid-status', 'Loading metadata…');
    body.append(status);

    try {
      const listId = await parentListId(currentPath);
      const fields = await listFields(listId);
      let item = {};
      let itemAsText = {};
      try {
        item = await client.get(fileApi(row.ServerRelativeUrl, '/ListItemAllFields'), {
          expand: 'FieldValuesAsText',
        });
        itemAsText = item.FieldValuesAsText || {};
      } catch {
        try {
          item = await client.get(fileApi(row.ServerRelativeUrl, '/ListItemAllFields'));
        } catch { item = {}; }
      }
      status.remove();
      const form = createFieldEditorForm({
        fields,
        item,
        itemAsText,
        onSave: (formValues) => spWrite.validateUpdateListItem(
          { fileServerRelativeUrl: row.ServerRelativeUrl },
          formValues,
          { newDocumentUpdate: true },
        ),
      });
      body.append(form.el);
    } catch (err) {
      showFailure(status, err, 'this file’s metadata');
    }
  }

  // ---- load ----
  async function load(route) {
    await loadLibraries();
    let path = route?.path;
    if (!path) {
      path = libraries[0]?.RootFolder?.ServerRelativeUrl || webRootPath();
    }
    try {
      await listFolder(path);
    } catch (err) {
      if (!grid) makeGrid();
      grid.setError(err);
    }
  }

  // shell.reset() calls this on every inspected-web change.
  function destroy() {
    crumbObserver?.disconnect();
  }

  return { el: root, load, destroy };
}
