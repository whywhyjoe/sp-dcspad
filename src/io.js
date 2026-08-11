// File save/load helpers: turning pad state into files on the user's
// disk and back. Persistence (localStorage) stays in state.js — this
// module only moves bytes through downloads and file pickers.

// A legitimate DCSPad file can't exceed the localStorage quota it came
// from, so anything bigger is a mis-pick (and would lock the main
// thread in JSON.parse).
export const MAX_IMPORT_BYTES = 5 * 1024 * 1024;

export function downloadText(filename, text, type = 'application/json') {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.append(a);
  a.click();
  a.remove();
  // Revoke on a delay: some engines (Safari) abort a download whose
  // blob URL is revoked before the download manager has claimed it.
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// Wire a hidden <input type="file"> to a JSON handler. Oversize files
// and JSON syntax errors are rejected here with their own messages;
// the handler receives only parsed objects and owns shape validation.
// The input is reset afterwards so picking the same file twice fires.
export function wireJsonImport(inputId, onDoc) {
  const input = document.getElementById(inputId);
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    input.value = '';
    if (!file) return;
    if (file.size > MAX_IMPORT_BYTES) {
      alert(`"${file.name}" is ${(file.size / 1048576).toFixed(1)} MB — too large to be a DCSPad file.`);
      return;
    }
    let doc;
    try { doc = JSON.parse(await file.text()); }
    catch { alert(`"${file.name}" isn't valid JSON.`); return; }
    onDoc(doc, file.name);
  });
  return input;
}

export const BUILT_IN_SHAREPOINT_FILE_TYPES = Object.freeze([
  Object.freeze({
    id: 'html',
    label: 'HTML',
    extensions: Object.freeze(['html', 'htm']),
    pane: 'html',
    defaultExtension: 'html',
  }),
  Object.freeze({
    id: 'css',
    label: 'CSS',
    extensions: Object.freeze(['css']),
    pane: 'css',
    defaultExtension: 'css',
  }),
  Object.freeze({
    id: 'javascript',
    label: 'JavaScript',
    extensions: Object.freeze(['js']),
    pane: 'js',
    defaultExtension: 'js',
  }),
]);

export function sharePointFileTypes(additionalTypes = []) {
  const usedExtensions = new Set(
    BUILT_IN_SHAREPOINT_FILE_TYPES.flatMap((type) => type.extensions),
  );
  const types = [...BUILT_IN_SHAREPOINT_FILE_TYPES];

  for (const [index, raw] of additionalTypes.entries()) {
    if (!raw || typeof raw !== 'object') continue;
    const extensions = (Array.isArray(raw.extensions) ? raw.extensions : [])
      .map((extension) => String(extension || '').trim().replace(/^\./, '').toLowerCase())
      .filter((extension) => extension && !usedExtensions.has(extension));
    if (!extensions.length || !['html', 'css', 'js'].includes(raw.pane)) continue;
    for (const extension of extensions) usedExtensions.add(extension);
    types.push(Object.freeze({
      id: `additional-${index}-${extensions[0]}`,
      label: String(raw.label || '').trim() || extensions[0].toUpperCase(),
      extensions: Object.freeze(extensions),
      pane: raw.pane,
      defaultExtension: extensions[0],
    }));
  }

  return types;
}

export function fileTypeForFileName(fileName, additionalTypes = []) {
  const match = /\.([^.]+)$/i.exec(String(fileName || '').trim());
  const extension = match?.[1]?.toLowerCase();
  if (!extension) return null;
  return sharePointFileTypes(additionalTypes)
    .find((type) => type.extensions.includes(extension)) || null;
}

export function paneForFileName(fileName, additionalTypes = []) {
  return fileTypeForFileName(fileName, additionalTypes)?.pane || '';
}

// Wire one picker for all three editor file types. The callback receives a
// fully-read candidate but does not mutate an editor; main.js owns the shared
// local/SharePoint replacement confirmation.
export function wirePaneImport(inputId, onCandidate, onError = () => {}) {
  const input = document.getElementById(inputId);
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    input.value = '';
    if (!file) return;

    const pane = paneForFileName(file.name);
    if (!pane) {
      onError(`"${file.name}" is not an HTML, CSS, or JavaScript file.`);
      return;
    }
    if (file.size > MAX_IMPORT_BYTES) {
      onError(
        `"${file.name}" is ${(file.size / 1048576).toFixed(1)} MB — `
        + 'HTML, CSS, and JavaScript imports are limited to 5 MB.',
      );
      return;
    }

    try {
      await onCandidate({ fileName: file.name, pane, text: await file.text() });
    } catch (error) {
      onError(`"${file.name}" could not be read (${error.message || error}).`);
    }
  });
  return input;
}
