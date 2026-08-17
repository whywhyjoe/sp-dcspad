// DCSPad-style pre-upload metadata dialog for the Files browser.
//
// Mirrors the pad's "File metadata" dialog (index.html #sp-metadata-dialog)
// in markup classes and behavior: probe the destination library for the
// three curated columns (Title, _ExtendedDescription, DocVersion), grey out
// whatever the library lacks, prefill from the file being overwritten, then
// upload and write metadata. A metadata failure keeps the uploaded file and
// offers a retry that never re-uploads, plus "Keep file without metadata".
//
// metadataFieldStates() is pure and mirrors sp-files.js's availability
// rules; the dialog itself is plain DOM riding app.css's .app-dialog and
// .sp-metadata-* styles (workbench.html loads app.css before workbench.css).

export const FILE_METADATA_SPECS = Object.freeze([
  { key: 'title', label: 'Title', internalName: 'Title', types: ['Text'] },
  { key: 'description', label: 'Description', internalName: '_ExtendedDescription', types: ['Note', 'Text'] },
  { key: 'docVersion', label: 'DocVersion', internalName: 'DocVersion', types: ['Text'] },
]);

export function metadataFieldStates(libraryFields) {
  const fields = Array.isArray(libraryFields) ? libraryFields : [];
  const states = {};
  for (const spec of FILE_METADATA_SPECS) {
    const match = fields.find((f) =>
      String(f.InternalName || '').toLowerCase() === spec.internalName.toLowerCase());
    let reason = '';
    if (!match) reason = `${spec.internalName} is not available in this library.`;
    else if (match.ReadOnlyField) reason = `${spec.internalName} is read-only.`;
    else if (match.Hidden) reason = `${spec.internalName} is hidden in this library.`;
    else if (!spec.types.includes(String(match.TypeAsString || ''))) {
      reason = `${spec.internalName} is not a supported text field.`;
    }
    states[spec.key] = {
      key: spec.key,
      label: spec.label,
      internalName: match?.InternalName || spec.internalName,
      entityPropertyName: match?.EntityPropertyName || match?.InternalName || spec.internalName,
      available: !reason,
      reason,
    };
  }
  return states;
}

export const anyMetadataAvailable = (states) =>
  Object.values(states || {}).some((s) => s.available);

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

// Opens the dialog and drives the upload lifecycle.
//   doUpload:   async () => void — performs the binary upload. A throw
//               closes the dialog and rejects with the error, so the caller
//               keeps its consent handling (409 race included).
//   doMetadata: async (values) => void — writes metadata for the uploaded
//               file. A throw keeps the dialog open with Retry / Keep.
// Resolves 'saved' | 'kept' | 'cancelled'.
export function openUploadMetadataDialog({
  fileName, overwrite = false, states, values = {}, doUpload, doMetadata,
}) {
  return new Promise((resolve, reject) => {
    const dialog = el('dialog', 'app-dialog sp-metadata-dialog wb-upload-metadata');
    const panel = el('div', 'app-dialog__panel');
    const head = el('div', 'app-dialog__head');
    head.append(el('h2', '', 'File metadata'));
    const closeBtn = el('button', 'btn btn-ghost btn-xs', '✕');
    closeBtn.type = 'button';
    closeBtn.setAttribute('aria-label', 'Close');
    head.append(closeBtn);
    const context = el('p', 'app-dialog__context', overwrite
      ? `Review metadata before replacing ${fileName}.`
      : `Add metadata before uploading ${fileName}.`);
    panel.append(head, context);

    const inputs = {};
    for (const state of Object.values(states)) {
      const label = el('label',
        `app-dialog__field sp-metadata-field ${state.available ? 'available' : 'unavailable'}`);
      const headRow = el('span', 'sp-metadata-field__head');
      headRow.append(el('span', '', state.label));
      headRow.append(el('span', 'sp-metadata-field__state',
        state.available ? 'Available' : 'Unavailable'));
      const input = state.key === 'description' ? el('textarea') : el('input');
      if (state.key === 'description') input.rows = 4;
      else {
        input.type = 'text';
        input.maxLength = 255;
        input.autocomplete = 'off';
      }
      input.className = `wb-upload-meta-${state.key}`;
      input.disabled = !state.available;
      input.value = state.available ? String(values[state.key] ?? '') : '';
      const hint = el('span', 'sp-metadata-field__hint', state.available
        ? `Writes to the ${state.internalName} field.`
        : state.reason);
      label.append(headRow, input, hint);
      panel.append(label);
      inputs[state.key] = input;
    }

    const error = el('div', 'sp-files-error');
    error.setAttribute('role', 'alert');
    error.hidden = true;
    const actions = el('div', 'app-dialog__actions sp-metadata-actions');
    const cancel = el('button', 'btn btn-ghost', 'Cancel');
    cancel.type = 'button';
    const keep = el('button', 'btn wb-upload-meta-keep', 'Keep file without metadata');
    keep.type = 'button';
    keep.hidden = true;
    const primary = el('button', 'btn btn-run wb-upload-meta-go', 'Upload file');
    primary.type = 'button';
    actions.append(cancel, keep, primary);
    panel.append(error, actions);
    dialog.append(panel);
    document.body.append(dialog);

    const readValues = () => Object.fromEntries(
      Object.entries(inputs).map(([key, input]) => [key, input.value]));
    const finish = (outcome) => { dialog.close(); dialog.remove(); resolve(outcome); };
    const fail = (err) => { dialog.close(); dialog.remove(); reject(err); };

    let uploaded = false;
    let busy = false;

    async function run() {
      if (busy) return;
      busy = true;
      error.hidden = true;
      primary.disabled = true;
      cancel.disabled = true;
      keep.disabled = true;
      if (!uploaded) {
        primary.textContent = 'Uploading…';
        try {
          await doUpload();
          uploaded = true;
        } catch (err) {
          // Let the caller re-run the flow (409-race consent) without
          // losing what the user already typed.
          if (err && typeof err === 'object') err.uploadMetadataValues = readValues();
          fail(err);
          return;
        }
      }
      primary.textContent = 'Saving metadata…';
      try {
        await doMetadata(readValues());
        finish('saved');
      } catch (err) {
        busy = false;
        cancel.hidden = true;      // the file exists now — cancelling is over
        keep.hidden = false;
        keep.disabled = false;
        primary.textContent = 'Retry metadata';
        primary.disabled = false;
        error.textContent =
          `The file was uploaded, but its metadata could not be saved: ${err?.message || err}`;
        error.hidden = false;
      }
    }

    primary.addEventListener('click', run);
    keep.addEventListener('click', () => { if (!busy) finish('kept'); });
    const dismiss = () => {
      if (busy) return;
      if (uploaded) finish('kept');
      else finish('cancelled');
    };
    cancel.addEventListener('click', dismiss);
    closeBtn.addEventListener('click', dismiss);
    dialog.addEventListener('cancel', (e) => { e.preventDefault(); dismiss(); });

    dialog.showModal();
    Object.values(inputs).find((input) => !input.disabled)?.focus();
  });
}
