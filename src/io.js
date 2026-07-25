// File save/load helpers: turning pad state into files on the user's
// disk and back. Persistence (localStorage) stays in state.js — this
// module only moves bytes through downloads and file pickers.

// A legitimate DCSPad file can't exceed the localStorage quota it came
// from, so anything bigger is a mis-pick (and would lock the main
// thread in JSON.parse).
const MAX_IMPORT_BYTES = 5 * 1024 * 1024;

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
