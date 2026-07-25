// File save/load helpers: turning pad state into files on the user's
// disk and back. Persistence (localStorage) stays in state.js — this
// module only moves bytes through downloads and file pickers.

export function downloadText(filename, text, type = 'application/json') {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.append(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// Wire a hidden <input type="file"> to a JSON handler. The handler gets
// the parsed object, or null when the file wasn't valid JSON — it owns
// shape validation and user feedback. The input is reset afterwards so
// picking the same file twice still fires.
export function wireJsonImport(inputId, onDoc) {
  const input = document.getElementById(inputId);
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    input.value = '';
    if (!file) return;
    let doc = null;
    try { doc = JSON.parse(await file.text()); } catch { /* handler shows the error */ }
    onDoc(doc, file.name);
  });
  return input;
}
