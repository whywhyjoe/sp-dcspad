// Runner: assembles one complete, correctly-ordered HTML document per run
// and executes it in a fresh same-origin srcdoc iframe.
//
// Assembly order is the contract that makes pad behavior match a real
// SharePoint page: harness first, then SP context + <base>, library CSS,
// user CSS, user HTML, library JS (ordered, blocking), user JS last.

let harnessText = null;
let currentToken = null;
let currentFrame = null;
let runCounter = 0;
let userJsLine = 0;          // 1-based line where user JS starts in the doc
const evalCallbacks = new Map();
let evalCounter = 0;
let handlers = {};

export async function initRunner(messageHandlers) {
  handlers = messageHandlers;
  const res = await fetch(new URL('./bridge/harness.js', import.meta.url));
  harnessText = await res.text();

  window.addEventListener('message', (e) => {
    const d = e.data;
    if (!d || d.dcspad !== currentToken) return;
    if (d.kind === 'eval-result') {
      const cb = evalCallbacks.get(d.id);
      if (cb) { evalCallbacks.delete(d.id); cb(d); }
      return;
    }
    handlers[d.kind]?.(d);
  });
}

const escScript = (s) => s.replace(/<\/script/gi, '<\\/script');
const escStyle = (s) => s.replace(/<\/style/gi, '<\\/style');

function assemble({ docs, libraries, spContext, settings, token }) {
  const cssLinks = libraries
    .filter((l) => l.css)
    .map((l) => (Array.isArray(l.css) ? l.css : [l.css]).map((u) => `<link rel="stylesheet" href="${u}">`).join('\n'))
    .join('\n');
  const jsTags = libraries
    .filter((l) => l.js)
    .map((l) => (Array.isArray(l.js) ? l.js : [l.js]).map((u) => `<script src="${u}"><\/script>`).join('\n'))
    .join('\n');

  const contextScript = spContext
    ? `<script>window._spPageContextInfo = ${JSON.stringify(spContext.pageContext)};<\/script>\n` +
      (spContext.baseHref ? `<base href="${spContext.baseHref}">\n` : '')
    : '';

  const head = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<script>${escScript(harnessText.replaceAll('__DCSPAD_TOKEN__', token))}<\/script>
${contextScript}${cssLinks}
<style>
${escStyle(docs.css)}
</style>
</head>
<body>
${docs.html}
${jsTags}
`;
  const scriptOpen = settings.jsAsModule ? '<script type="module">' : '<script>';
  // User JS begins on the line right after the opening script tag.
  userJsLine = head.split('\n').length + 1;
  return `${head}${scriptOpen}
${escScript(docs.js)}
<\/script>
</body>
</html>`;
}

export function run(opts) {
  runCounter += 1;
  currentToken = `run-${runCounter}-${Math.random().toString(36).slice(2)}`;

  const doc = assemble({ ...opts, token: currentToken });

  const host = document.getElementById('preview-host');
  document.getElementById('preview-empty')?.remove();
  if (currentFrame) currentFrame.remove();
  evalCallbacks.clear();

  const frame = document.createElement('iframe');
  frame.setAttribute('sandbox', 'allow-scripts allow-same-origin allow-forms allow-modals allow-popups');
  frame.srcdoc = doc;
  host.appendChild(frame);
  currentFrame = frame;

  return { runNumber: runCounter, token: currentToken };
}

export function evalInFrame(code) {
  return new Promise((resolve) => {
    if (!currentFrame) {
      resolve({ ok: false, value: { t: 'str', v: 'Nothing is running — press Run first.' }, noRun: true });
      return;
    }
    const id = ++evalCounter;
    evalCallbacks.set(id, resolve);
    currentFrame.contentWindow.postMessage({ dcspad: currentToken, kind: 'eval', code, id }, '*');
  });
}

// Map a line number from the assembled srcdoc document back to the user's
// JS editor (1-based). Returns null for lines outside the user script.
export function mapSrcdocLineToUserJs(line) {
  const mapped = line - userJsLine + 1;
  return mapped >= 1 ? mapped : null;
}

export function hasRun() { return !!currentFrame; }
