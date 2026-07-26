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
  // no-cache: SharePoint serves library files with max-age=86400; a
  // revalidation (304 when unchanged) keeps the harness in step with the
  // rest of the graph after a deploy. Hosted mode runs from the bundled
  // artifact, where import.meta.url points at the bundle — boot.js provides
  // the real src/ base; standalone (unbundled) resolves relatively.
  const harnessUrl = window.__DCSPAD_SRC_BASE__
    ? window.__DCSPAD_SRC_BASE__ + 'bridge/harness.js'
    : new URL('./bridge/harness.js', import.meta.url);
  const res = await fetch(harnessUrl, { cache: 'no-cache' });
  if (!res.ok) {
    throw new Error(`preview harness failed to load (HTTP ${res.status} for bridge/harness.js) — check the deployed folder structure`);
  }
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

// Host-page CSP nonce. Modern SharePoint pages ship a nonce-based
// script-src (no 'unsafe-inline'), and about:srcdoc documents inherit the
// parent's CSP — so every script the runner assembles must carry the host
// nonce or the preview frame silently runs nothing. The nonce property is
// readable same-origin; on a standalone page there is none and this stays
// '' (attribute omitted, behavior unchanged).
function hostNonce() {
  for (const s of document.scripts) if (s.nonce) return s.nonce;
  return '';
}

function assemble({ docs, libraries, spContext, settings, token }) {
  const nonce = hostNonce();
  const nonceAttr = nonce ? ` nonce="${nonce}"` : '';
  const cssLinks = libraries
    .filter((l) => l.css)
    .map((l) => (Array.isArray(l.css) ? l.css : [l.css]).map((u) => `<link rel="stylesheet" href="${u}">`).join('\n'))
    .join('\n');
  const jsTags = libraries
    .filter((l) => l.js)
    .map((l) => (Array.isArray(l.js) ? l.js : [l.js]).map((u) => `<script src="${u}"${nonceAttr}><\/script>`).join('\n'))
    .join('\n');

  // Pad-only canvas color, injected BEFORE library/user CSS so anything
  // the user styles wins by cascade — same layering a real page gives.
  const chromeStyle = settings.previewDark
    ? `<style data-dcspad-chrome>
:root { color-scheme: dark; }
html { background: #1d2026; color: #d6d9e0; }
</style>\n`
    : '';

  const contextScript = spContext
    ? `<script${nonceAttr}>window._spPageContextInfo = ${JSON.stringify(spContext.pageContext)};<\/script>\n` +
      (spContext.baseHref ? `<base href="${spContext.baseHref}">\n` : '')
    : '';

  const head = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<script${nonceAttr}>${escScript(harnessText.replaceAll('__DCSPAD_TOKEN__', token))}<\/script>
${contextScript}${chromeStyle}${cssLinks}
<style>
${escStyle(docs.css)}
</style>
</head>
<body>
${docs.html}
${jsTags}
`;
  const scriptOpen = settings.jsAsModule ? `<script type="module"${nonceAttr}>` : `<script${nonceAttr}>`;
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
  // The old frame can never answer now — settle its pending REPL evals
  // instead of leaving their promises hanging forever.
  for (const cb of evalCallbacks.values()) {
    cb({ ok: false, cancelled: true, value: { t: 'str', v: '(cancelled — a new run replaced the frame before this settled)' } });
  }
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
