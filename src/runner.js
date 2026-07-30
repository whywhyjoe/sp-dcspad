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
const escAttr = (s) => String(s)
  .replaceAll('&', '&amp;')
  .replaceAll('"', '&quot;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;');

function externalScript(url, nonceAttr) {
  return `<script src="${escAttr(url)}"${nonceAttr}><\/script>`;
}

// A fallback is inserted with document.write while the srcdoc parser is still
// at the failed library's position. That makes the fallback parser-blocking,
// preserving the catalog's dependency order before later libraries/user JS.
// The global probe avoids inline event handlers, which SharePoint's CSP blocks.
function libraryScript(entry, nonceAttr) {
  const primary = externalScript(entry.js, nonceAttr);
  if (!entry.fallbackJs || !entry.probeGlobal) return primary;

  const path = entry.probeGlobal.split('.').filter(Boolean);
  const fallbackTag = externalScript(entry.fallbackJs, nonceAttr);
  const message = `DCSPad: ${entry.name || entry.probeGlobal} did not expose ${entry.probeGlobal}; loading configured fallback`;
  const probe = `(function(){var value=window;var path=${JSON.stringify(path)};`
    + `for(var i=0;i<path.length&&value!=null;i+=1)value=value[path[i]];`
    + `if(value==null){console.warn(${JSON.stringify(message)});`
    + `document.write(${JSON.stringify(fallbackTag)});}})();`;
  return `${primary}\n<script${nonceAttr}>${escScript(probe)}<\/script>`;
}

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
    .map((l) => (Array.isArray(l.css) ? l.css : [l.css]).map((u) => `<link rel="stylesheet" href="${escAttr(u)}">`).join('\n'))
    .join('\n');
  const libraryStyles = libraries
    .filter((l) => l.cssText)
    .map((l) => `<style data-dcspad-library="${escAttr(l.name || 'configured')}">
${escStyle(l.cssText)}
</style>`)
    .join('\n');
  const jsTags = libraries
    .filter((l) => l.js)
    .map((l) => {
      if (Array.isArray(l.js)) {
        return l.js.map((u) => externalScript(u, nonceAttr)).join('\n');
      }
      return libraryScript(l, nonceAttr);
    })
    .join('\n');

  // Pad-only canvas color, injected BEFORE library/user CSS so anything
  // the user styles wins by cascade — same layering a real page gives.
  const chromeStyle = settings.previewDark
    ? `<style data-dcspad-chrome>
:root { color-scheme: dark; }
html { background: #1a1d23; color: #e6e9ef; }
</style>\n`
    : '';

  const contextScript = spContext
    ? `<script${nonceAttr}>window._spPageContextInfo = ${JSON.stringify(spContext.pageContext)};<\/script>\n` +
      (spContext.baseHref ? `<base href="${escAttr(spContext.baseHref)}">\n` : '')
    : '';

  const head = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<script${nonceAttr}>${escScript(harnessText.replaceAll('__DCSPAD_TOKEN__', token))}<\/script>
${contextScript}${chromeStyle}${cssLinks}
${libraryStyles}
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

export function reset() {
  currentToken = null;
  if (currentFrame) currentFrame.remove();
  currentFrame = null;
  for (const cb of evalCallbacks.values()) {
    cb({ ok: false, cancelled: true, value: { t: 'str', v: '(cancelled — the project was reset)' } });
  }
  evalCallbacks.clear();
  const host = document.getElementById('preview-host');
  host.replaceChildren();
  const empty = document.createElement('div');
  empty.id = 'preview-empty';
  empty.className = 'preview-empty';
  empty.innerHTML = 'Nothing to preview yet — press <kbd>Ctrl/Cmd + Enter</kbd>';
  host.append(empty);
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
