// Same-tenant resource Browser. SharePoint may return .html files as downloads,
// so HTML, Markdown, code, and text resources are fetched and rendered through srcdoc.
// A base URL preserves relative CSS, images, scripts, and cross-resource links.

import { getState, updateNested } from './state.js';

const BROWSER_PATH_RE = /\.(?:html?|md|markdown|txt|css|js|json|csv)$/i;
const BROWSER_LINK_RE = /\.(?:html?|md|markdown|txt|css|js|json|csv)(?:$|[?#])/i;
const BROWSER_NAVIGATION_MESSAGE = 'dcspad:browser-navigate';

const escapeHtml = (value) => String(value)
  .replaceAll('&', '&amp;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;')
  .replaceAll("'", '&#39;');

function safeHref(value, { image = false } = {}) {
  const href = String(value || '').trim();
  if (!href || /^(?:javascript|vbscript):/i.test(href)) return '#';
  if (/^data:/i.test(href) && !(image && /^data:image\//i.test(href))) return '#';
  return href;
}

function renderInline(source) {
  const tokens = [];
  const stash = (html) => {
    const token = `\u0000${tokens.length}\u0000`;
    tokens.push(html);
    return token;
  };

  let value = String(source || '');
  value = value.replace(/`([^`\n]+)`/g, (_, code) =>
    stash(`<code>${escapeHtml(code)}</code>`));
  value = value.replace(/!\[([^\]]*)\]\((\S+?)(?:\s+["']([^"']*)["'])?\)/g,
    (_, alt, href, title) => stash(
      `<img src="${escapeHtml(safeHref(href, { image: true }))}" alt="${escapeHtml(alt)}"`
      + `${title ? ` title="${escapeHtml(title)}"` : ''} loading="lazy">`,
    ));
  value = value.replace(/\[([^\]]+)\]\((\S+?)(?:\s+["']([^"']*)["'])?\)/g,
    (_, label, href, title) => stash(
      `<a href="${escapeHtml(safeHref(href))}"${title ? ` title="${escapeHtml(title)}"` : ''}>`
      + `${escapeHtml(label)}</a>`,
    ));

  value = escapeHtml(value)
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/__([^_]+)__/g, '<strong>$1</strong>')
    .replace(/(^|[\s(])\*([^*\n]+)\*/g, '$1<em>$2</em>')
    .replace(/(^|[\s(])_([^_\n]+)_/g, '$1<em>$2</em>')
    .replace(/~~([^~]+)~~/g, '<del>$1</del>')
    .replace(/ {2}\n/g, '<br>\n');

  return value.replace(/\u0000(\d+)\u0000/g, (_, index) => tokens[Number(index)]);
}

function markdownToHtml(markdown) {
  const lines = String(markdown || '').replace(/\r\n?/g, '\n').split('\n');
  const out = [];
  let paragraph = [];
  const flushParagraph = () => {
    if (!paragraph.length) return;
    out.push(`<p>${renderInline(paragraph.join('\n'))}</p>`);
    paragraph = [];
  };

  for (let index = 0; index < lines.length;) {
    const line = lines[index];
    const fence = line.match(/^\s*```([\w-]*)\s*$/);
    if (fence) {
      flushParagraph();
      const code = [];
      index += 1;
      while (index < lines.length && !/^\s*```\s*$/.test(lines[index])) {
        code.push(lines[index]);
        index += 1;
      }
      if (index < lines.length) index += 1;
      const language = fence[1] ? ` class="language-${escapeHtml(fence[1])}"` : '';
      out.push(`<pre><code${language}>${escapeHtml(code.join('\n'))}</code></pre>`);
      continue;
    }

    if (!line.trim()) {
      flushParagraph();
      index += 1;
      continue;
    }

    const heading = line.match(/^(#{1,6})\s+(.+?)\s*#*\s*$/);
    if (heading) {
      flushParagraph();
      const level = heading[1].length;
      const id = heading[2].toLowerCase().replace(/<[^>]+>/g, '')
        .replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
      out.push(`<h${level}${id ? ` id="${escapeHtml(id)}"` : ''}>${renderInline(heading[2])}</h${level}>`);
      index += 1;
      continue;
    }

    if (/^\s*(?:---+|\*\*\*+|___+)\s*$/.test(line)) {
      flushParagraph();
      out.push('<hr>');
      index += 1;
      continue;
    }

    if (/^\s*>\s?/.test(line)) {
      flushParagraph();
      const quoted = [];
      while (index < lines.length && /^\s*>\s?/.test(lines[index])) {
        quoted.push(lines[index].replace(/^\s*>\s?/, ''));
        index += 1;
      }
      out.push(`<blockquote>${markdownToHtml(quoted.join('\n'))}</blockquote>`);
      continue;
    }

    const list = line.match(/^\s*([-+*]|\d+\.)\s+(.+)$/);
    if (list) {
      flushParagraph();
      const ordered = /\d+\./.test(list[1]);
      const items = [];
      while (index < lines.length) {
        const item = lines[index].match(/^\s*([-+*]|\d+\.)\s+(.+)$/);
        if (!item || (/\d+\./.test(item[1]) !== ordered)) break;
        items.push(`<li>${renderInline(item[2])}</li>`);
        index += 1;
      }
      out.push(`<${ordered ? 'ol' : 'ul'}>${items.join('')}</${ordered ? 'ol' : 'ul'}>`);
      continue;
    }

    const next = lines[index + 1] || '';
    if (line.includes('|') && /^\s*\|?\s*:?-{3,}/.test(next)) {
      flushParagraph();
      const splitCells = (row) => row.trim().replace(/^\||\|$/g, '')
        .split('|').map((cell) => cell.trim());
      const headers = splitCells(line);
      index += 2;
      const rows = [];
      while (index < lines.length && lines[index].includes('|') && lines[index].trim()) {
        rows.push(splitCells(lines[index]));
        index += 1;
      }
      out.push('<div class="table-wrap"><table><thead><tr>'
        + headers.map((cell) => `<th>${renderInline(cell)}</th>`).join('')
        + '</tr></thead><tbody>'
        + rows.map((row) => `<tr>${headers.map((_, cellIndex) =>
          `<td>${renderInline(row[cellIndex] || '')}</td>`).join('')}</tr>`).join('')
        + '</tbody></table></div>');
      continue;
    }

    paragraph.push(line);
    index += 1;
  }
  flushParagraph();
  return out.join('\n');
}

function hostNonce() {
  return document.querySelector('script[nonce]')?.nonce
    || document.querySelector('style[nonce]')?.nonce
    || '';
}

async function inlineSameTenantScripts(doc, documentUrl, signal) {
  const documentOrigin = new URL(documentUrl, location.href).origin;
  const scripts = [...doc.querySelectorAll('script[src]')];
  await Promise.all(scripts.map(async (script) => {
    const type = (script.getAttribute('type') || '').trim().toLowerCase();
    if (type && !/^(?:module|text\/javascript|application\/javascript)$/.test(type)) return;

    let scriptUrl;
    try {
      scriptUrl = new URL(script.getAttribute('src'), documentUrl);
    } catch {
      return;
    }
    if (!/^https?:$/.test(scriptUrl.protocol) || scriptUrl.origin !== documentOrigin) return;

    try {
      const response = await fetch(scriptUrl.href, {
        credentials: 'same-origin',
        cache: 'no-cache',
        signal,
      });
      if (!response.ok) return;
      const source = await response.text();
      script.removeAttribute('src');
      script.removeAttribute('integrity');
      script.removeAttribute('crossorigin');
      script.removeAttribute('referrerpolicy');
      // The serialized srcdoc must not contain a literal closing script tag,
      // even when one appears inside a JavaScript string.
      const embeddableSource = source.replace(/<\/script/gi, '<\\/script');
      script.textContent = `${embeddableSource}\n//# sourceURL=${scriptUrl.href}`;
    } catch (error) {
      if (error.name === 'AbortError') throw error;
      // Leave the original src in place. The iframe can still attempt the
      // browser request when the server already supplies an executable MIME.
    }
  }));
}

function addNavigationBridge(doc) {
  const bridge = doc.createElement('script');
  bridge.dataset.dcspadBrowserBridge = '';
  bridge.textContent = `(function () {
  document.addEventListener('click', function (event) {
    if (event.defaultPrevented || event.button !== 0
        || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
    var path = event.composedPath ? event.composedPath() : [event.target];
    var link = null;
    for (var index = 0; index < path.length; index += 1) {
      if (path[index] && path[index].matches && path[index].matches('a[href]')) {
        link = path[index];
        break;
      }
    }
    if (!link) return;
    var raw = link.getAttribute('href') || '';
    if (!raw) return;
    var href;
    try { href = new URL(raw, document.baseURI).href; } catch (_) { return; }
    event.preventDefault();
    parent.postMessage({ type: '${BROWSER_NAVIGATION_MESSAGE}', href: href }, '*');
  }, true);
}());`;
  (doc.head || doc.documentElement).prepend(bridge);
}

async function prepareHtmlDocument(
  source,
  url,
  { allowScripts = false, signal } = {},
) {
  const doc = new DOMParser().parseFromString(source, 'text/html');
  const nonce = hostNonce();
  let base = doc.querySelector('base');
  if (!base) {
    base = doc.createElement('base');
    (doc.head || doc.documentElement).prepend(base);
  }
  base.href = url;
  if (!doc.querySelector('meta[name="viewport"]')) {
    const viewport = doc.createElement('meta');
    viewport.name = 'viewport';
    viewport.content = 'width=device-width, initial-scale=1';
    (doc.head || doc.documentElement).prepend(viewport);
  }
  if (!allowScripts) {
    for (const script of doc.querySelectorAll('script')) script.remove();
  } else {
    // SharePoint commonly serves library .js files as downloads or
    // application/octet-stream. Fetch same-tenant scripts as text, just as
    // the Browser does for the HTML document, then execute them inline.
    await inlineSameTenantScripts(doc, url, signal);
    // This bridge lives inside the parsed page, so it cannot miss the
    // document's load event. The parent validates both the sender frame and
    // destination before routing the link through loadDoc().
    addNavigationBridge(doc);
  }
  if (nonce) {
    for (const style of doc.querySelectorAll('style')) style.nonce = nonce;
    if (allowScripts) {
      for (const script of doc.querySelectorAll('script')) script.nonce = nonce;
    }
  }
  return `<!doctype html>\n${doc.documentElement.outerHTML}`;
}

function prepareMarkdownDocument(source, url, title) {
  const nonce = hostNonce();
  const nonceAttr = nonce ? ` nonce="${escapeHtml(nonce)}"` : '';
  const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <base href="${escapeHtml(url)}">
  <title>${escapeHtml(title)}</title>
  <style${nonceAttr}>
    :root { color-scheme: dark; font: 15px/1.65 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
    * { box-sizing: border-box; }
    body { max-width: 920px; margin: 0 auto; padding: clamp(24px, 5vw, 64px); color: #d7dbe4; background: #16191f; }
    h1, h2, h3, h4, h5, h6 { color: #f5f7fb; line-height: 1.2; margin: 1.6em 0 .65em; scroll-margin-top: 24px; }
    h1 { margin-top: 0; font-size: clamp(1.8rem, 4vw, 2.6rem); } h2 { font-size: 1.55rem; border-bottom: 1px solid #343a46; padding-bottom: .35em; }
    h3 { font-size: 1.25rem; } p, ul, ol, blockquote, pre, table { margin: 0 0 1.1rem; }
    a { color: #78dcc3; } a:hover { color: #a0ead7; }
    code { font-family: "SFMono-Regular", Consolas, monospace; font-size: .9em; color: #d5f4ec; background: #262c34; border: 1px solid #343d48; border-radius: 4px; padding: .12em .34em; }
    pre { overflow: auto; padding: 16px; background: #0f1217; border: 1px solid #343a46; border-radius: 7px; }
    pre code { padding: 0; color: #e5e9f0; background: none; border: 0; }
    blockquote { margin-left: 0; padding: .15em 1em; color: #aeb5c2; border-left: 3px solid #4ac9aa; }
    img { display: block; max-width: 100%; height: auto; }
    hr { border: 0; border-top: 1px solid #343a46; margin: 2rem 0; }
    .table-wrap { overflow-x: auto; margin-bottom: 1.1rem; }
    table { width: 100%; border-collapse: collapse; } th, td { padding: 8px 10px; text-align: left; border: 1px solid #343a46; } th { background: #22262e; }
  </style>
</head>
<body>
  <main>${markdownToHtml(source)}</main>
</body>
</html>`;
  return prepareHtmlDocument(html, url);
}

function prepareTextDocument(source, url, title) {
  const nonce = hostNonce();
  const nonceAttr = nonce ? ` nonce="${escapeHtml(nonce)}"` : '';
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <base href="${escapeHtml(url)}">
  <title>${escapeHtml(title)}</title>
  <style${nonceAttr}>
    :root { color-scheme: dark; font: 14px/1.65 "SFMono-Regular", Consolas, monospace; }
    * { box-sizing: border-box; }
    body { margin: 0; padding: clamp(20px, 4vw, 48px); color: #d7dbe4; background: #16191f; }
    pre { margin: 0; white-space: pre-wrap; overflow-wrap: anywhere; tab-size: 2; }
  </style>
</head>
<body><pre>${escapeHtml(source)}</pre></body>
</html>`;
}

function documentType(doc) {
  if (/\.(?:txt|css|js|json|csv)(?:$|[?#])/i.test(doc.url)) return 'text';
  if (doc.type === 'markdown' || doc.type === 'text' || doc.type === 'html') return doc.type;
  if (/\.(?:md|markdown)(?:$|[?#])/i.test(doc.url)) return 'markdown';
  return 'html';
}

function resourceBadge(doc) {
  const path = new URL(doc.url, location.href).pathname.toLowerCase();
  if (/\.(?:md|markdown)$/.test(path)) return 'MD';
  if (/\.css$/.test(path)) return 'CSS';
  if (/\.js$/.test(path)) return 'JS';
  if (/\.json$/.test(path)) return 'JSON';
  if (/\.csv$/.test(path)) return 'CSV';
  if (/\.txt$/.test(path)) return 'TXT';
  return 'HTML';
}

function resourceTitle(url) {
  const filename = decodeURIComponent(url.pathname.split('/').pop() || '').trim();
  return filename || 'SharePoint resource';
}

function htmlDocumentTitle(source) {
  try {
    const title = new DOMParser().parseFromString(source, 'text/html')
      .querySelector('title')?.textContent?.replace(/\s+/g, ' ').trim();
    return title || '';
  } catch {
    return '';
  }
}

export function initDocs({ config, layoutApi, onBrowse, onError } = {}) {
  const configuredDocs = Array.isArray(config?.docs) ? config.docs : [];
  const copilot = config?.copilot || {};
  const main = document.getElementById('main');
  const menu = document.getElementById('docs-menu');
  const menuItems = document.getElementById('docs-menu-items');
  const favoritesMenuItems = document.getElementById('favorites-menu-items');
  const menuDivider = document.getElementById('docs-menu-divider');
  const menuEmpty = document.getElementById('docs-menu-empty');

  const addressForm = document.getElementById('browser-address-form');
  const addressInput = document.getElementById('browser-address-input');
  const historySelect = document.getElementById('browser-history');
  const backButton = document.getElementById('browser-back');
  const refreshButton = document.getElementById('browser-refresh');
  const browseButton = document.getElementById('browser-browse');
  const addFavoriteButton = document.getElementById('btn-browser-add-favorite');
  const favoriteDialog = document.getElementById('favorite-name-dialog');
  const favoriteForm = document.getElementById('favorite-name-form');
  const favoriteInput = document.getElementById('favorite-name-input');
  const favoriteContext = document.getElementById('favorite-name-context');
  let frame = document.getElementById('docs-frame');
  const state = document.getElementById('docs-state');
  const cache = new Map();
  let current = null;
  let loadController = null;
  let history = [];
  let favorites = [];
  let pendingFavorite = null;
  let navigationEntries = [];
  let navigationIndex = -1;

  function readFavorites() {
    const values = Array.isArray(getState().settings.browserFavorites)
      ? getState().settings.browserFavorites
      : [];
    const seen = new Set();
    favorites = [];
    for (const value of values) {
      if (!value || typeof value !== 'object') continue;
      try {
        const url = normalizeTenantUrl(value.url);
        if (seen.has(url.href)) continue;
        seen.add(url.href);
        favorites.push({
          title: String(value.title || '').trim().slice(0, 80) || resourceTitle(url),
          url: url.href,
          type: String(value.type || 'auto'),
        });
      } catch { /* Ignore stale or no-longer-supported favorites. */ }
    }
  }

  function persistFavorites() {
    updateNested('settings', {
      browserFavorites: favorites.map((favorite) => ({ ...favorite })),
    });
  }

  function renderFavorites() {
    favoritesMenuItems.replaceChildren();
    for (const favorite of favorites) {
      const row = document.createElement('div');
      row.className = 'favorite-menu-row';
      row.setAttribute('role', 'none');

      const open = document.createElement('button');
      open.type = 'button';
      open.className = 'menu-item docs-menu-item';
      open.setAttribute('role', 'menuitem');
      open.textContent = favorite.title;
      open.title = favorite.url;
      open.addEventListener('click', () => loadDoc({
        id: `favorite:${favorite.url}`,
        ...favorite,
      }));

      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'favorite-menu-remove';
      remove.setAttribute('aria-label', `Remove ${favorite.title} from favorites`);
      remove.title = 'Remove from favorites';
      remove.innerHTML = '<svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" aria-hidden="true"><path d="m4 4 8 8M12 4l-8 8"/></svg>';
      remove.addEventListener('click', (event) => {
        event.stopPropagation();
        if (!confirm(`Remove "${favorite.title}" from favorites?`)) return;
        favorites = favorites.filter((item) => item.url !== favorite.url);
        persistFavorites();
        renderFavorites();
      });

      row.append(open, remove);
      favoritesMenuItems.append(row);
    }
    menuDivider.hidden = !(configuredDocs.length && favorites.length);
    menuEmpty.hidden = configuredDocs.length + favorites.length > 0;
  }

  function closeFavoriteDialog() {
    pendingFavorite = null;
    if (favoriteDialog.open) favoriteDialog.close();
  }

  function readHistory() {
    const values = Array.isArray(getState().settings.browserHistory)
      ? getState().settings.browserHistory
      : [];
    const seen = new Set();
    history = [];
    for (const value of values) {
      try {
        const href = normalizeTenantUrl(value).href;
        if (seen.has(href)) continue;
        seen.add(href);
        history.push(href);
      } catch { /* Ignore stale or no-longer-supported history entries. */ }
      if (history.length === 10) break;
    }
  }

  function renderHistory() {
    historySelect.replaceChildren();
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = history.length ? 'Recent URLs' : 'No recent URLs';
    historySelect.append(placeholder);
    for (const url of history) {
      const option = document.createElement('option');
      option.value = url;
      option.textContent = url;
      historySelect.append(option);
    }
    historySelect.value = '';
    historySelect.disabled = history.length === 0;
  }

  function recordHistory(url) {
    history = [url, ...history.filter((item) => item !== url)].slice(0, 10);
    updateNested('settings', { browserHistory: [...history] });
    renderHistory();
  }

  function updateBackButton() {
    backButton.disabled = navigationIndex <= 0;
  }

  function commitNavigation(doc) {
    const active = navigationEntries[navigationIndex];
    if (active?.url === doc.url) {
      navigationEntries[navigationIndex] = { ...doc };
      updateBackButton();
      return;
    }
    navigationEntries = navigationEntries.slice(0, navigationIndex + 1);
    navigationEntries.push({ ...doc });
    navigationIndex = navigationEntries.length - 1;
    updateBackButton();
  }

  function followPageFragment(url) {
    if (!current?.url || !url.hash) return false;
    const loadedUrl = new URL(current.url, location.href);
    const targetUrl = new URL(url.href);
    loadedUrl.hash = '';
    targetUrl.hash = '';
    if (loadedUrl.href !== targetUrl.href) return false;

    let fragment = url.hash.slice(1);
    try { fragment = decodeURIComponent(fragment); } catch { /* Keep the encoded fragment. */ }
    const doc = frame.contentDocument;
    const target = doc?.getElementById(fragment) || doc?.getElementsByName(fragment)?.[0];
    target?.scrollIntoView();
    current = { ...current, url: url.href };
    addressInput.value = url.href;
    return true;
  }

  function followBrowserLink(url) {
    if (url.origin !== location.origin) {
      onError?.('Browser links are limited to this SharePoint tenant.');
      return;
    }
    if (followPageFragment(url)) return;
    if (BROWSER_LINK_RE.test(url.href)) {
      const configured = configuredDocs.find((entry) => entry.url === url.href);
      const title = configured?.title || resourceTitle(url);
      loadDoc(configured || {
        id: `linked:${url.href}`,
        title,
        url: url.href,
        type: 'auto',
      });
      return;
    }
    onError?.('Browser supports same-tenant HTML, Markdown, code, and text files.');
  }

  function wireFrameLinks(targetFrame) {
    const wiredDocuments = new WeakSet();
    const wireCurrentDocument = () => {
      const doc = targetFrame.contentDocument;
      if (!doc || wiredDocuments.has(doc)) return;
      wiredDocuments.add(doc);
      doc.addEventListener('click', (event) => {
        if (event.defaultPrevented || event.button !== 0
            || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
        const link = event.target?.closest?.('a[href]')
          || event.target?.parentElement?.closest?.('a[href]');
        if (!link) return;
        const raw = link.getAttribute('href') || '';
        if (!raw) return;
        let url;
        try { url = new URL(raw, doc.baseURI); } catch (_) { return; }
        event.preventDefault();
        followBrowserLink(url);
      }, true);
    };
    // A newly inserted iframe can emit an initial about:blank load before its
    // srcdoc navigation. Wire every distinct document rather than consuming a
    // one-shot listener on whichever load happens first.
    targetFrame.addEventListener('load', wireCurrentDocument);
    return wireCurrentDocument;
  }

  window.addEventListener('message', (event) => {
    if (event.source !== frame.contentWindow
        || event.data?.type !== BROWSER_NAVIGATION_MESSAGE
        || typeof event.data.href !== 'string') return;
    let url;
    try { url = new URL(event.data.href, location.href); } catch (_) { return; }
    followBrowserLink(url);
  });

  function setMode(name) {
    for (const tab of document.querySelectorAll('#extras-tabs .extras-tab')) {
      const active = tab.dataset.extra === name;
      tab.classList.toggle('active', active);
      tab.setAttribute('aria-selected', String(active));
    }
    for (const view of document.querySelectorAll('#sidebar > .extras-view')) {
      const active = view.id === `extras-${name}`;
      view.hidden = !active;
      view.classList.toggle('active', active);
    }
    if (name === 'docs') layoutApi?.setPaneVisible?.('resources', true);
    else main.classList.remove('max-docs');
  }

  function showState(message, tone = '') {
    state.textContent = message;
    state.className = `docs-state${tone ? ` ${tone}` : ''}`;
    state.hidden = false;
    frame.hidden = true;
  }

  function normalizeTenantUrl(value) {
    const source = String(value || '').trim();
    if (!source) throw new Error('Enter a SharePoint resource URL.');
    const url = new URL(source, location.href);
    if (url.origin !== location.origin) {
      throw new Error(`Browser is limited to ${location.origin}.`);
    }
    if (!BROWSER_PATH_RE.test(url.pathname)) {
      throw new Error('Browser supports .html, .htm, .md, .markdown, .css, .js, .json, .csv, and .txt files.');
    }
    return url;
  }

  function loadAddress(value, options) {
    try {
      const url = normalizeTenantUrl(value);
      const configured = configuredDocs.find((entry) => entry.url === url.href);
      return loadDoc(configured || {
        id: `address:${url.href}`,
        title: resourceTitle(url),
        url: url.href,
        type: 'auto',
      }, options);
    } catch (error) {
      setMode('docs');
      showState(error.message || String(error), 'error');
      onError?.(error.message || String(error));
      return Promise.resolve();
    }
  }

  async function loadDoc(doc, { force = false, record = true, track = true } = {}) {
    if (!doc?.url) return;
    let url;
    try {
      url = normalizeTenantUrl(doc.url);
    } catch (error) {
      setMode('docs');
      showState(error.message || String(error), 'error');
      onError?.(error.message || String(error));
      return;
    }
    doc = { ...doc, url: url.href };
    current = doc;
    setMode('docs');
    addressInput.value = doc.url;
    menu.hidden = true;
    document.getElementById('btn-docs').setAttribute('aria-expanded', 'false');
    refreshButton.disabled = false;
    addFavoriteButton.disabled = true;
    showState(`Loading ${doc.title}…`, 'loading');
    loadController?.abort();
    loadController = new AbortController();

    try {
      if (force) cache.delete(doc.url);
      let source = cache.get(doc.url);
      if (source === undefined) {
        const response = await fetch(doc.url, {
          credentials: 'same-origin',
          cache: 'no-cache',
          signal: loadController.signal,
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        source = await response.text();
        cache.set(doc.url, source);
      }
      const type = documentType(doc);
      const allowScripts = type === 'html';
      const srcdoc = type === 'markdown'
        ? await prepareMarkdownDocument(source, doc.url, doc.title)
        : type === 'text'
          ? prepareTextDocument(source, doc.url, doc.title)
          : await prepareHtmlDocument(source, doc.url, {
            allowScripts,
            signal: loadController.signal,
          });
      // Use a fresh frame for every document. Replacing srcdoc in an existing
      // sandboxed frame can leave Chromium with a zero-size child viewport.
      const nextFrame = frame.cloneNode(false);
      nextFrame.hidden = false;
      nextFrame.title = doc.title;
      const sandbox = [
        'allow-same-origin',
        'allow-forms',
        'allow-modals',
        'allow-popups',
        'allow-popups-to-escape-sandbox',
        'allow-downloads',
      ];
      if (allowScripts) sandbox.push('allow-scripts');
      nextFrame.setAttribute('sandbox', sandbox.join(' '));
      wireFrameLinks(nextFrame);
      frame.replaceWith(nextFrame);
      frame = nextFrame;
      // Assign after insertion. The persistent load hook above wires both a
      // possible initial about:blank document and the final srcdoc document.
      frame.srcdoc = srcdoc;
      state.hidden = true;
      current = {
        ...doc,
        favoriteTitle: type === 'html'
          ? htmlDocumentTitle(source) || resourceTitle(url)
          : resourceTitle(url),
      };
      if (track) commitNavigation(current);
      addFavoriteButton.disabled = false;
      if (record) recordHistory(doc.url);
    } catch (error) {
      if (error.name === 'AbortError') return;
      showState(`Could not load ${doc.title}: ${error.message || error}`, 'error');
      onError?.(`Could not load ${doc.title}.`);
    }
  }

  for (const doc of configuredDocs) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'menu-item docs-menu-item';
    button.setAttribute('role', 'menuitem');
    button.dataset.docId = doc.id;
    button.textContent = doc.title;
    button.addEventListener('click', () => loadDoc(doc));
    menuItems.append(button);
  }
  if (!configuredDocs.length) {
    showState('Paste a same-tenant HTML, Markdown, code, or text URL in the address bar.');
  }

  const copilotBtn = document.getElementById('btn-copilot');
  if (copilotBtn) {
    copilotBtn.hidden = !(copilot.enabled && copilot.url);
    copilotBtn.addEventListener('click', () => {
      if (!copilot.enabled || !copilot.url) return;
      const opened = window.open(copilot.url, 'dcspad-copilot');
      opened?.focus?.();
    });
  }

  document.getElementById('extras-tabs').addEventListener('click', (event) => {
    const tab = event.target.closest('.extras-tab');
    if (!tab) return;
    setMode(tab.dataset.extra);
    if (tab.dataset.extra === 'docs' && !current && configuredDocs[0]) {
      loadDoc(configuredDocs[0]);
    }
  });
  addressForm.addEventListener('submit', (event) => {
    event.preventDefault();
    loadAddress(addressInput.value);
  });
  historySelect.addEventListener('change', () => {
    const url = historySelect.value;
    historySelect.value = '';
    if (url) loadAddress(url);
  });
  backButton.addEventListener('click', () => {
    if (navigationIndex <= 0) return;
    navigationIndex -= 1;
    updateBackButton();
    loadDoc(navigationEntries[navigationIndex], { record: false, track: false });
  });
  refreshButton.addEventListener('click', () => {
    if (current) loadDoc(current, { force: true, record: false, track: false });
  });
  browseButton.addEventListener('click', () => onBrowse?.());
  addFavoriteButton.addEventListener('click', () => {
    if (!current?.url || addFavoriteButton.disabled) return;
    const existing = favorites.find((favorite) => favorite.url === current.url);
    pendingFavorite = {
      title: existing?.title || current.favoriteTitle || resourceTitle(new URL(current.url)),
      url: current.url,
      type: current.type || 'auto',
    };
    favoriteInput.value = pendingFavorite.title;
    favoriteContext.textContent = current.url;
    document.getElementById('favorite-name-title').textContent = existing
      ? 'Update favorite'
      : 'Add to favorites';
    document.getElementById('favorite-name-save').textContent = existing
      ? 'Update favorite'
      : 'Save favorite';
    if (!favoriteDialog.open) favoriteDialog.showModal();
    requestAnimationFrame(() => {
      favoriteInput.focus();
      favoriteInput.select();
    });
  });
  favoriteForm.addEventListener('submit', (event) => {
    event.preventDefault();
    const title = favoriteInput.value.trim();
    if (!pendingFavorite || !title) return;
    const favorite = { ...pendingFavorite, title };
    const existingIndex = favorites.findIndex((item) => item.url === favorite.url);
    if (existingIndex >= 0) favorites[existingIndex] = favorite;
    else favorites.push(favorite);
    persistFavorites();
    renderFavorites();
    pendingFavorite = null;
    favoriteDialog.close();
  });
  document.getElementById('favorite-name-cancel').addEventListener('click', closeFavoriteDialog);
  document.getElementById('favorite-name-close').addEventListener('click', closeFavoriteDialog);
  favoriteDialog.addEventListener('cancel', () => { pendingFavorite = null; });
  document.getElementById('btn-max-docs').addEventListener('click', () => {
    main.classList.remove('max-preview', 'max-diag', 'max-editor');
    main.classList.toggle('max-docs');
  });
  readFavorites();
  renderFavorites();
  readHistory();
  renderHistory();

  return {
    loadDoc,
    loadAddress,
    refresh: () => current && loadDoc(current, {
      force: true,
      record: false,
      track: false,
    }),
    setMode,
  };
}
