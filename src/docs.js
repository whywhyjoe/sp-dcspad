// Same-tenant resource Browser. SharePoint may return .html files as downloads,
// so HTML, Markdown, and text resources are fetched and rendered through srcdoc.
// A base URL preserves relative CSS, images, scripts, and cross-resource links.

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

function prepareHtmlDocument(source, url, { allowScripts = false } = {}) {
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
  if (doc.type === 'markdown' || doc.type === 'text' || doc.type === 'html') return doc.type;
  if (/\.(?:md|markdown)(?:$|[?#])/i.test(doc.url)) return 'markdown';
  if (/\.txt(?:$|[?#])/i.test(doc.url)) return 'text';
  return 'html';
}

function resourceTitle(url) {
  const filename = decodeURIComponent(url.pathname.split('/').pop() || '').trim();
  return filename || 'SharePoint resource';
}

export function initDocs({ config, layoutApi, onError } = {}) {
  const configuredDocs = Array.isArray(config?.docs) ? config.docs : [];
  const copilot = config?.copilot || {};
  const main = document.getElementById('main');
  const menu = document.getElementById('docs-menu');
  const menuItems = document.getElementById('docs-menu-items');
  const menuEmpty = document.getElementById('docs-menu-empty');
  const aiGroup = document.getElementById('docs-ai-group');
  const addressForm = document.getElementById('browser-address-form');
  const addressInput = document.getElementById('browser-address-input');
  let frame = document.getElementById('docs-frame');
  const state = document.getElementById('docs-state');
  const openSource = document.getElementById('btn-docs-open-source');
  const cache = new Map();
  let current = null;
  let loadController = null;

  function wireFrameLinks(targetFrame) {
    targetFrame.addEventListener('load', () => {
      const doc = targetFrame.contentDocument;
      if (!doc) return;
      doc.addEventListener('click', (event) => {
        if (event.defaultPrevented || event.button !== 0
            || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
        const link = event.target.closest?.('a[href]');
        if (!link) return;
        const raw = link.getAttribute('href') || '';
        if (!raw || raw.startsWith('#')) return;
        let url;
        try { url = new URL(raw, doc.baseURI); } catch (_) { return; }
        if (url.origin !== location.origin) {
          event.preventDefault();
          onError?.('Browser links are limited to this SharePoint tenant.');
          return;
        }
        if (/\.(?:html?|md|markdown|txt)(?:$|[?#])/i.test(url.href)) {
          event.preventDefault();
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
        event.preventDefault();
        onError?.('Browser supports same-tenant HTML, Markdown, and text files.');
      }, true);
    }, { once: true });
  }

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
    if (!source) throw new Error('Enter a SharePoint HTML, Markdown, or text URL.');
    const url = new URL(source, location.href);
    if (url.origin !== location.origin) {
      throw new Error(`Browser is limited to ${location.origin}.`);
    }
    if (!/\.(?:html?|md|markdown|txt)$/i.test(url.pathname)) {
      throw new Error('Browser supports .html, .htm, .md, .markdown, and .txt files.');
    }
    return url;
  }

  function loadAddress(value) {
    try {
      const url = normalizeTenantUrl(value);
      const configured = configuredDocs.find((entry) => entry.url === url.href);
      return loadDoc(configured || {
        id: `address:${url.href}`,
        title: resourceTitle(url),
        url: url.href,
        type: 'auto',
      });
    } catch (error) {
      setMode('docs');
      showState(error.message || String(error), 'error');
      onError?.(error.message || String(error));
      return Promise.resolve();
    }
  }

  async function loadDoc(doc) {
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
    openSource.disabled = false;
    openSource.title = `Open ${doc.title} source in a new tab`;
    showState(`Loading ${doc.title}…`, 'loading');
    loadController?.abort();
    loadController = new AbortController();

    try {
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
        ? prepareMarkdownDocument(source, doc.url, doc.title)
        : type === 'text'
          ? prepareTextDocument(source, doc.url, doc.title)
          : prepareHtmlDocument(source, doc.url, { allowScripts });
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
      frame.srcdoc = srcdoc;
      state.hidden = true;
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
    const type = documentType(doc);
    const badge = type === 'markdown' ? 'MD' : type === 'text' ? 'TXT' : 'HTML';
    button.innerHTML = `<span class="docs-menu-icon" aria-hidden="true">${badge}</span>`
      + `<span>${escapeHtml(doc.title)}</span>`;
    button.addEventListener('click', () => loadDoc(doc));
    menuItems.append(button);
  }
  menuEmpty.hidden = configuredDocs.length > 0;
  if (!configuredDocs.length) {
    showState('Paste a same-tenant HTML, Markdown, or text URL in the address bar.');
  }

  aiGroup.hidden = !(copilot.enabled && copilot.url);
  document.getElementById('mi-open-copilot').addEventListener('click', () => {
    if (!copilot.enabled || !copilot.url) return;
    menu.hidden = true;
    document.getElementById('btn-docs').setAttribute('aria-expanded', 'false');
    const opened = window.open(copilot.url, 'dcspad-copilot');
    opened?.focus?.();
  });

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
  openSource.addEventListener('click', () => {
    if (current?.url) window.open(current.url, '_blank', 'noopener,noreferrer');
  });
  document.getElementById('btn-max-docs').addEventListener('click', () => {
    main.classList.remove('max-preview', 'max-diag', 'max-editor');
    main.classList.toggle('max-docs');
  });
  return { loadDoc, loadAddress, setMode };
}
