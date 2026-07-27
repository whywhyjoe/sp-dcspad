// Generated BMO SharePoint Design System intelligence for Monaco's CSS and
// HTML surfaces. Runtime code loads only the compact vendor artifact; source
// CSS is parsed by tools/build-design-intelligence.mjs during development.

export const BSP_PACK_ID = 'bsp-design';

let dataPromise = null;

function artifactUrl() {
  const root = window.__DCSPAD_ASSET_BASE__
    || new URL('../../', import.meta.url).href;
  const url = new URL('vendor/intelligence/bsp-design.json', root);
  const version = window.__DCSPAD_INTELLIGENCE_VERSION__;
  if (version) url.searchParams.set('v', version);
  return url.href;
}

function prepareData(raw) {
  if (raw?.schemaVersion !== 1 || raw?.pack !== BSP_PACK_ID
      || !Array.isArray(raw.tokens) || !Array.isArray(raw.classes)) {
    throw new Error('unsupported or malformed bsp-design.json');
  }
  return {
    ...raw,
    tokenByName: new Map(raw.tokens.map((token) => [token.name, token])),
    classByName: new Map(raw.classes.map((item) => [item.name, item])),
  };
}

export function fetchBspIntelligence() {
  if (!dataPromise) {
    dataPromise = fetch(artifactUrl(), {
      credentials: 'same-origin',
      cache: 'no-cache',
    })
      .then((response) => {
        if (!response.ok) throw new Error(`HTTP ${response.status} loading BSP intelligence`);
        return response.json();
      })
      .then(prepareData)
      .catch((error) => {
        dataPromise = null;
        throw error;
      });
  }
  return dataPromise;
}

function markdownToken(token) {
  const scope = token.scope === 'editorial'
    ? '\n\nScoped to **`.editorial`**.'
    : '';
  return [
    `**\`${token.name}\`** · ${token.category}`,
    token.description,
    `Value: \`${token.value}\`${scope}`,
    `Source: \`${token.source.file}:${token.source.line}\``,
  ].filter(Boolean).join('\n\n');
}

function markdownClass(item) {
  const labels = {
    base: 'base/component',
    element: 'BEM element',
    modifier: 'BEM modifier',
    state: 'state',
    utility: 'utility',
  };
  const requirements = item.base
    ? `\n\nCompose with **\`.${item.base}\`**.`
    : '';
  const scope = item.scopes?.includes('editorial')
    ? '\n\nAvailable in **Editorial mode**.'
    : '';
  return [
    `**\`.${item.name}\`** · ${labels[item.kind] || item.kind}`,
    `${item.description}${requirements}${scope}`,
    `Source: \`${item.source.file}:${item.source.line}\``,
  ].filter(Boolean).join('\n\n');
}

function cssTokenAt(model, position) {
  const line = model.getLineContent(position.lineNumber);
  const offset = position.column - 1;
  const pattern = /--[\w-]+/g;
  let match;
  while ((match = pattern.exec(line))) {
    if (offset >= match.index && offset <= match.index + match[0].length) {
      return {
        name: match[0],
        range: {
          startLineNumber: position.lineNumber,
          startColumn: match.index + 1,
          endLineNumber: position.lineNumber,
          endColumn: match.index + match[0].length + 1,
        },
      };
    }
  }
  return null;
}

function cssCompletionContext(model, position) {
  const before = model.getLineContent(position.lineNumber).slice(0, position.column - 1);
  const variable = before.match(/var\(\s*(--[\w-]*)?$/);
  const declaration = before.match(/^\s*(--[\w-]*)$/);
  const prefix = variable ? (variable[1] || '') : declaration?.[1];
  if (prefix === undefined) return null;
  return {
    prefix,
    range: {
      startLineNumber: position.lineNumber,
      startColumn: position.column - prefix.length,
      endLineNumber: position.lineNumber,
      endColumn: position.column,
    },
  };
}

export function createBspCssCompletionProvider(monaco, getData) {
  return {
    triggerCharacters: ['-', '('],
    provideCompletionItems(model, position) {
      const data = getData();
      const context = data && cssCompletionContext(model, position);
      if (!context) return { suggestions: [] };
      return {
        suggestions: data.tokens.map((token) => ({
          label: token.name,
          kind: monaco.languages.CompletionItemKind.Variable,
          detail: `${token.category} · ${token.value}`,
          documentation: { value: markdownToken(token) },
          insertText: token.name,
          filterText: token.name,
          sortText: token.name,
          range: context.range,
        })),
      };
    },
  };
}

export function createBspCssHoverProvider(getData) {
  return {
    provideHover(model, position) {
      const data = getData();
      const target = data && cssTokenAt(model, position);
      const token = target && data.tokenByName.get(target.name);
      if (!token) return null;
      return {
        range: target.range,
        contents: [{ value: markdownToken(token) }],
      };
    },
  };
}

function classAttributeBeforeCursor(model, position) {
  const offset = model.getOffsetAt(position);
  const start = Math.max(0, offset - 6000);
  const before = model.getValue().slice(start, offset);
  const match = before.match(/\bclass\s*=\s*(["'])([^"']*)$/i);
  if (!match) return null;
  const value = match[2];
  const prefix = value.match(/[^\s]*$/)?.[0] || '';
  return {
    value,
    prefix,
    startOffset: offset - prefix.length,
  };
}

function classTokenAt(model, position) {
  const value = model.getValue();
  const offset = model.getOffsetAt(position);
  const before = value.slice(Math.max(0, offset - 6000), offset);
  const open = before.match(/\bclass\s*=\s*(["'])([^"']*)$/i);
  if (!open) return null;
  const quote = open[1];
  const valueStart = offset - open[2].length;
  const close = value.indexOf(quote, offset);
  if (close < 0) return null;
  let start = offset;
  let end = offset;
  while (start > valueStart && !/\s/.test(value[start - 1])) start--;
  while (end < close && !/\s/.test(value[end])) end++;
  const name = value.slice(start, end);
  if (!name) return null;
  const startPosition = model.getPositionAt(start);
  const endPosition = model.getPositionAt(end);
  return {
    name,
    range: {
      startLineNumber: startPosition.lineNumber,
      startColumn: startPosition.column,
      endLineNumber: endPosition.lineNumber,
      endColumn: endPosition.column,
    },
  };
}

export function createBspHtmlClassCompletionProvider(monaco, getData) {
  return {
    triggerCharacters: [' ', '-', '_'],
    provideCompletionItems(model, position) {
      const data = getData();
      const context = data && classAttributeBeforeCursor(model, position);
      if (!context) return { suggestions: [] };
      const present = new Set(context.value.trim().split(/\s+/).filter(Boolean));
      present.delete(context.prefix);
      const startPosition = model.getPositionAt(context.startOffset);
      const range = new monaco.Range(
        startPosition.lineNumber,
        startPosition.column,
        position.lineNumber,
        position.column,
      );
      return {
        suggestions: data.classes
          .filter((item) => !present.has(item.name))
          .map((item) => ({
            label: item.name,
            kind: monaco.languages.CompletionItemKind.Class,
            detail: `BMO design system · ${item.kind}`,
            documentation: { value: markdownClass(item) },
            insertText: item.name,
            filterText: item.name,
            sortText: `${item.kind === 'base' ? '0' : '1'}-${item.name}`,
            range,
          })),
      };
    },
  };
}

export function createBspHtmlClassHoverProvider(getData) {
  return {
    provideHover(model, position) {
      const data = getData();
      const target = data && classTokenAt(model, position);
      const item = target && data.classByName.get(target.name);
      if (!item) return null;
      return {
        range: target.range,
        contents: [{ value: markdownClass(item) }],
      };
    },
  };
}
