// Full Microsoft Fluent System Icons intelligence for Monaco's HTML model.
// The generated artifact is a compact projection of the vendored catalog;
// no SVG folders, font CSS, or multi-megabyte source JSON are parsed here.

export const FLUENT_ICONS_PACK_ID = 'fluent-icons';

export const FLUENT_ICONS_HTML_DATA = {
  version: 1.1,
  tags: [{
    name: 'fluent-icon',
    description: [
      'Font-backed Fluent System Icon. The name uses',
      '`{icon-name}-{source-size}-{regular|filled|light}`.',
      'Decorative by default; add `label` when the icon itself conveys meaning.',
    ].join(' '),
    attributes: [
      {
        name: 'name',
        description: 'A real Fluent icon token, for example `home-24-regular`.',
      },
      {
        name: 'label',
        description: 'Accessible label for a meaningful icon. Omit for decorative icons.',
      },
    ],
  }],
};

let dataPromise = null;

function artifactUrl() {
  const root = window.__DCSPAD_ASSET_BASE__
    || new URL('../../', import.meta.url).href;
  const url = new URL('vendor/intelligence/fluent-icons.json', root);
  const version = window.__DCSPAD_INTELLIGENCE_VERSION__;
  if (version) url.searchParams.set('v', version);
  return url.href;
}

function variantPreference(variant) {
  return (variant.size === 24 ? 0 : variant.size === 20 ? 1 : 2)
    + (variant.style === 'regular' ? 0
      : variant.style === 'filled' ? 0.1
        : variant.style === 'light' ? 0.2 : 0.3);
}

function prepareData(raw) {
  if (raw?.schemaVersion !== 1 || raw?.pack !== FLUENT_ICONS_PACK_ID
      || !Array.isArray(raw.icons)) {
    throw new Error('unsupported or malformed fluent-icons.json');
  }

  const variants = [];
  const variantByToken = new Map();
  const variantById = new Map();
  const variantByClass = new Map();
  const defaultVariants = [];
  const defaultFontVariants = [];
  const fontVariants = [];

  for (const icon of raw.icons) {
    if (!icon?.slug || !icon?.idBase || !Array.isArray(icon.variants)) continue;
    const iconVariants = [];
    const svgOnly = new Set(icon.svgOnly || []);
    for (const suffix of icon.variants) {
      const match = String(suffix).match(
        /^(\d+)-(regular|filled|light|color)(?:-(ltr|rtl))?$/,
      );
      if (!match) continue;
      const size = Number(match[1]);
      const style = match[2];
      const direction = match[3] || '';
      const token = `${icon.slug}-${size}-${style}`;
      const directionalToken = direction ? `${token}-${direction}` : token;
      const id = `ic_fluent_${icon.idBase}_${size}_${style}${direction ? `_${direction}` : ''}`;
      const className = `icon-${id}`;
      const variant = {
        name: icon.name,
        slug: icon.slug,
        description: icon.description || '',
        metaphors: icon.metaphors || [],
        size,
        style,
        direction,
        token: directionalToken,
        id,
        className,
        filename: `${id}.svg`,
        fontAvailable: !svgOnly.has(suffix),
      };
      variant.searchText = [
        icon.name,
        icon.slug,
        icon.description,
        ...(icon.metaphors || []),
        directionalToken,
      ].filter(Boolean).join(' ').toLowerCase();
      variants.push(variant);
      iconVariants.push(variant);
      variantByToken.set(directionalToken, variant);
      variantById.set(id, variant);
      if (variant.fontAvailable) {
        fontVariants.push(variant);
        variantByClass.set(className, variant);
      }
    }
    iconVariants.sort((a, b) =>
      variantPreference(a) - variantPreference(b)
      || a.token.localeCompare(b.token));
    if (iconVariants[0]) defaultVariants.push(iconVariants[0]);
    const defaultFont = iconVariants.find((variant) => variant.fontAvailable);
    if (defaultFont) defaultFontVariants.push(defaultFont);
  }

  defaultVariants.sort((a, b) => a.name.localeCompare(b.name));
  defaultFontVariants.sort((a, b) => a.name.localeCompare(b.name));
  return {
    ...raw,
    variants,
    fontVariants,
    defaultVariants,
    defaultFontVariants,
    variantByToken,
    variantById,
    variantByClass,
  };
}

export function fetchFluentIconIntelligence() {
  if (!dataPromise) {
    dataPromise = fetch(artifactUrl(), {
      credentials: 'same-origin',
      cache: 'no-cache',
    })
      .then((response) => {
        if (!response.ok) throw new Error(`HTTP ${response.status} loading Fluent icon intelligence`);
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

function markdownVariant(variant) {
  const summary = variant.description
    || (variant.metaphors.length
      ? `Related concepts: ${variant.metaphors.slice(0, 8).join(', ')}.`
      : 'Microsoft Fluent System Icon.');
  const fontMarkup = variant.fontAvailable
    ? [
      '**Fluent icon font**',
      `\`<i class="${variant.className}" aria-hidden="true"></i>\``,
      'Keep the `<i>` decorative and label its parent control.',
    ]
    : [
      '**SVG only**',
      'This color/direction-specific asset is not available in the configured icon fonts.',
    ];
  return [
    `**${variant.name}** · ${variant.size}px ${variant.style}`,
    summary,
    `Token: \`${variant.token}\``,
    `Font class: \`${variant.className}\``,
    `SVG file: \`${variant.filename}\``,
    '**Custom element**',
    `\`<fluent-icon name="${variant.token}"></fluent-icon>\``,
    '**Fluent SVG sprite**',
    `\`<svg class="icon"><use href="#${variant.id}"></use></svg>\``,
    'The `<use>` form requires a Fluent sprite containing that symbol id.',
    ...fontMarkup,
  ].join('\n\n');
}

function completionRange(monaco, model, position, startOffset) {
  const start = model.getPositionAt(startOffset);
  return new monaco.Range(
    start.lineNumber,
    start.column,
    position.lineNumber,
    position.column,
  );
}

function currentTagFragment(model, position) {
  const offset = model.getOffsetAt(position);
  const start = Math.max(0, offset - 12000);
  const before = model.getValue().slice(start, offset);
  const lastOpen = before.lastIndexOf('<');
  const lastClose = before.lastIndexOf('>');
  if (lastOpen <= lastClose) return null;
  return {
    offset,
    fragment: before.slice(lastOpen),
  };
}

function nameAttributeContext(model, position) {
  const tag = currentTagFragment(model, position);
  if (!tag || !/^<fluent-icon\b/i.test(tag.fragment)) return null;
  const match = tag.fragment.match(/\bname\s*=\s*(["'])([^"']*)$/i);
  if (!match) return null;
  return {
    kind: 'name',
    prefix: match[2],
    startOffset: tag.offset - match[2].length,
  };
}

function useHrefContext(model, position) {
  const tag = currentTagFragment(model, position);
  if (!tag || !/^<use\b/i.test(tag.fragment)) return null;
  const match = tag.fragment.match(/\bhref\s*=\s*(["'])([^"']*)$/i);
  if (!match) return null;
  const hash = match[2].lastIndexOf('#');
  if (hash < 0) return null;
  const prefix = match[2].slice(hash + 1);
  return {
    kind: 'use',
    prefix,
    startOffset: tag.offset - prefix.length,
  };
}

function fontClassContext(model, position) {
  const tag = currentTagFragment(model, position);
  if (!tag) return null;
  const match = tag.fragment.match(/\bclass\s*=\s*(["'])([^"']*)$/i);
  if (!match) return null;
  const prefix = match[2].match(/[^\s]*$/)?.[0] || '';
  if (!prefix.startsWith('icon-')) return null;
  return {
    kind: 'class',
    prefix,
    startOffset: tag.offset - prefix.length,
  };
}

function normalizedQuery(context) {
  if (context.kind === 'class') {
    return context.prefix
      .replace(/^icon-ic_fluent_/, '')
      .replaceAll('_', '-')
      .toLowerCase();
  }
  if (context.kind === 'use') {
    return context.prefix
      .replace(/^ic_fluent_/, '')
      .replaceAll('_', '-')
      .toLowerCase();
  }
  return context.prefix.toLowerCase().replaceAll('_', '-');
}

function matchingVariants(data, context) {
  const query = normalizedQuery(context);
  const supportsSvgOnly = context.kind === 'use';
  const source = query
    ? (supportsSvgOnly ? data.variants : data.fontVariants)
    : (supportsSvgOnly ? data.defaultVariants : data.defaultFontVariants);
  const ranked = [];
  for (const variant of source) {
    const token = variant.token.toLowerCase();
    let score = 9;
    if (!query) score = 3;
    else if (token.startsWith(query)) score = 0;
    else if (token.includes(query)) score = 1;
    else if (variant.searchText.includes(query.replaceAll('-', ' '))) score = 2;
    if (score < 9) ranked.push({ variant, score });
  }
  ranked.sort((a, b) =>
    a.score - b.score
    || variantPreference(a.variant) - variantPreference(b.variant)
    || a.variant.token.localeCompare(b.variant.token));
  return ranked.slice(0, 600);
}

export function createFluentIconCompletionProvider(monaco, getData) {
  return {
    triggerCharacters: ['"', "'", '-', '_', '#'],
    provideCompletionItems(model, position) {
      const data = getData();
      if (!data) return { suggestions: [] };
      const context = nameAttributeContext(model, position)
        || useHrefContext(model, position)
        || fontClassContext(model, position);
      if (!context) return { suggestions: [] };
      const range = completionRange(monaco, model, position, context.startOffset);
      return {
        suggestions: matchingVariants(data, context).map(({ variant, score }) => {
          const insertText = context.kind === 'class'
            ? variant.className
            : context.kind === 'use'
              ? variant.id
              : variant.token;
          return {
            label: insertText,
            kind: context.kind === 'class'
              ? monaco.languages.CompletionItemKind.Class
              : monaco.languages.CompletionItemKind.EnumMember,
            detail: `${variant.name} · ${variant.size}px ${variant.style}`,
            documentation: { value: markdownVariant(variant) },
            insertText,
            filterText: `${insertText} ${variant.searchText}`,
            sortText: `${score}-${String(variantPreference(variant)).padStart(3, '0')}-${variant.token}`,
            range,
          };
        }),
      };
    },
  };
}

function wordAt(model, position) {
  const value = model.getValue();
  const offset = model.getOffsetAt(position);
  let start = offset;
  let end = offset;
  while (start > 0 && /[\w-]/.test(value[start - 1])) start--;
  while (end < value.length && /[\w-]/.test(value[end])) end++;
  if (start === end) return null;
  const startPosition = model.getPositionAt(start);
  const endPosition = model.getPositionAt(end);
  return {
    value: value.slice(start, end),
    range: {
      startLineNumber: startPosition.lineNumber,
      startColumn: startPosition.column,
      endLineNumber: endPosition.lineNumber,
      endColumn: endPosition.column,
    },
  };
}

export function createFluentIconHoverProvider(getData) {
  return {
    provideHover(model, position) {
      const data = getData();
      const target = data && wordAt(model, position);
      if (!target) return null;
      const variant = data.variantByToken.get(target.value)
        || data.variantById.get(target.value)
        || data.variantByClass.get(target.value);
      if (!variant) return null;
      return {
        range: target.range,
        contents: [{ value: markdownVariant(variant) }],
      };
    },
  };
}

function markerForValue(monaco, model, startOffset, value, message) {
  const start = model.getPositionAt(startOffset);
  const end = model.getPositionAt(startOffset + value.length);
  return {
    severity: monaco.MarkerSeverity.Warning,
    message,
    startLineNumber: start.lineNumber,
    startColumn: start.column,
    endLineNumber: end.lineNumber,
    endColumn: end.column,
  };
}

function attributeValues(source, tagName, attributeName) {
  const results = [];
  const pattern = new RegExp(
    `<${tagName}\\b[^>]*\\b${attributeName}\\s*=\\s*([\"'])([^\"']+)\\1`,
    'gi',
  );
  let match;
  while ((match = pattern.exec(source))) {
    const value = match[2];
    const relative = match[0].lastIndexOf(value);
    results.push({ value, start: match.index + relative });
  }
  return results;
}

export function collectFluentIconMarkers(monaco, model, data) {
  if (!data) return [];
  const source = model.getValue();
  const markers = [];

  for (const item of attributeValues(source, 'fluent-icon', 'name')) {
    const variant = data.variantByToken.get(item.value);
    if (!variant) {
      markers.push(markerForValue(
        monaco,
        model,
        item.start,
        item.value,
        `Unknown Fluent icon token "${item.value}". Use completion to select a real size/style variant.`,
      ));
    } else if (!variant.fontAvailable) {
      markers.push(markerForValue(
        monaco,
        model,
        item.start,
        item.value,
        `"${item.value}" is SVG-only and cannot render through the configured font-backed <fluent-icon>. Use a Fluent sprite <use> reference.`,
      ));
    }
  }

  for (const item of attributeValues(source, 'use', 'href')) {
    const hash = item.value.lastIndexOf('#');
    const id = hash >= 0 ? item.value.slice(hash + 1) : '';
    if (id.startsWith('ic_fluent_') && !data.variantById.has(id)) {
      markers.push(markerForValue(
        monaco,
        model,
        item.start + hash + 1,
        id,
        `Unknown Fluent sprite symbol "${id}".`,
      ));
    }
  }

  const classPattern = /\bclass\s*=\s*(["'])([^"']+)\1/gi;
  let classMatch;
  while ((classMatch = classPattern.exec(source))) {
    const value = classMatch[2];
    const valueStart = classMatch.index + classMatch[0].lastIndexOf(value);
    const tokens = value.matchAll(/\S+/g);
    for (const tokenMatch of tokens) {
      const className = tokenMatch[0];
      if (className.startsWith('icon-ic_fluent_')
          && !data.variantByClass.has(className)) {
        markers.push(markerForValue(
          monaco,
          model,
          valueStart + tokenMatch.index,
          className,
          `Unknown Fluent icon font class "${className}".`,
        ));
      }
    }
  }
  return markers;
}
