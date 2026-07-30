// Generates the compact Monaco data consumed by DCSPad from the local BMO
// SharePoint Design System sources. The design system remains buildless for
// consumers; this is development tooling for the workbench itself.

import { createHash } from 'node:crypto';
import { mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const toolsDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(toolsDir, '..');
const configPath = path.join(repoRoot, 'dcspad.config.json');
const outputDir = path.join(repoRoot, 'vendor', 'intelligence');

function optionValue(name) {
  const exactIndex = process.argv.indexOf(name);
  if (exactIndex !== -1) {
    const value = process.argv[exactIndex + 1];
    if (!value || value.startsWith('--')) {
      throw new Error(`${name} requires a directory path`);
    }
    return value;
  }
  const prefix = `${name}=`;
  const inline = process.argv.find((arg) => arg.startsWith(prefix));
  return inline ? inline.slice(prefix.length) : '';
}

function resolveBuildRoot(option, environmentName, siblingRepo) {
  const configured = optionValue(option)
    || process.env[environmentName]
    || path.resolve(repoRoot, '..', siblingRepo);
  return path.isAbsolute(configured)
    ? path.normalize(configured)
    : path.resolve(repoRoot, configured);
}

async function requireDirectory(root, label, option, environmentName) {
  try {
    if ((await stat(root)).isDirectory()) return;
  } catch {
    // Use the actionable error below for missing and inaccessible paths.
  }
  throw new Error(
    `${label} source directory was not found: ${root}\n`
    + `Pass ${option} <path> or set ${environmentName}.`,
  );
}

// Runtime asset URLs belong in dcspad.config.json. Intelligence is generated
// from local source repositories, which are deliberately configured
// separately so a SharePoint URL fragment is never mistaken for a disk path.
const sourceRoot = resolveBuildRoot(
  '--design-root',
  'DCSPAD_DESIGN_SYSTEM_ROOT',
  'bsp-design-system',
);
const fluentIconsRoot = resolveBuildRoot(
  '--fluent-icons-root',
  'DCSPAD_FLUENT_ICONS_ROOT',
  'bsp-fluent-icon-lib',
);
await requireDirectory(
  sourceRoot,
  'BSP design system',
  '--design-root',
  'DCSPAD_DESIGN_SYSTEM_ROOT',
);
await requireDirectory(
  fluentIconsRoot,
  'Fluent icon library',
  '--fluent-icons-root',
  'DCSPAD_FLUENT_ICONS_ROOT',
);

const config = JSON.parse(await readFile(configPath, 'utf8'));
const designConfig = config?.assets?.designSystem;
const fluentIconsConfig = config?.assets?.fluentIcons;
if (!designConfig) {
  throw new Error('dcspad.config.json must define assets.designSystem');
}

const relativeSource = (key, fallback) =>
  designConfig.files?.[key] || fallback;
const sourceSpecs = [
  { key: 'tokens', file: relativeSource('tokens', 'colors_and_type.css'), tokens: true, classes: true },
  { key: 'components', file: relativeSource('components', 'components.css'), classes: true },
  { key: 'editorial', file: relativeSource('editorial', 'editorial.css'), tokens: true, classes: true, scope: 'editorial' },
];

if (!fluentIconsConfig) {
  throw new Error('dcspad.config.json must define assets.fluentIcons');
}
const fluentCatalogFile = fluentIconsConfig.files?.catalog || 'fluent-font-library.json';

const cleanComment = (raw) => {
  const text = String(raw || '')
    .split(/\r?\n/)
    .map((line) => line.replace(/^\s*\*?\s?/, '').trim())
    .filter((line) => line && !/^[-=]{4,}$/.test(line))
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
  return text.length > 420 ? `${text.slice(0, 417).trimEnd()}…` : text;
};

const lineNumberAt = (source, index) =>
  source.slice(0, index).split('\n').length;

function contextualComments(source) {
  const comments = [];
  const pattern = /\/\*([\s\S]*?)\*\//g;
  let match;
  while ((match = pattern.exec(source))) {
    const lineStart = source.lastIndexOf('\n', match.index) + 1;
    const isStandalone = source.slice(lineStart, match.index).trim() === '';
    comments.push({
      start: match.index,
      end: pattern.lastIndex,
      text: cleanComment(match[1]),
      isStandalone,
    });
  }
  return comments;
}

function nearestContext(comments, index) {
  for (let i = comments.length - 1; i >= 0; i--) {
    if (comments[i].end <= index && comments[i].isStandalone && comments[i].text) {
      return comments[i].text;
    }
  }
  return '';
}

function tokenCategory(name) {
  if (name.startsWith('--bmo-chart-')) return 'BMO chart color';
  if (name.startsWith('--bmo-')) return 'BMO brand token';
  if (name.startsWith('--f2-')) return 'Fluent 2 token';
  if (name.startsWith('--surface-')) return 'Surface';
  if (name.startsWith('--fg-')) return 'Foreground';
  if (name.startsWith('--stroke-')) return 'Stroke';
  if (name.startsWith('--accent-')) return 'Interactive accent';
  if (name.startsWith('--space-')) return 'Spacing';
  if (name.startsWith('--radius-')) return 'Radius';
  if (name.startsWith('--shadow-')) return 'Elevation';
  if (name.startsWith('--type-') || name.startsWith('--font-') || name.startsWith('--weight-')) return 'Typography';
  if (name.startsWith('--motion-') || name.startsWith('--ease-') || name.startsWith('--reveal-') || name.startsWith('--lift-')) return 'Motion';
  if (name.startsWith('--ed-')) return 'Editorial mode';
  return 'Design token';
}

function extractTokens(source, file, scope) {
  const comments = contextualComments(source);
  const tokens = [];
  const declaration = /(^|\n)\s*(--[\w-]+)\s*:\s*([^;]+);(?:[ \t]*\/\*([^*]*?)\*\/)?/g;
  let match;
  while ((match = declaration.exec(source))) {
    const index = match.index + match[1].length;
    const name = match[2];
    const value = match[3].replace(/\s+/g, ' ').trim();
    const inline = cleanComment(match[4]);
    const context = nearestContext(comments, index);
    tokens.push({
      name,
      value,
      category: tokenCategory(name),
      description: inline || context || tokenCategory(name),
      ...(scope ? { scope } : {}),
      source: { file, line: lineNumberAt(source, index) },
    });
  }
  return tokens;
}

function classKind(name, file) {
  if (name.startsWith('is-')) return 'state';
  if (name.includes('__')) return 'element';
  if (name.includes('--')) return 'modifier';
  if (file === 'colors_and_type.css') return 'utility';
  return 'base';
}

function inferredBase(name) {
  if (name.includes('__')) return name.split('__')[0];
  if (name.includes('--')) return name.split('--')[0];
  return '';
}

function defaultClassDescription(record) {
  if (record.kind === 'modifier') return `Modifier for .${record.base}.`;
  if (record.kind === 'element') return `Element of .${record.base}.`;
  if (record.kind === 'state') {
    const targets = record.appliesTo.slice(0, 5).map((name) => `.${name}`).join(', ');
    return targets ? `State class used with ${targets}.` : 'Design-system state class.';
  }
  if (record.kind === 'utility') return 'Design-system utility class.';
  return 'Design-system component or composition class.';
}

function extractClassOccurrences(source, file, scope) {
  const comments = contextualComments(source);
  const occurrences = [];
  const ruleStart = /([^{}]+)\{/g;
  let match;
  while ((match = ruleStart.exec(source))) {
    const selectorStart = match.index;
    const selector = match[1]
      .replace(/\/\*[\s\S]*?\*\//g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (!selector || selector.startsWith('@')) continue;

    const names = [...selector.matchAll(/\.(-?[_a-zA-Z][\w-]*)/g)]
      .map((item) => item[1]);
    const uniqueNames = [...new Set(names)];
    if (!uniqueNames.length) continue;

    const context = nearestContext(comments, selectorStart);
    for (const name of uniqueNames) {
      occurrences.push({
        name,
        selector: selector.length > 220 ? `${selector.slice(0, 217)}…` : selector,
        context,
        peers: uniqueNames.filter((peer) => peer !== name),
        ...(scope ? { scope } : {}),
        source: { file, line: lineNumberAt(source, selectorStart) },
      });
    }
  }
  return occurrences;
}

function mergeClasses(occurrences) {
  const byName = new Map();
  for (const occurrence of occurrences) {
    let record = byName.get(occurrence.name);
    if (!record) {
      const kind = classKind(occurrence.name, occurrence.source.file);
      record = {
        name: occurrence.name,
        kind,
        ...(inferredBase(occurrence.name) ? { base: inferredBase(occurrence.name) } : {}),
        description: occurrence.context,
        appliesTo: [],
        selectors: [],
        scopes: [],
        source: occurrence.source,
      };
      byName.set(occurrence.name, record);
    }
    if (!record.description && occurrence.context) record.description = occurrence.context;
    for (const peer of occurrence.peers) {
      if (!record.appliesTo.includes(peer)) record.appliesTo.push(peer);
    }
    if (!record.selectors.includes(occurrence.selector) && record.selectors.length < 4) {
      record.selectors.push(occurrence.selector);
    }
    if (occurrence.scope && !record.scopes.includes(occurrence.scope)) {
      record.scopes.push(occurrence.scope);
    }
  }

  for (const record of byName.values()) {
    if (!record.description) record.description = defaultClassDescription(record);
    if (record.base && !record.appliesTo.includes(record.base)) {
      record.appliesTo.unshift(record.base);
    }
    record.appliesTo.sort();
    if (!record.appliesTo.length) delete record.appliesTo;
    if (!record.scopes.length) delete record.scopes;
  }
  return [...byName.values()].sort((a, b) => a.name.localeCompare(b.name));
}

const sourceHashes = {};
const allTokens = [];
const allClassOccurrences = [];

for (const spec of sourceSpecs) {
  const sourcePath = path.resolve(sourceRoot, spec.file);
  if (!sourcePath.startsWith(sourceRoot + path.sep)) {
    throw new Error(`Refusing to read design-system source outside ${sourceRoot}: ${spec.file}`);
  }
  const source = await readFile(sourcePath, 'utf8');
  sourceHashes[spec.file] = createHash('sha256').update(source).digest('hex');
  if (spec.tokens) allTokens.push(...extractTokens(source, spec.file, spec.scope));
  if (spec.classes) {
    allClassOccurrences.push(...extractClassOccurrences(source, spec.file, spec.scope));
  }
}

const tokenMap = new Map();
for (const token of allTokens) {
  const existing = tokenMap.get(token.name);
  if (!existing || (!token.scope && existing.scope)) tokenMap.set(token.name, token);
}
const tokens = [...tokenMap.values()].sort((a, b) => a.name.localeCompare(b.name));
const classes = mergeClasses(allClassOccurrences);

function parseFluentFilename(filename) {
  const match = String(filename || '').match(
    /^ic_fluent_(.+)_(\d+)_(regular|filled|light|color)(?:_(ltr|rtl))?\.svg$/i,
  );
  if (!match) return null;
  return {
    idBase: match[1].toLowerCase(),
    size: Number(match[2]),
    style: match[3].toLowerCase(),
    direction: match[4]?.toLowerCase() || '',
  };
}

function buildFluentIcons(raw) {
  if (!raw || !Array.isArray(raw.icons)) {
    throw new Error('fluent-font-library.json must contain an icons array');
  }
  const byIdBase = new Map();
  for (const sourceIcon of raw.icons) {
    const grouped = new Map();
    for (const filename of Array.isArray(sourceIcon?.filenames) ? sourceIcon.filenames : []) {
      const parsed = parseFluentFilename(filename);
      if (!parsed) continue;
      if (!grouped.has(parsed.idBase)) grouped.set(parsed.idBase, []);
      const suffix = [
        parsed.size,
        parsed.style,
        parsed.direction,
      ].filter(Boolean).join('-');
      const fontKey = filename.replace(/\.svg$/i, '');
      grouped.get(parsed.idBase).push({
        suffix,
        font: Object.prototype.hasOwnProperty.call(
          sourceIcon.fontCodepoints || {},
          fontKey,
        ),
      });
    }
    for (const [idBase, rawVariants] of grouped) {
      let record = byIdBase.get(idBase);
      if (!record) {
        record = {
          name: String(sourceIcon.name || idBase.replaceAll('_', ' ')).trim(),
          slug: idBase.replaceAll('_', '-'),
          idBase,
          ...(String(sourceIcon.description || '').trim()
            ? { description: String(sourceIcon.description).trim() }
            : {}),
          metaphors: [],
          variants: [],
          svgOnly: [],
        };
        byIdBase.set(idBase, record);
      }
      for (const metaphor of Array.isArray(sourceIcon.metaphor) ? sourceIcon.metaphor : []) {
        const clean = String(metaphor || '').trim();
        if (clean && !record.metaphors.includes(clean)) record.metaphors.push(clean);
      }
      for (const variant of rawVariants) {
        if (!record.variants.includes(variant.suffix)) {
          record.variants.push(variant.suffix);
        }
        if (!variant.font && !record.svgOnly.includes(variant.suffix)) {
          record.svgOnly.push(variant.suffix);
        }
      }
    }
  }

  const styleOrder = { regular: 0, filled: 1, light: 2 };
  for (const record of byIdBase.values()) {
    record.metaphors.sort((a, b) => a.localeCompare(b));
    if (!record.metaphors.length) delete record.metaphors;
    const sortVariants = (a, b) => {
      const [aSize, aStyle, aDirection = ''] = a.split('-');
      const [bSize, bStyle, bDirection = ''] = b.split('-');
      return Number(aSize) - Number(bSize)
        || (styleOrder[aStyle] ?? 9) - (styleOrder[bStyle] ?? 9)
        || aDirection.localeCompare(bDirection);
    };
    record.variants.sort(sortVariants);
    record.svgOnly.sort(sortVariants);
    if (!record.svgOnly.length) delete record.svgOnly;
  }
  return [...byIdBase.values()].sort((a, b) => a.slug.localeCompare(b.slug));
}

const fluentCatalogPath = path.resolve(fluentIconsRoot, fluentCatalogFile);
if (!fluentCatalogPath.startsWith(fluentIconsRoot + path.sep)) {
  throw new Error(`Refusing to read Fluent icon catalog outside ${fluentIconsRoot}`);
}
const fluentCatalogSource = await readFile(fluentCatalogPath, 'utf8');
const fluentIcons = buildFluentIcons(JSON.parse(fluentCatalogSource));
const fluentVariantCount = fluentIcons.reduce(
  (count, icon) => count + icon.variants.length,
  0,
);
const fluentFontVariantCount = fluentIcons.reduce(
  (count, icon) => count + icon.variants.length - (icon.svgOnly?.length || 0),
  0,
);

const artifact = {
  schemaVersion: 1,
  pack: 'bsp-design',
  tokens,
  classes,
};
const artifactText = `${JSON.stringify(artifact, null, 2)}\n`;
const artifactHash = createHash('sha256').update(artifactText).digest('hex');
const fluentArtifact = {
  schemaVersion: 1,
  pack: 'fluent-icons',
  icons: fluentIcons,
};
const fluentArtifactText = `${JSON.stringify(fluentArtifact, null, 2)}\n`;
const fluentArtifactHash = createHash('sha256').update(fluentArtifactText).digest('hex');
const manifest = {
  schemaVersion: 1,
  pack: 'bsp-design',
  artifact: 'bsp-design.json',
  artifactSha256: artifactHash,
  counts: { tokens: tokens.length, classes: classes.length },
  sources: sourceHashes,
  fluentIcons: {
    artifact: 'fluent-icons.json',
    artifactSha256: fluentArtifactHash,
    counts: {
      icons: fluentIcons.length,
      variants: fluentVariantCount,
      fontVariants: fluentFontVariantCount,
    },
    source: {
      file: fluentCatalogFile,
      sha256: createHash('sha256').update(fluentCatalogSource).digest('hex'),
    },
  },
};

await mkdir(outputDir, { recursive: true });
await writeFile(path.join(outputDir, 'bsp-design.json'), artifactText);
await writeFile(path.join(outputDir, 'fluent-icons.json'), fluentArtifactText);
await writeFile(
  path.join(outputDir, 'manifest.json'),
  `${JSON.stringify(manifest, null, 2)}\n`,
);

console.log(
  `BSP intelligence generated: ${tokens.length} tokens, ${classes.length} classes; `
  + `${fluentIcons.length} Fluent icons, ${fluentVariantCount} variants`,
);
console.log(`Sources: ${sourceRoot}; ${fluentIconsRoot}`);
