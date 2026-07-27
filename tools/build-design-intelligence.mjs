// Generates the compact Monaco data consumed by DCSPad from the local BMO
// SharePoint Design System sources. The design system remains buildless for
// consumers; this is development tooling for the workbench itself.

import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const toolsDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(toolsDir, '..');
const configPath = path.join(repoRoot, 'dcspad.config.json');
const outputDir = path.join(repoRoot, 'vendor', 'intelligence');

const config = JSON.parse(await readFile(configPath, 'utf8'));
const designConfig = config?.assets?.designSystem;
if (!designConfig || typeof designConfig.localBaseUrl !== 'string') {
  throw new Error('dcspad.config.json must define assets.designSystem.localBaseUrl');
}
if (/^[a-z][a-z\d+.-]*:/i.test(designConfig.localBaseUrl)) {
  throw new Error('assets.designSystem.localBaseUrl must be a local path when generating intelligence');
}

const sourceRoot = path.resolve(repoRoot, designConfig.localBaseUrl);
const relativeSource = (key, fallback) =>
  designConfig.files?.[key] || fallback;
const sourceSpecs = [
  { key: 'tokens', file: relativeSource('tokens', 'colors_and_type.css'), tokens: true, classes: true },
  { key: 'components', file: relativeSource('components', 'components.css'), classes: true },
  { key: 'editorial', file: relativeSource('editorial', 'editorial.css'), tokens: true, classes: true, scope: 'editorial' },
];

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

const artifact = {
  schemaVersion: 1,
  pack: 'bsp-design',
  tokens,
  classes,
};
const artifactText = `${JSON.stringify(artifact, null, 2)}\n`;
const artifactHash = createHash('sha256').update(artifactText).digest('hex');
const manifest = {
  schemaVersion: 1,
  pack: 'bsp-design',
  artifact: 'bsp-design.json',
  artifactSha256: artifactHash,
  counts: { tokens: tokens.length, classes: classes.length },
  sources: sourceHashes,
};

await mkdir(outputDir, { recursive: true });
await writeFile(path.join(outputDir, 'bsp-design.json'), artifactText);
await writeFile(
  path.join(outputDir, 'manifest.json'),
  `${JSON.stringify(manifest, null, 2)}\n`,
);

console.log(`BSP intelligence generated: ${tokens.length} tokens, ${classes.length} classes`);
