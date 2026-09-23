// Vendors Turndown's browser ESM build into vendor/turndown/.
//
// Why a vendor step at all: src/ is loaded two ways — bundled into
// dcspad.app.js / dcspad.workbench.js by esbuild, and *unbundled* as plain ES
// modules by standalone index.html and every test suite. The unbundled paths
// have no import map (and a modern SP page freezes import-map registration
// anyway), so a bare 'turndown' specifier would 404 there. A vendored file
// behind a relative import resolves in all three.
//
// The copy is byte-for-byte from node_modules with a provenance header
// prepended: lib/turndown.browser.es.js has zero imports and a single default
// export, and its Node DOM shim (@mixmark-io/domino) is excluded by the
// package's own "browser" field, so nothing else comes along. Never hand-edit
// the output — SharePoint-specific rules live in src/html-markdown.js.

import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const toolsDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(toolsDir, '..');
const packageRoot = path.join(toolsDir, 'node_modules', 'turndown');
const outputDir = path.join(repoRoot, 'vendor', 'turndown');

// The browser ESM build specifically: the default entry pulls in domino.
const SOURCE = 'lib/turndown.browser.es.js';

// ---- the bare-import guard ------------------------------------------------

// Keywords after which a '/' opens a regex rather than dividing.
const REGEX_AFTER_WORD = new Set([
  'return', 'typeof', 'instanceof', 'in', 'of', 'new', 'delete', 'void',
  'throw', 'case', 'do', 'else', 'yield', 'await',
]);
// Punctuators after which a '/' opens a regex ('' = start of input).
const REGEX_AFTER_PUNCT = new Set(['', ...'(,=:[!&|?{};+-*%<>~^']);

// The source with its comments removed and everything else kept: strings,
// template literals and regex literals are walked over, not stripped, so a
// '//' in a URL string or a '`' in a regex cannot desync the scan. Comments
// become whitespace, as the language treats them, so tokens a comment
// separated stay separated. No dependency on purpose — this only has to
// find comments. Regex-vs-division is the usual previous-token heuristic; a
// wrong call leaves a literal unterminated, which throws rather than guesses.
export function stripComments(src) {
  let out = '';
  let i = 0;
  let prev = '';          // last significant character outside a literal
  let prevWord = '';      // the identifier/keyword that character ended
  let braces = 0;
  const templates = [];   // brace depth at each open ${ … }
  const fail = (what) => { throw new Error(`unterminated ${what} at offset ${i}`); };
  const isWord = (c) => /[\w$]/.test(c);
  const value = () => { prev = ')'; prevWord = ''; };   // a literal just ended

  // From just after an opening ` (or the } closing a ${…}): copies template
  // text; true at the closing `, false at a ${.
  const templateText = () => {
    while (i < src.length) {
      const c = src[i];
      if (c === '\\') { out += src.slice(i, i + 2); i += 2; continue; }
      if (c === '`') { out += c; i += 1; return true; }
      if (c === '$' && src[i + 1] === '{') { out += '${'; i += 2; return false; }
      out += c; i += 1;
    }
    return fail('template literal');
  };
  const enterTemplate = () => {
    if (templateText()) value();
    else { templates.push(braces); prev = '{'; prevWord = ''; }
  };

  while (i < src.length) {
    const c = src[i];
    const d = src[i + 1];
    if (c === '/' && d === '/') {
      while (i < src.length && src[i] !== '\n' && src[i] !== '\r') i += 1;
      continue;
    }
    if (c === '/' && d === '*') {
      const end = src.indexOf('*/', i + 2);
      if (end < 0) fail('block comment');
      out += /[\n\r\u2028\u2029]/.test(src.slice(i, end)) ? '\n' : ' ';
      i = end + 2;
      continue;
    }
    if (c === '"' || c === "'") {
      let j = i + 1;
      while (j < src.length && src[j] !== c) {
        if (src[j] === '\n') fail('string');
        j += src[j] === '\\' ? 2 : 1;
      }
      if (j >= src.length) fail('string');
      out += src.slice(i, j + 1);
      i = j + 1;
      value();
      continue;
    }
    if (c === '`') {
      out += c;
      i += 1;
      enterTemplate();
      continue;
    }
    if (c === '/' && (isWord(prev) ? REGEX_AFTER_WORD.has(prevWord) : REGEX_AFTER_PUNCT.has(prev))) {
      let j = i + 1;
      let inClass = false;
      for (; j < src.length; j += 1) {
        const r = src[j];
        if (r === '\n') fail('regex literal');
        if (r === '\\') { j += 1; continue; }
        if (r === '[') inClass = true;
        else if (r === ']') inClass = false;
        else if (r === '/' && !inClass) break;
      }
      if (j >= src.length) fail('regex literal');
      j += 1;
      while (j < src.length && isWord(src[j])) j += 1;   // flags
      out += src.slice(i, j);
      i = j;
      value();
      continue;
    }
    if (c === '{') braces += 1;
    if (c === '}') {
      if (templates.length && templates[templates.length - 1] === braces) {
        templates.pop();
        out += c;
        i += 1;
        enterTemplate();
        continue;
      }
      braces -= 1;
    }
    out += c;
    i += 1;
    if (/\s/.test(c)) continue;
    // A word continues only across adjacent characters, never whitespace.
    prevWord = isWord(c) ? (isWord(src[i - 2] ?? '') ? prevWord : '') + c : '';
    prev = c;
  }
  if (templates.length) fail('template expression');
  return out;
}

// Every form that names a module: `import … from 'x'`, `export … from 'x'`,
// a side-effect `import 'x'` and a dynamic `import('x')`. Matched on the
// comment-stripped text, so `import/*c*/('x')` and `from/*c*/'x'` read as
// the plain forms and a commented-out import does not count. Anything not
// relative counts — an absolute URL would be a CDN dependency, which the pad
// has none of either.
// Accepted false failure: a string literal that merely reads like one
// ("import('pkg')") still trips it. That fails the build loudly and a human
// looks, which is the right direction for this guard to be wrong in.
const BARE_IMPORT = /(?:\bfrom\s*|\bimport\s*\(?\s*)(['"`])(?![./])([^'"`]*)\1/;

// The bare specifier the source names, or null.
export function findBareImport(source) {
  const found = stripComments(source).match(BARE_IMPORT);
  return found ? found[2] : null;
}

// ---- the vendor step -------------------------------------------------------

async function readPackageJson() {
  try {
    return JSON.parse(await readFile(path.join(packageRoot, 'package.json'), 'utf8'));
  } catch {
    throw new Error(
      'turndown is not installed. Run `npm install` in tools/ first.\n'
      + 'On Windows ARM64 use the system Node — see the gotchas in CLAUDE.md.',
    );
  }
}

async function main() {
  const pkg = await readPackageJson();
  const source = await readFile(path.join(packageRoot, SOURCE), 'utf8');

  // A bare specifier surviving into the vendored file would break the
  // unbundled paths silently — the whole reason this step exists.
  const bare = findBareImport(source);
  if (bare !== null) {
    throw new Error(`${SOURCE} carries a bare import ('${bare}'); `
      + 'it can no longer be vendored as one file');
  }
  const code = stripComments(source);
  if (!/export\s*\{[^}]*\bas default\b/.test(code) && !/export\s+default\b/.test(code)) {
    throw new Error(`${SOURCE} has no default export`);
  }

  const header = `// GENERATED by tools/build-vendor-turndown.mjs — do not edit.\n`
    + `// turndown ${pkg.version} (${pkg.license}) — ${SOURCE}\n`
    + `// SharePoint-specific rules belong in src/html-markdown.js, not here.\n`;
  const body = `${header}${source}`;

  await mkdir(outputDir, { recursive: true });
  await writeFile(path.join(outputDir, 'turndown.js'), body, 'utf8');
  await writeFile(path.join(outputDir, 'LICENSE'),
    await readFile(path.join(packageRoot, 'LICENSE'), 'utf8'), 'utf8');
  await writeFile(path.join(outputDir, 'version.json'), `${JSON.stringify({
    name: 'turndown',
    version: pkg.version,
    license: pkg.license,
    source: SOURCE,
    sha256: createHash('sha256').update(source).digest('hex'),
  }, null, 2)}\n`, 'utf8');

  const kb = (body.length / 1024).toFixed(1);
  console.log(`vendor/turndown/turndown.js  turndown ${pkg.version}  ${kb} KB`);
}

// Runs only as a script, so the guard can be imported and exercised alone.
if (path.resolve(process.argv[1] || '').toLowerCase()
  === fileURLToPath(import.meta.url).toLowerCase()) {
  main().catch((error) => {
    console.error(error.message);
    process.exitCode = 1;
  });
}
