// Rebuilds the complete Monaco runtime used by DCSPad:
//   vendor/monaco/monaco.js          self-contained ESM editor bundle
//   vendor/monaco/monaco.css         editor/codicon styles
//   vendor/monaco/*.worker.js        same-origin classic workers
//   vendor/monaco/pnpjs-types.json   PnPjs 2.15.0 declaration graph
//   vendor/monaco/version.json       one cache-busting stamp for the set
//
// All browser artifacts deliberately use .js rather than .mjs because
// SharePoint serves .mjs as application/octet-stream on the target tenant.

import { build } from 'esbuild';
import ts from 'typescript';
import { createHash } from 'node:crypto';
import {
  existsSync, mkdtempSync, mkdirSync, readFileSync, readdirSync, rmSync,
  statSync, writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const toolsDir = dirname(fileURLToPath(import.meta.url));
const repoDir = resolve(toolsDir, '..');
const outDir = join(repoDir, 'vendor', 'monaco');
const nodeModules = join(toolsDir, 'node_modules');

const MONACO_VERSION = '0.55.1';
const PNPJS_VERSION = '2.15.0';

const mainEntry = `
import 'monaco-editor/esm/vs/language/css/monaco.contribution.js';
import 'monaco-editor/esm/vs/language/html/monaco.contribution.js';
import 'monaco-editor/esm/vs/language/typescript/monaco.contribution.js';
import 'monaco-editor/esm/vs/basic-languages/css/css.contribution.js';
import 'monaco-editor/esm/vs/basic-languages/html/html.contribution.js';
import 'monaco-editor/esm/vs/basic-languages/javascript/javascript.contribution.js';
import 'monaco-editor/esm/vs/basic-languages/typescript/typescript.contribution.js';
export * from 'monaco-editor/esm/vs/editor/editor.api.js';
export * as css from 'monaco-editor/esm/vs/language/css/monaco.contribution.js';
export * as html from 'monaco-editor/esm/vs/language/html/monaco.contribution.js';
export * as typescript from 'monaco-editor/esm/vs/language/typescript/monaco.contribution.js';
`;

const workerEntries = {
  'editor.worker': join(nodeModules, 'monaco-editor', 'esm', 'vs', 'editor', 'editor.worker.js'),
  'css.worker': join(nodeModules, 'monaco-editor', 'esm', 'vs', 'language', 'css', 'css.worker.js'),
  'html.worker': join(nodeModules, 'monaco-editor', 'esm', 'vs', 'language', 'html', 'html.worker.js'),
  'ts.worker': join(nodeModules, 'monaco-editor', 'esm', 'vs', 'language', 'typescript', 'ts.worker.js'),
};

const globalTypes = `
import type * as PnPjs from "@pnp/pnpjs";

declare global {
  const pnp: typeof PnPjs;
  interface Window {
    pnp: typeof PnPjs;
    _spPageContextInfo?: {
      webAbsoluteUrl?: string;
      webServerRelativeUrl?: string;
      siteAbsoluteUrl?: string;
      userDisplayName?: string;
      userEmail?: string;
      formDigestValue?: string;
      [key: string]: unknown;
    };
  }
  const _spPageContextInfo: NonNullable<Window["_spPageContextInfo"]>;
}

export {};
`.trimStart();

function walkFiles(root, predicate, out = []) {
  if (!existsSync(root)) return out;
  for (const name of readdirSync(root)) {
    const path = join(root, name);
    const stat = statSync(path);
    if (stat.isDirectory()) walkFiles(path, predicate, out);
    else if (predicate(path)) out.push(path);
  }
  return out;
}

function virtualPath(path) {
  return `file:///node_modules/${relative(nodeModules, path).split(sep).join('/')}`;
}

function buildPnpTypePayload() {
  const packageRoots = [
    join(nodeModules, '@pnp', 'common'),
    join(nodeModules, '@pnp', 'config-store'),
    join(nodeModules, '@pnp', 'graph'),
    join(nodeModules, '@pnp', 'logging'),
    join(nodeModules, '@pnp', 'odata'),
    join(nodeModules, '@pnp', 'pnpjs'),
    join(nodeModules, '@pnp', 'sp'),
    join(nodeModules, '@pnp', 'sp-addinhelpers'),
  ];
  const files = packageRoots.flatMap((root) =>
    walkFiles(root, (path) => path.endsWith('.d.ts')));
  const tslib = join(nodeModules, 'tslib', 'tslib.d.ts');
  if (existsSync(tslib)) files.push(tslib);

  const libs = files
    .sort()
    .map((path) => ({ filePath: virtualPath(path), content: readFileSync(path, 'utf8') }));
  libs.push({ filePath: 'file:///dcspad/globals.d.ts', content: globalTypes });
  return { version: PNPJS_VERSION, libs };
}

function validatePnpGlobal() {
  const dir = mkdtempSync(join(tmpdir(), 'dcspad-pnp-types-'));
  try {
    const globalsPath = join(dir, 'globals.d.ts');
    const probePath = join(dir, 'probe.ts');
    writeFileSync(globalsPath, globalTypes);
    writeFileSync(probePath, `
const webTitle: PromiseLike<unknown> = pnp.sp.web.select("Title").get();
const listItems = pnp.sp.web.lists.getByTitle("Documents").items.select("Id", "Title");
const { sp } = pnp;
const explicitWeb = pnp.SPNS.Web("https://contoso.sharepoint.com/sites/dev");
void webTitle; void listItems; void sp; void explicitWeb;
`);
    const program = ts.createProgram([globalsPath, probePath], {
      target: ts.ScriptTarget.ES2020,
      module: ts.ModuleKind.ESNext,
      moduleResolution: ts.ModuleResolutionKind.Node10,
      skipLibCheck: true,
      strict: true,
      noEmit: true,
      baseUrl: toolsDir,
      paths: { '@pnp/*': ['node_modules/@pnp/*'] },
      types: [],
    });
    const diagnostics = ts.getPreEmitDiagnostics(program);
    if (diagnostics.length) {
      const message = ts.formatDiagnosticsWithColorAndContext(diagnostics, {
        getCanonicalFileName: (f) => f,
        getCurrentDirectory: () => toolsDir,
        getNewLine: () => '\n',
      });
      throw new Error(`PnPjs global declaration validation failed:\n${message}`);
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });

await build({
  stdin: {
    contents: mainEntry,
    resolveDir: toolsDir,
    sourcefile: 'monaco-entry.js',
  },
  bundle: true,
  format: 'esm',
  minify: true,
  sourcemap: false,
  outfile: join(outDir, 'monaco.js'),
  loader: { '.ttf': 'file' },
  assetNames: 'assets/[name]-[hash]',
  logLevel: 'info',
});

await build({
  entryPoints: workerEntries,
  bundle: true,
  format: 'iife',
  minify: true,
  sourcemap: false,
  outdir: outDir,
  entryNames: '[name]',
  logLevel: 'info',
});

validatePnpGlobal();
const pnpTypes = buildPnpTypePayload();
writeFileSync(join(outDir, 'pnpjs-types.json'), JSON.stringify(pnpTypes));
const runtimeHash = createHash('sha256');
for (const path of walkFiles(outDir, (candidate) => !candidate.endsWith('version.json')).sort()) {
  runtimeHash.update(relative(outDir, path).split(sep).join('/'));
  runtimeHash.update(readFileSync(path));
}
writeFileSync(join(outDir, 'version.json'), JSON.stringify({
  monaco: MONACO_VERSION,
  pnpjs: PNPJS_VERSION,
  runtimeHash: runtimeHash.digest('hex').slice(0, 16),
}, null, 2));

console.log(`Monaco ${MONACO_VERSION} + PnPjs ${PNPJS_VERSION} runtime rebuilt`);
console.log(`${pnpTypes.libs.length} declaration files bundled`);
