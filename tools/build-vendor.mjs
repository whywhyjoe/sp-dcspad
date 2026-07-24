// One-time (re)build of vendor/codemirror.mjs — a single ESM bundle of
// CodeMirror 6 so the app has no CDN dependency and cannot hit CM6's
// duplicate-@codemirror/state pitfall.
//
// Regenerate with:
//   cd devpad/tools && npm init -y && npm i esbuild codemirror @codemirror/lang-html @codemirror/lang-css @codemirror/lang-javascript @codemirror/theme-one-dark
//   node build-vendor.mjs

import { build } from 'esbuild';
import { writeFileSync } from 'fs';

writeFileSync('_entry.js', `
export { EditorView, keymap, lineNumbers } from '@codemirror/view';
export { EditorState, Compartment } from '@codemirror/state';
export { basicSetup } from 'codemirror';
export { html } from '@codemirror/lang-html';
export { css } from '@codemirror/lang-css';
export { javascript } from '@codemirror/lang-javascript';
export { oneDark } from '@codemirror/theme-one-dark';
export { indentWithTab } from '@codemirror/commands';
`);

await build({
  entryPoints: ['_entry.js'],
  bundle: true,
  format: 'esm',
  minify: true,
  outfile: '../vendor/codemirror.mjs',
  logLevel: 'info',
});
console.log('vendor/codemirror.mjs rebuilt');
