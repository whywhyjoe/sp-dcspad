// JSON language wiring.
//
// Highlighting, validation, schema support and formatting all come from
// Monaco's own JSON contribution — VS Code's "JSON Language Features"
// (vscode-json-languageservice) running in vendor/monaco/json.worker.js.
// tools/build-monaco.mjs bundles it; this module only supplies the two things
// the vendored runtime cannot know about:
//
//   * theme rules, because dcspad-dark is ours (without them a property key
//     is coloured by the generic `string` rule and reads the same as its
//     value, which is not how VS Code renders JSON);
//   * language adoption for models the pad creates with a .json URI, since
//     the language service only lights up models whose language id is json.

export const JSON_LANGUAGE_ID = 'json';

// Token names come from Monaco's JSON tokenizer (vs/language/json/
// tokenization.js) and are stable across its releases. Hues are VS Code Dark
// Modern on the #17191f ground, matching the HTML/CSS/JS palette in
// editors.js. Comment tokens (comment.line.json / comment.block.json) are
// deliberately absent: the generic `comment` rule already resolves them.
export const JSON_THEME_RULES = [
  { token: 'string.key.json', foreground: '9CDCFE' },
  { token: 'string.value.json', foreground: 'CE9178' },
  { token: 'number.json', foreground: 'B5CEA8' },
  { token: 'keyword.json', foreground: '569CD6' },
  { token: 'delimiter.bracket.json', foreground: 'D4D4D4' },
  { token: 'delimiter.array.json', foreground: 'D4D4D4' },
  { token: 'delimiter.colon.json', foreground: 'D4D4D4' },
  { token: 'delimiter.comma.json', foreground: 'D4D4D4' },
];

function looksLikeJsonUri(model) {
  return /\.jsonc?$/i.test(model.uri?.path || '');
}

/**
 * Configure JSON diagnostics and adopt JSON models as they appear.
 *
 * @param {*} monaco
 * @param {{allowComments?: boolean, trailingCommas?: 'error'|'warning'|'ignore',
 *          enableSchemaRequest?: boolean, adoptByUri?: boolean}} [options]
 * @returns {{dispose: () => void}}
 */
export function installJsonLanguage(monaco, options = {}) {
  const allowComments = options.allowComments === true;

  monaco.json.jsonDefaults.setDiagnosticsOptions({
    validate: true,
    allowComments,
    trailingCommas: options.trailingCommas || (allowComments ? 'ignore' : 'error'),
    schemaValidation: 'error',
    // Off by default: fetching a $schema URL would be an outbound request from
    // a worker on a SharePoint page, which either fails quietly or surprises
    // the tenant. Callers that want SchemaStore can opt in.
    enableSchemaRequest: options.enableSchemaRequest === true,
  });

  if (options.adoptByUri === false) return { dispose: () => {} };

  const adopt = (model) => {
    if (!model.isDisposed() && looksLikeJsonUri(model) && model.getLanguageId() !== JSON_LANGUAGE_ID) {
      monaco.editor.setModelLanguage(model, JSON_LANGUAGE_ID);
    }
  };
  for (const model of monaco.editor.getModels()) adopt(model);
  const subscription = monaco.editor.onDidCreateModel(adopt);

  return { dispose: () => subscription.dispose() };
}
