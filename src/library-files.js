// File signatures and structural guards for the two user-managed libraries.
// Both documents historically used { v, items }, so item-level validation is
// retained for unsigned legacy exports while new exports identify themselves.

export const FRAMEWORK_CATALOG_KIND = 'dcspad-framework-catalog';
export const SNIPPET_LIBRARY_KIND = 'dcspad-snippet-library';

const KIND_LABELS = {
  [FRAMEWORK_CATALOG_KIND]: 'framework catalog',
  [SNIPPET_LIBRARY_KIND]: 'snippet library',
};

const isObject = (value) =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const isNonEmptyString = (value) =>
  typeof value === 'string' && value.trim().length > 0;

function validateRoot(doc, expectedKind, expectedLabel, allowUnsignedEmpty) {
  if (!isObject(doc) || !Array.isArray(doc.items)) {
    return { ok: false, message: `Not a DCSPad ${expectedLabel} file.` };
  }

  if (doc.kind && doc.kind !== expectedKind) {
    const actual = KIND_LABELS[doc.kind];
    return {
      ok: false,
      message: actual
        ? `This is a DCSPad ${actual}, not a ${expectedLabel}.`
        : `Unsupported DCSPad file type "${doc.kind}".`,
    };
  }

  if (!doc.kind && doc.items.length === 0 && !allowUnsignedEmpty) {
    return {
      ok: false,
      message: `This empty legacy file has no "kind" signature, so it cannot be safely identified as a ${expectedLabel}.`,
    };
  }

  return { ok: true };
}

function duplicateId(items) {
  const ids = new Set();
  for (const item of items) {
    if (!item.id) continue;
    if (ids.has(item.id)) return item.id;
    ids.add(item.id);
  }
  return '';
}

export function validateFrameworkCatalog(doc, { allowUnsignedEmpty = false } = {}) {
  const root = validateRoot(
    doc,
    FRAMEWORK_CATALOG_KIND,
    'framework catalog',
    allowUnsignedEmpty,
  );
  if (!root.ok) return root;

  for (let index = 0; index < doc.items.length; index += 1) {
    const item = doc.items[index];
    const validSource = isNonEmptyString(item?.js)
      || isNonEmptyString(item?.css)
      || item?.needsConfig === true;
    if (!isObject(item)
        || !isNonEmptyString(item.id)
        || !isNonEmptyString(item.name)
        || !validSource
        || (item.js !== undefined && typeof item.js !== 'string')
        || (item.css !== undefined && typeof item.css !== 'string')
        || (item.order !== undefined
          && (!Number.isInteger(item.order) || item.order < 1))) {
      return {
        ok: false,
        message: `Item ${index + 1} is not a valid framework entry (expected id, name, a JS/CSS source, and an optional positive integer order).`,
      };
    }
  }

  const repeated = duplicateId(doc.items);
  if (repeated) {
    return { ok: false, message: `Framework id "${repeated}" appears more than once.` };
  }

  return {
    ok: true,
    doc: {
      kind: FRAMEWORK_CATALOG_KIND,
      v: Number.isInteger(doc.v) && doc.v > 0 ? doc.v : 1,
      items: doc.items.map((item) => ({ ...item })),
    },
  };
}

export function validateSnippetLibrary(doc, { allowUnsignedEmpty = false } = {}) {
  const root = validateRoot(
    doc,
    SNIPPET_LIBRARY_KIND,
    'snippet library',
    allowUnsignedEmpty,
  );
  if (!root.ok) return root;

  for (let index = 0; index < doc.items.length; index += 1) {
    const item = doc.items[index];
    if (!isObject(item)
        || !isNonEmptyString(item.name)
        || typeof item.code !== 'string'
        || !['html', 'css', 'js'].includes(item.lang)
        || (item.id !== undefined && !isNonEmptyString(item.id))) {
      return {
        ok: false,
        message: `Item ${index + 1} is not a valid snippet entry (expected name, html/css/js language, and code).`,
      };
    }
  }

  const repeated = duplicateId(doc.items);
  if (repeated) {
    return { ok: false, message: `Snippet id "${repeated}" appears more than once.` };
  }

  return {
    ok: true,
    doc: {
      kind: SNIPPET_LIBRARY_KIND,
      v: 1,
      items: doc.items.map((item) => ({ ...item })),
    },
  };
}
