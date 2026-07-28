// Convert a parsed JSON value into the harness's serialized-node format so
// the inspector/tree renderer (tree-view.js + sp-shapes.js) can be reused on
// data that never crossed the preview-frame boundary: network response
// bodies, and the workbench's direct /_api results.

export function toNode(v, depth = 0, { maxDepth = 6, maxItems = 100 } = {}) {
  if (v === null) return { t: 'null' };
  switch (typeof v) {
    case 'string': return { t: 'str', v };
    case 'number': return { t: 'num', v };
    case 'boolean': return { t: 'bool', v };
    case 'undefined': return { t: 'undef' };
  }
  if (depth >= maxDepth) return { t: 'maxdepth', v: Array.isArray(v) ? `Array(${v.length})` : '{…}' };
  const opts = { maxDepth, maxItems };
  if (Array.isArray(v)) {
    return { t: 'arr', n: v.length, items: v.slice(0, maxItems).map((x) => toNode(x, depth + 1, opts)), trunc: v.length > maxItems };
  }
  const keys = Object.keys(v);
  return {
    t: 'obj', cls: 'Object',
    keys: keys.slice(0, maxItems).map((k) => [k, toNode(v[k], depth + 1, opts)]),
    trunc: keys.length > maxItems,
  };
}
