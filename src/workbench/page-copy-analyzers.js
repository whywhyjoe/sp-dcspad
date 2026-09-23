// Page copy analyzers — the registry (design/PAGE-COPY.md §5.1–5.2).
//
// The generic scanner in page-copy.js only DISCOVERS references; an analyzer
// is the only thing allowed to authorize a patch. Each one is selected by web
// part id, reads the instance's real configuration (property paths recorded
// from live pages during the spike, one fixture per shape), and classifies
// every reference it recognizes:
//
//   refs(instance, ctx) → [{ path, value, class, asset?, note }]
//     path    keys inside the instance (the raw canvas control, or the layout
//             part for the header)
//     class   'asset' | 'data' | 'config' | 'link' | 'dynamic'
//     asset   { path, ids: { siteId, webId, listId, uniqueId } } — the source
//             file, when class is 'asset'
//   patch(instance, ctx) → number of values changed (mutates the clone it is
//             handed). ctx.mapAsset(identity) → { path, ids, url } | null for
//             a transferred file; ctx.mapLink(value) → rewritten link | null
//             (null unless the operator opted into link rewriting).
//
// Unrecognized references stay unchanged and are reported by the scanner as
// "unverified"; an id with no analyzer here gets exactly that treatment.
// One module per web part under ./page-copy-analyzers/, so each shape is
// written, reviewed and fixture-tested on its own.

import text from './page-copy-analyzers/text.js';
import quickLinks from './page-copy-analyzers/quick-links.js';
import news from './page-copy-analyzers/news.js';
import listLibrary from './page-copy-analyzers/list-library.js';

export function createAnalyzers(list = []) {
  const byId = new Map();
  for (const analyzer of list) {
    if (analyzer.header || analyzer.text) continue;
    for (const id of [].concat(analyzer.id)) byId.set(String(id).toLowerCase(), analyzer);
  }
  return {
    header: list.find((a) => a.header) || null,
    text: list.find((a) => a.text) || null,
    forId: (webPartId) => byId.get(String(webPartId || '').toLowerCase()) || null,
    ids: () => [...byId.keys()],
  };
}

// The registry the dialog uses.
export const ANALYZERS = createAnalyzers([text, quickLinks, news, listLibrary]);
