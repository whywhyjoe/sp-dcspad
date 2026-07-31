// Durable site favorites + recents for the workbench site switcher.
//
// Persistence goes through state.js loadDoc/saveDoc — the sanctioned seam
// (invariant 6: state.js is the only module that touches localStorage), so
// favorites ride the future SharePoint storage swap for free. Importing
// state.js here is safe: its import-time side effects are a localStorage
// *read* plus a pagehide flush that only fires for pending workspace saves.
//
// Doc shape (loadDoc requires an items array):
//   { kind: 'dcspad-workbench-sites', version: 1,
//     items:   [{ id, url, title, addedAt }],   // favorites; url '' = host web
//     recents: [{ url, title, lastAt }] }       // most-recent first, capped

import { loadDoc, saveDoc, newId } from '../state.js';

export const FAVORITES_KEY = 'dcspad.v2.wbsites';
const RECENTS_CAP = 8;

function emptyDoc() {
  return { kind: 'dcspad-workbench-sites', version: 1, items: [], recents: [] };
}

function readDoc() {
  const doc = loadDoc(FAVORITES_KEY);
  if (!doc || doc.kind !== 'dcspad-workbench-sites') return emptyDoc();
  return {
    ...emptyDoc(),
    ...doc,
    items: Array.isArray(doc.items) ? doc.items : [],
    recents: Array.isArray(doc.recents) ? doc.recents : [],
  };
}

const canonical = (url) => String(url || '').replace(/\/+$/, '').toLowerCase();

let quotaListener = null;
export function onQuotaError(fn) { quotaListener = fn; }

function writeDoc(doc) {
  if (!saveDoc(FAVORITES_KEY, doc)) quotaListener?.();
  return doc;
}

export function getFavorites() { return readDoc().items; }
export function getRecents() { return readDoc().recents; }

export function isFavorite(url) {
  const key = canonical(url);
  return readDoc().items.some((item) => canonical(item.url) === key);
}

export function addFavorite({ url = '', title = '' } = {}) {
  const doc = readDoc();
  const key = canonical(url);
  if (doc.items.some((item) => canonical(item.url) === key)) return doc.items;
  doc.items.push({ id: newId('fav'), url: String(url || ''), title: String(title || ''), addedAt: new Date().toISOString() });
  return writeDoc(doc).items;
}

export function removeFavorite(url) {
  const doc = readDoc();
  const key = canonical(url);
  doc.items = doc.items.filter((item) => canonical(item.url) !== key);
  return writeDoc(doc).items;
}

export function pushRecent({ url = '', title = '' } = {}) {
  const doc = readDoc();
  const key = canonical(url);
  doc.recents = [
    { url: String(url || ''), title: String(title || ''), lastAt: new Date().toISOString() },
    ...doc.recents.filter((item) => canonical(item.url) !== key),
  ].slice(0, RECENTS_CAP);
  return writeDoc(doc).recents;
}
