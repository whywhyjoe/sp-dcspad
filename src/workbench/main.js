// SP Workbench bootstrap — the second DCSPad entry point.
//
// Wires context, REST client, shell, and views. Standalone workbench.html
// loads this unbundled; the hosted page loads dcspad.workbench.js (built by
// tools/build-workbench.mjs) through boot-workbench.js.

import { getSpContext } from '../bridge/sp-context.js';
import { createSpRestClient } from './sp-rest.js';
import { mockResolver } from './mock-data.js';
import { createShell } from './shell.js';
import { createListsView } from './views/lists.js';
import { createSecurityView } from './views/security.js';
import { createSiteView } from './views/site.js';
import { createSiteHomeView } from './views/site-home.js';
import { createLinksView } from './views/links.js';
import { createQueryView } from './views/query.js';
import { createPagesView } from './views/pages.js';
import { createBrowserView } from './views/browser.js';
import {
  getFavorites, getRecents, isFavorite, addFavorite, removeFavorite,
  pushRecent, onQuotaError,
} from './favorites.js';

const GLYPHS = {
  lists: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" aria-hidden="true"><path d="M5.5 4h8M5.5 8h8M5.5 12h8"/><circle cx="2.7" cy="4" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="8" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="12" r=".9" fill="currentColor" stroke="none"/></svg>',
  security: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 1.8 13 3.6v3.6c0 3.2-2.1 5.6-5 6.9-2.9-1.3-5-3.7-5-6.9V3.6z"/><path d="m5.8 7.8 1.6 1.6 2.9-3"/></svg>',
  site: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" aria-hidden="true"><path d="M2.2 13.3V6.5L8 2.3l5.8 4.2v6.8z"/><path d="M6.2 13.3V9.4h3.6v3.9"/></svg>',
  links: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6.5 9.5 9.5 6.5"/><path d="M7.5 4.6 9 3.1a2.6 2.6 0 0 1 3.7 0l.2.2a2.6 2.6 0 0 1 0 3.7L11.4 8.5"/><path d="M8.5 11.4 7 12.9a2.6 2.6 0 0 1-3.7 0l-.2-.2a2.6 2.6 0 0 1 0-3.7L4.6 7.5"/></svg>',
  query: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="7" cy="7" r="4.4"/><path d="m13.5 13.5-3.2-3.2"/><path d="M5.2 7h3.6M7 5.2v3.6"/></svg>',
  pages: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3.4 1.8h6.4l2.8 2.8v9.6H3.4z"/><path d="M9.6 1.8v3h3"/><path d="M5.4 8h5.2M5.4 10.4h5.2"/></svg>',
  files: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M1.8 4.2v8.2h12.4V5.8H7.6L6.2 4.2H1.8z"/><path d="M1.8 4.2V2.9h4.4l1.4 1.6"/></svg>',
  advanced: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="8" cy="8" r="2.2"/><path d="M8 1.9v2M8 12.1v2M1.9 8h2M12.1 8h2M3.7 3.7l1.4 1.4M10.9 10.9l1.4 1.4M12.3 3.7l-1.4 1.4M5.1 10.9l-1.4 1.4"/></svg>',
};

const SITE_KEY = 'dcspad.workbench.site';

function applyWorkbenchContext(ctx, inspecting = '') {
  const chip = document.getElementById('wb-chip');
  const chipText = document.getElementById('wb-chip-text');
  const statusCtx = document.getElementById('wb-status-context');
  chip.classList.toggle('sp-chip-live', ctx.live);
  chip.classList.toggle('sp-chip-mock', !ctx.live);
  chipText.textContent = ctx.live ? 'SP' : 'SP: Mock';
  chip.title = ctx.live
    ? `Connected to ${ctx.label}${ctx.user ? ` as ${ctx.user}` : ''} · context: ${ctx.source}`
    : 'Not connected to a SharePoint web — showing built-in mock data';
  const inspectingNote = inspecting ? ` · inspecting ${inspecting}` : '';
  statusCtx.textContent = ctx.live
    ? `SP: ${ctx.label}${ctx.user ? ` · ${ctx.user}` : ''}${inspectingNote}`
    : `SP: mock data (deploy to SharePoint for live inspection)${inspectingNote}`;
}

const ctx = getSpContext();
applyWorkbenchContext(ctx);

const client = createSpRestClient({
  mockResolver: ctx.live ? null : mockResolver,
});

const shell = createShell({
  mount: document.getElementById('wb-main'),
  // inspectSite is a hoisted declaration below; views get a late-bound ref.
  deps: { client, inspectSite: (url) => inspectSite(url) },
  views: [
    // Nav order and grouping are Joe's spec (2026-07-31): identity first,
    // then content, then query, then jump-off/diagnostic sections.
    { id: 'site', label: 'Site', glyph: GLYPHS.site, group: 'identity', create: createSiteHomeView },
    { id: 'security', label: 'Permissions', glyph: GLYPHS.security, group: 'identity', create: createSecurityView },
    { id: 'lists', label: 'Lists', glyph: GLYPHS.lists, group: 'content', create: createListsView },
    { id: 'pages', label: 'Pages', glyph: GLYPHS.pages, group: 'content', create: createPagesView },
    { id: 'files', label: 'Files', glyph: GLYPHS.files, group: 'content', create: createBrowserView },
    { id: 'query', label: 'Query', glyph: GLYPHS.query, group: 'query', create: createQueryView },
    { id: 'links', label: 'Panels', glyph: GLYPHS.links, group: 'jump', create: createLinksView },
    { id: 'advanced', label: 'Advanced', glyph: GLYPHS.advanced, group: 'jump', create: createSiteView },
  ],
});

// ---- site switcher --------------------------------------------------------

const siteForm = document.getElementById('wb-site-form');
const siteInput = document.getElementById('wb-site-input');
const siteOpen = document.getElementById('wb-site-open');
const siteError = document.getElementById('wb-site-error');

function rememberSite(value) {
  try {
    if (value) sessionStorage.setItem(SITE_KEY, value);
    else sessionStorage.removeItem(SITE_KEY);
  } catch { /* private mode */ }
}

async function inspectSite(input, { reset = true } = {}) {
  siteError.hidden = true;
  siteOpen.disabled = true;
  siteOpen.textContent = 'Opening…';
  try {
    const web = await client.connectWeb(input);
    const inspectingHost = client.webUrl() === client.hostWebUrl();
    siteInput.value = inspectingHost ? '' : client.webUrl();
    rememberSite(inspectingHost ? '' : client.webUrl());
    applyWorkbenchContext(ctx, inspectingHost ? '' : `${web?.Title || 'web'} (${client.webUrl()})`);
    currentSite = {
      url: inspectingHost ? '' : client.webUrl(),
      title: web?.Title || (inspectingHost ? 'This site' : client.webUrl()),
    };
    pushRecent(currentSite);
    refreshFavStar();
    refreshCurrentUser();   // admin state can differ per web
    if (reset) shell.reset();
    return true;
  } catch (err) {
    siteError.textContent = err?.message || String(err);
    siteError.hidden = false;
    return false;
  } finally {
    siteOpen.disabled = false;
    siteOpen.textContent = 'Inspect';
  }
}

siteForm.addEventListener('submit', (e) => {
  e.preventDefault();
  inspectSite(siteInput.value);
});

// ---- site favorites + recents ---------------------------------------------

const favBtn = document.getElementById('wb-site-fav');
const favListBtn = document.getElementById('wb-site-favlist');
const favMenu = document.getElementById('wb-site-menu');
const statusCtx = document.getElementById('wb-status-context');

// The web currently inspected; url '' means the host web.
let currentSite = { url: '', title: 'This site' };

onQuotaError(() => {
  const previous = statusCtx.textContent;
  statusCtx.textContent = 'Could not save workbench favorites (storage quota).';
  setTimeout(() => { statusCtx.textContent = previous; }, 4000);
});

function refreshFavStar() {
  const fav = isFavorite(currentSite.url);
  favBtn.textContent = fav ? '★' : '☆';
  favBtn.classList.toggle('active', fav);
  favBtn.setAttribute('aria-pressed', fav ? 'true' : 'false');
  favBtn.title = fav
    ? 'Remove the inspected site from favorites'
    : 'Favorite the inspected site';
}

favBtn.addEventListener('click', () => {
  if (isFavorite(currentSite.url)) removeFavorite(currentSite.url);
  else addFavorite(currentSite);
  refreshFavStar();
});

function siteMenuItem(entry, hint) {
  const item = document.createElement('button');
  item.type = 'button';
  item.className = 'wb-menu-item';
  const label = entry.url ? (entry.title || entry.url) : (entry.title || 'This site');
  item.textContent = entry.url ? `${label} — ${entry.url}` : `${label} (host web)`;
  if (hint) item.title = hint;
  item.addEventListener('click', () => {
    favMenu.hidden = true;
    siteInput.value = entry.url;
    inspectSite(entry.url);
  });
  return item;
}

function menuHeading(text) {
  const h = document.createElement('div');
  h.className = 'wb-menu-heading';
  h.textContent = text;
  return h;
}

function rebuildSiteMenu() {
  favMenu.textContent = '';
  const favorites = getFavorites();
  const recents = getRecents();
  if (favorites.length) {
    favMenu.append(menuHeading('Favorites'));
    for (const entry of favorites) favMenu.append(siteMenuItem(entry));
  }
  const favUrls = new Set(favorites.map((f) => (f.url || '').toLowerCase()));
  const rest = recents.filter((r) => !favUrls.has((r.url || '').toLowerCase()));
  if (rest.length) {
    favMenu.append(menuHeading('Recent'));
    for (const entry of rest) favMenu.append(siteMenuItem(entry));
  }
  if (!favMenu.childElementCount) {
    const empty = document.createElement('div');
    empty.className = 'wb-menu-empty';
    empty.textContent = 'No favorite or recent sites yet.';
    favMenu.append(empty);
  }
}

favListBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  if (favMenu.hidden) rebuildSiteMenu();
  favMenu.hidden = !favMenu.hidden;
});
document.addEventListener('click', () => { favMenu.hidden = true; });

refreshFavStar();

// ---- status bar: current user + role chip ---------------------------------

const statusUser = document.getElementById('wb-status-user');
const statusRole = document.getElementById('wb-status-role');

async function refreshCurrentUser() {
  try {
    const user = await client.get('web/currentuser', {
      select: ['Title', 'Email', 'IsSiteAdmin'],
    });
    statusUser.textContent = user.Email
      ? `${user.Title} · ${user.Email}` : (user.Title || '');
    statusRole.textContent = user.IsSiteAdmin ? 'Site admin' : 'Site user';
    statusRole.className = user.IsSiteAdmin
      ? 'wb-role-chip wb-role-admin' : 'wb-role-chip wb-role-user';
    statusRole.hidden = false;
  } catch {
    statusUser.textContent = '';
    statusRole.hidden = true;
  }
}

refreshCurrentUser();

// Boot: reconnect this tab's last inspected site before the first view loads,
// falling back to the host web (and a clean route) if it no longer resolves.
(async () => {
  let saved = '';
  try { saved = sessionStorage.getItem(SITE_KEY) || ''; } catch { /* ignore */ }
  if (saved) {
    siteInput.value = saved;
    const ok = await inspectSite(saved, { reset: false });
    if (!ok) {
      rememberSite('');
      siteInput.value = '';
      try { sessionStorage.removeItem('dcspad.workbench.route'); } catch { /* ignore */ }
    }
  }
  shell.restore();
})();
