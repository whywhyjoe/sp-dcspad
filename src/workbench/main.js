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

const GLYPHS = {
  lists: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" aria-hidden="true"><path d="M5.5 4h8M5.5 8h8M5.5 12h8"/><circle cx="2.7" cy="4" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="8" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="12" r=".9" fill="currentColor" stroke="none"/></svg>',
  security: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 1.8 13 3.6v3.6c0 3.2-2.1 5.6-5 6.9-2.9-1.3-5-3.7-5-6.9V3.6z"/><path d="m5.8 7.8 1.6 1.6 2.9-3"/></svg>',
  site: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" aria-hidden="true"><path d="M2.2 13.3V6.5L8 2.3l5.8 4.2v6.8z"/><path d="M6.2 13.3V9.4h3.6v3.9"/></svg>',
};

const SITE_KEY = 'dcspad.workbench.site';

function applyWorkbenchContext(ctx, inspecting = '') {
  const chip = document.getElementById('wb-chip');
  const chipText = document.getElementById('wb-chip-text');
  const statusCtx = document.getElementById('wb-status-context');
  chip.classList.toggle('sp-chip-live', ctx.live);
  chip.classList.toggle('sp-chip-mock', !ctx.live);
  chipText.textContent = ctx.live ? 'SP: Live' : 'SP: Mock';
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
  deps: { client },
  views: [
    { id: 'lists', label: 'Lists', glyph: GLYPHS.lists, create: createListsView },
    { id: 'security', label: 'Security', glyph: GLYPHS.security, create: createSecurityView },
    { id: 'site', label: 'Site', glyph: GLYPHS.site, create: createSiteView },
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
