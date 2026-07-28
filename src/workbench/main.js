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

function applyWorkbenchContext(ctx) {
  const chip = document.getElementById('wb-chip');
  const chipText = document.getElementById('wb-chip-text');
  const statusCtx = document.getElementById('wb-status-context');
  chip.classList.toggle('sp-chip-live', ctx.live);
  chip.classList.toggle('sp-chip-mock', !ctx.live);
  chipText.textContent = ctx.live ? 'SP: Live' : 'SP: Mock';
  chip.title = ctx.live
    ? `Connected to ${ctx.label}${ctx.user ? ` as ${ctx.user}` : ''} · context: ${ctx.source}`
    : 'Not connected to a SharePoint web — showing built-in mock data';
  statusCtx.textContent = ctx.live
    ? `SP: ${ctx.label}${ctx.user ? ` · ${ctx.user}` : ''}`
    : 'SP: mock data (deploy to SharePoint for live inspection)';
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

shell.restore();
