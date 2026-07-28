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

const GLYPHS = {
  lists: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" aria-hidden="true"><path d="M5.5 4h8M5.5 8h8M5.5 12h8"/><circle cx="2.7" cy="4" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="8" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="12" r=".9" fill="currentColor" stroke="none"/></svg>',
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
    // M3: { id: 'security', label: 'Security', ... }
    // M4: { id: 'site', label: 'Site', ... }
  ],
});

shell.restore();
