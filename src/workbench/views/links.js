// Config links view: curated one-click jumps to the SharePoint settings
// pages for the inspected web. Pure navigation — no REST calls, so the view
// works identically in mock and live modes (mock links just point at the
// local origin).

import { LINK_GROUPS, linkUrl } from '../config-links.js';

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

export function createLinksView({ client }) {
  const root = el('section', 'wb-view wb-view-links');
  const head = el('div', 'wb-view-head');
  head.innerHTML = '<h2>Panels</h2>'
    + '<p class="wb-view-hint">Quick jumps to the SharePoint configuration '
    + 'panels you actually reach for. Links open in a new tab; hover for the '
    + 'underlying page.</p>';
  const body = el('div', 'wb-links');
  root.append(head, body);

  // Rebuilt on every load() so site switching just works — shell.reset()
  // drops the instance anyway, but a cheap rebuild keeps this correct even
  // if the view is ever mounted without a reset.
  function load() {
    const webUrl = client.webUrl();
    body.textContent = '';
    for (const group of LINK_GROUPS) {
      const card = el('div', 'wb-linkgroup');
      card.append(el('h3', '', group.title));
      for (const link of group.links) {
        const row = el('a', 'wb-link');
        row.href = linkUrl(webUrl, link);
        row.target = '_blank';
        row.rel = 'noopener';
        // Label-first for scanning; the path (and any hint) lives in the
        // tooltip instead of competing with the label.
        row.append(el('span', 'wb-link-label', link.label));
        row.append(el('span', 'wb-link-go', '↗'));
        row.title = link.hint ? `${link.path}\n${link.hint}` : link.path;
        card.append(row);
      }
      body.append(card);
    }
  }

  return { el: root, load };
}
