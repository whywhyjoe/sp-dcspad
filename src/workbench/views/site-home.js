// Site view — the workbench landing page. A friendly at-a-glance overview:
// who and where you are, the essential web facts, and subwebs you can jump
// into with one click. The exhaustive property sheets live under Advanced.

import { createGrid } from '../grid.js';
import { copyText } from '../export.js';

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

const fmtDate = (v) => (v ? String(v).slice(0, 10) : '');

export function createSiteHomeView({ client, navigate, inspectSite }) {
  const root = el('section', 'wb-view wb-view-sitehome');
  const head = el('div', 'wb-view-head');
  head.innerHTML = '<h2>Site</h2>'
    + '<p class="wb-view-hint">The inspected web at a glance. '
    + 'Full property sheets are under Advanced.</p>';

  const cards = el('div', 'wb-home-cards');
  const webCard = el('div', 'wb-home-card');
  const userCard = el('div', 'wb-home-card');
  cards.append(webCard, userCard);

  const subwebsBox = el('div', 'wb-home-subwebs');
  root.append(head, cards, subwebsBox);

  let loadedForWeb = '';

  function factRow(label, value, { copyFull = '' } = {}) {
    const row = el('div', 'wb-home-fact');
    row.append(el('span', 'wb-home-fact-label', label));
    const v = el('span', copyFull ? 'wb-home-fact-value sp-copy' : 'wb-home-fact-value', value || '—');
    if (copyFull) {
      v.title = 'Click to copy the full URL';
      v.addEventListener('click', () => copyText(copyFull, v));
    }
    row.append(v);
    return row;
  }

  async function load() {
    const webUrl = client.webUrl();
    if (loadedForWeb === webUrl) return;
    loadedForWeb = webUrl;

    webCard.textContent = '';
    webCard.append(el('h3', 'wb-home-card-title', 'This web'));
    userCard.textContent = '';
    userCard.append(el('h3', 'wb-home-card-title', 'You'));
    subwebsBox.textContent = '';

    try {
      const web = await client.get('web', {
        select: [
          'Title', 'Description', 'Url', 'ServerRelativeUrl', 'WebTemplate',
          'Created', 'LastItemModifiedDate', 'Language',
        ],
      });
      webCard.append(
        factRow('Title', web.Title),
        factRow('Description', web.Description),
        factRow('URL', web.ServerRelativeUrl || '/', { copyFull: web.Url || webUrl }),
        factRow('Template', web.WebTemplate),
        factRow('Created', fmtDate(web.Created)),
        factRow('Last modified', fmtDate(web.LastItemModifiedDate)),
      );
    } catch (err) {
      webCard.append(el('div', 'wb-grid-status wb-error', err?.message || String(err)));
    }

    try {
      const user = await client.get('web/currentuser', {
        select: ['Title', 'Email', 'LoginName', 'IsSiteAdmin'],
      });
      userCard.append(
        factRow('Name', user.Title),
        factRow('Email', user.Email),
        factRow('Login', user.LoginName),
      );
      const roleRow = el('div', 'wb-home-fact');
      roleRow.append(el('span', 'wb-home-fact-label', 'Role'));
      roleRow.append(el('span',
        user.IsSiteAdmin ? 'wb-role-chip wb-role-admin' : 'wb-role-chip wb-role-user',
        user.IsSiteAdmin ? 'Site admin' : 'Site user'));
      userCard.append(roleRow);
    } catch (err) {
      userCard.append(el('div', 'wb-grid-status wb-error', err?.message || String(err)));
    }

    // Subwebs with one-click inspection.
    const grid = createGrid({
      columns: [
        { key: 'Title', label: 'Subweb' },
        { key: 'ServerRelativeUrl', label: 'Url', mono: true, copyable: true },
        { key: 'Created', label: 'Created', format: fmtDate },
        {
          key: 'Inspect',
          label: '',
          value: (row) => row.ServerRelativeUrl,
          format: () => '',
          render: (url) => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-xs';
            btn.textContent = 'Inspect';
            btn.addEventListener('click', (e) => {
              e.stopPropagation();
              inspectSite?.(url);
            });
            return btn;
          },
        },
      ],
      emptyText: 'No subwebs under this web.',
      filterPlaceholder: 'Filter subwebs…',
      exportName: 'sp-subwebs',
      descriptor: {
        path: 'web/webs',
        options: { select: ['Id', 'Title', 'ServerRelativeUrl', 'Created', 'WebTemplate'] },
        webUrl: client.webUrl(),
      },
    });
    subwebsBox.append(el('h3', 'wb-home-card-title', 'Subwebs'), grid.el);
    grid.setLoading('Loading subwebs…');
    client.getAll('web/webs', {
      select: ['Id', 'Title', 'ServerRelativeUrl', 'Created', 'WebTemplate'],
    })
      .then(({ items, partial }) => grid.setRows(items, { partial }))
      .catch((err) => grid.setError(err));
  }

  // Site switches drop the instance via shell.reset(), but stay correct if
  // load() ever runs again for a different web.
  function loadRoute() {
    if (loadedForWeb && loadedForWeb !== client.webUrl()) loadedForWeb = '';
    load();
  }

  return { el: root, load: loadRoute };
}
