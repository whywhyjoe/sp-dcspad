// SP Workbench REST client — read-only /_api access for the inspector views.
//
// Raw fetch, no runtime dependency: pnpjs is an *output format* of the script
// generator, never a data layer. Shares the OData plumbing with sp-files.js
// via sp-odata.js. v1 is GET-only, so no request digest is needed; when edit
// tools arrive they should reuse the digest cache in sp-files.js.
//
// Off SharePoint the client runs against an injected mock resolver so every
// view stays exercisable (and testable) with zero network.

import { getSpContext } from '../bridge/sp-context.js';
import { ACCEPT_JSON, SpFileError, requireOk } from '../sp-odata.js';

const PAGE_CAP = 5000;          // max items accumulated across pages (default ceiling)
const LARGE_PAGE_CAP = 100000;  // opt-in ceiling — see getAll's allowLargeCap
const MAX_CONCURRENT = 3;       // polite ceiling for parallel view loads
const RETRY_STATUSES = new Set([429, 503]);

function buildQuery({ select, expand, filter, orderby, top } = {}) {
  const parts = [];
  const join = (v) => (Array.isArray(v) ? v.join(',') : String(v));
  if (select) parts.push(`$select=${join(select)}`);
  if (expand) parts.push(`$expand=${join(expand)}`);
  if (filter) parts.push(`$filter=${encodeURIComponent(String(filter))}`);
  if (orderby) parts.push(`$orderby=${join(orderby)}`);
  if (top) parts.push(`$top=${top}`);
  return parts.length ? `?${parts.join('&')}` : '';
}

// Collection responses arrive as nometadata {value:[]}, verbose {d:{results:[]}},
// or (from mocks/stubs) bare arrays. Entities arrive bare or under d.
function collectionOf(data) {
  if (Array.isArray(data)) return data;
  if (Array.isArray(data?.value)) return data.value;
  if (Array.isArray(data?.d?.results)) return data.d.results;
  if (Array.isArray(data?.results)) return data.results;
  return null;
}

function nextLinkOf(data) {
  return data?.['odata.nextLink'] || data?.['@odata.nextLink'] || data?.d?.__next || '';
}

function entityOf(data) {
  return data?.d ?? data;
}

// One request queue for the whole app, shared by every client. A second
// client (createClient(): the schema dialog's target web, the EEEU audit's
// subsites) used to get its own three slots, so the app could run six or more
// requests at once against one tenant — the throttling this ceiling exists to
// avoid. withSlot never nests (a slot holder never waits on another slot), so
// sharing cannot deadlock.
//
// A finished request hands its slot straight to the next waiter instead of
// releasing it: releasing first let a newcomer take the slot before the woken
// waiter ran, briefly putting MAX_CONCURRENT + 1 requests in flight.
let inFlight = 0;
const waiters = [];

async function withSlot(work) {
  if (inFlight >= MAX_CONCURRENT) {
    await new Promise((resolve) => waiters.push(resolve));   // slot handed over
  } else {
    inFlight++;
  }
  try { return await work(); }
  finally {
    const next = waiters.shift();
    if (next) next();
    else inFlight--;
  }
}

export function createSpRestClient({
  getContext = getSpContext,
  fetchImpl = (...args) => fetch(...args),
  mockResolver = null,
} = {}) {

  let targetWebUrl = '';   // '' = the host web the workbench runs on

  function context() {
    const ctx = getContext();
    if (!ctx?.live && !mockResolver) {
      throw new SpFileError(
        'The SP Workbench needs an SP: Live context (or an injected mock).',
        { code: 'not-live' },
      );
    }
    return ctx;
  }

  function hostWebUrl() {
    return context().pageContext.webAbsoluteUrl.replace(/\/+$/, '');
  }

  function webUrl() {
    return targetWebUrl || hostWebUrl();
  }

  // Same-tenant only: a server-relative path ("/sites/Project") or an
  // absolute URL on the host's origin. Anything else is rejected — the
  // workbench authenticates with the page's own cookies, which don't
  // travel cross-origin.
  function normalizeTarget(input) {
    const raw = String(input || '').trim();
    if (!raw) return '';
    const host = hostWebUrl();
    let candidate;
    try {
      candidate = new URL(raw, host);
    } catch {
      throw new SpFileError(
        'Enter a site URL on this tenant, such as /sites/ProjectName.',
        { code: 'invalid-web-url' },
      );
    }
    if (!/^https?:$/.test(candidate.protocol)
        || candidate.origin !== new URL(host).origin) {
      throw new SpFileError(
        'That URL is on a different tenant — the workbench can only inspect sites on its own origin.',
        { code: 'invalid-web-url' },
      );
    }
    candidate.hash = '';
    candidate.search = '';
    return candidate.href.replace(/\/+$/, '');
  }

  // Validate a target web by asking it for /_api/web, then make it the web
  // every later call inspects. Empty input returns to the host web.
  // Resolves to the web entity ({ Title, Url, ServerRelativeUrl, ... }).
  async function connectWeb(input) {
    const candidate = normalizeTarget(input);
    if (!candidate) {
      targetWebUrl = '';
      return entityOf(await rawGet(`${hostWebUrl()}/_api/web?$select=Id,Title,Url,ServerRelativeUrl`));
    }
    const web = entityOf(await rawGet(`${candidate}/_api/web?$select=Id,Title,Url,ServerRelativeUrl`));
    // Prefer the canonical URL SharePoint reports (fixes casing, trailing
    // segments); fall back to the candidate for mocks that omit Url.
    targetWebUrl = normalizeTarget(web?.Url) || candidate;
    return web;
  }

  function apiUrl(path, opts) {
    const clean = String(path).replace(/^\/+/, '');
    return `${webUrl()}/_api/${clean}${buildQuery(opts)}`;
  }

  async function rawGet(url) {
    if (mockResolver && !getContext().live) {
      const data = mockResolver(url);
      if (data == null) {
        throw new SpFileError(`No mock data for ${url}`, { code: 'not-found', status: 404 });
      }
      return structuredClone(data);
    }
    return withSlot(async () => {
      const attempt = async () => {
        let response;
        try {
          response = await fetchImpl(url, {
            credentials: 'same-origin',
            headers: { Accept: ACCEPT_JSON },
          });
        } catch (cause) {
          throw new SpFileError(
            `Could not reach SharePoint (${cause.message || cause}).`,
            { code: 'network', cause },
          );
        }
        return response;
      };
      let response = await attempt();
      if (RETRY_STATUSES.has(response.status)) {
        const after = Number(response.headers.get('Retry-After')) || 2;
        await new Promise((r) => setTimeout(r, Math.min(after, 30) * 1000));
        response = await attempt();
      }
      await requireOk(response, 'SharePoint request failed', 'get');
      return response.json();
    });
  }

  // Single entity (or raw endpoint payload). `path` is relative to /_api/.
  async function get(path, opts) {
    return entityOf(await rawGet(apiUrl(path, opts)));
  }

  // ONE request that never follows a paging link: { items, nextLink }. Use it
  // for a `top: 1` style lookup (newest row, oldest row, "is there any row?").
  // getAll treats `top` as the PAGE size, not a total, so
  // getAll(path, { orderby, top: 1 }) on a large list walks one-row pages up
  // to the 5000 cap. Measured live at 5,011 requests (sp-traffic-analytics,
  // dev tenant, 2026-10). Passing getAll a `cap` also stops it, but a lookup
  // should not be one forgotten option away from five thousand requests.
  async function getPage(path, opts) {
    const data = await rawGet(apiUrl(path, opts));
    const page = collectionOf(data);
    return page
      ? { items: page, nextLink: nextLinkOf(data) }
      : { items: [entityOf(data)], nextLink: '' };
  }

  // Full collection: follows paging links up to `cap` items (default and
  // ceiling PAGE_CAP — callers can only lower it, e.g. a max-items export).
  // `allowLargeCap` is an explicit, opt-in escape hatch for one caller (the
  // list-data capture, which can legitimately need every item in a large
  // list): it raises the ceiling to LARGE_PAGE_CAP without changing the
  // default 5000 ceiling anyone else sees. Returns { items, partial } —
  // partial=true means more rows remained.
  // No `cap` means the ceiling itself — with allowLargeCap that is
  // LARGE_PAGE_CAP. (A `cap = PAGE_CAP` default here once silently held an
  // allowLargeCap caller that passed no cap — the EEEU item scan — to 5000.)
  // `shouldStop` (optional) is checked between pages: a cancelled caller gets
  // what was read so far with partial=true and stopped=true instead of
  // waiting out the whole collection (the EEEU item scan of a 34k-item
  // library otherwise held Cancel for ~2.5 minutes on live SPO).
  //
  // `opts.top` is the page size requested from SharePoint, NOT a total. The
  // large-cap readers (the EEEU item scan, the list-data capture) pass
  // top: 5000 and rely on every later page being followed. For a fixed
  // number of rows, use getPage (one request) or pass `cap`.
  //
  // Large lists: SharePoint Online refuses ANY query whose LEADING indexed
  // filter condition matches more than 5,000 rows (the list view threshold).
  // $top doesn't help, and neither do later clauses or an Id window that
  // narrow the result. Verified live on a 6,201-item list, 2026-10. A narrow
  // indexed range placed FIRST in $filter is served. Unfiltered paging works
  // at any size; a filtered read of a busy list must lead with a bounded
  // indexed range and read it in slices (sp-traffic-analytics
  // src/aggregate.js getRowsInRange: DateTime slices that halve on refusal).
  async function getAll(path, opts, { cap, allowLargeCap = false, shouldStop = null } = {}) {
    const ceiling = allowLargeCap ? LARGE_PAGE_CAP : PAGE_CAP;
    const limit = Math.min(Math.max(1, Number(cap) || ceiling), ceiling);
    let url = apiUrl(path, opts);
    const items = [];
    let partial = false;
    let stopped = false;
    while (url) {
      if (shouldStop?.()) { partial = true; stopped = true; break; }
      const data = await rawGet(url);
      const page = collectionOf(data);
      if (!page) {
        // An entity endpoint queried through getAll — treat as one item.
        items.push(entityOf(data));
        break;
      }
      const remaining = limit - items.length;
      if (page.length > remaining) {
        items.push(...page.slice(0, remaining));
        partial = true;
        break;
      }
      items.push(...page);
      const next = nextLinkOf(data);
      if (!next) break;
      if (items.length >= limit) { partial = true; break; }
      url = next;
    }
    return { items, partial, stopped };
  }

  return { context, webUrl, hostWebUrl, connectWeb, apiUrl, get, getPage, getAll };
}
