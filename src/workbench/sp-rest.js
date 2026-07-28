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

const PAGE_CAP = 5000;          // max items accumulated across pages
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

export function createSpRestClient({
  getContext = getSpContext,
  fetchImpl = (...args) => fetch(...args),
  mockResolver = null,
} = {}) {
  let inFlight = 0;
  const waiters = [];

  async function withSlot(work) {
    if (inFlight >= MAX_CONCURRENT) {
      await new Promise((resolve) => waiters.push(resolve));
    }
    inFlight++;
    try { return await work(); }
    finally {
      inFlight--;
      waiters.shift()?.();
    }
  }

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

  function webUrl() {
    return context().pageContext.webAbsoluteUrl.replace(/\/+$/, '');
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

  // Full collection: follows paging links up to PAGE_CAP items.
  // Returns { items, partial } — partial=true means a paging link remained.
  async function getAll(path, opts) {
    let url = apiUrl(path, opts);
    const items = [];
    let partial = false;
    while (url) {
      const data = await rawGet(url);
      const page = collectionOf(data);
      if (!page) {
        // An entity endpoint queried through getAll — treat as one item.
        items.push(entityOf(data));
        break;
      }
      items.push(...page);
      const next = nextLinkOf(data);
      if (!next) break;
      if (items.length >= PAGE_CAP) { partial = true; break; }
      url = next;
    }
    return { items, partial };
  }

  return { context, webUrl, apiUrl, get, getAll };
}
