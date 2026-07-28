// SharePoint context bridge.
//
// DCSPad accepts a deliberately small host contract before falling back to
// SharePoint globals:
//
//   window.__DCSPAD_SP_CONTEXT__ = {
//     webAbsoluteUrl: 'https://tenant.sharepoint.com/sites/example'
//   };
//
// A custom script editor or SPFx host may also include pageContext and user
// display fields. The web URL is the only required value; sp-files.js obtains
// a fresh digest and canonical web/site URLs from /_api/contextinfo.

export const MODERN_SITE_PAGES_FEATURE_ID =
  'b6917cb1-93a0-4b97-a84d-7cf49975d4ec';

const SERIALIZABLE_FIELDS = [
  'webAbsoluteUrl', 'webServerRelativeUrl',
  'siteAbsoluteUrl', 'siteServerRelativeUrl',
  'webTitle', 'userId', 'userLoginName', 'userDisplayName',
  'currentLanguage', 'currentCultureName', 'layoutsUrl', 'webUIVersion',
  'siteClientTag', 'formDigestValue', 'formDigestTimeoutSeconds',
];

let cached = null;

const isRecord = (value) =>
  value !== null && typeof value === 'object' && !Array.isArray(value);

function safeSameOriginUrl(value) {
  if (typeof value !== 'string' || !value.trim()) return '';
  try {
    const url = new URL(value.trim(), location.href);
    if (url.origin !== location.origin) return '';
    return url.href.replace(/\/+$/, '');
  } catch {
    return '';
  }
}

function serverRelativeUrl(absoluteUrl) {
  try {
    return decodeURIComponent(new URL(absoluteUrl).pathname).replace(/\/+$/, '') || '/';
  } catch {
    return '/';
  }
}

function candidateWindows() {
  const candidates = [window];
  for (const key of ['parent', 'top']) {
    try {
      const candidate = window[key];
      if (candidate && !candidates.includes(candidate)) {
        // Reading location.href forces a same-origin check now, so every
        // later property access can remain a simple guarded lookup.
        void candidate.location.href;
        candidates.push(candidate);
      }
    } catch {
      // Cross-origin frames are expected; continue with the current window.
    }
  }
  return candidates;
}

function hostContext(candidate) {
  try {
    const host = candidate.__DCSPAD_SP_CONTEXT__;
    if (!isRecord(host)) return null;
    const pageContext = isRecord(host.pageContext) ? host.pageContext : host;
    const webAbsoluteUrl = safeSameOriginUrl(
      host.webAbsoluteUrl || pageContext.webAbsoluteUrl,
    );
    return webAbsoluteUrl ? { raw: host, pageContext, webAbsoluteUrl } : null;
  } catch {
    return null;
  }
}

function globalContext(candidate) {
  try {
    const pageContext = candidate._spPageContextInfo;
    const webAbsoluteUrl = safeSameOriginUrl(pageContext?.webAbsoluteUrl);
    return webAbsoluteUrl ? { raw: pageContext, pageContext, webAbsoluteUrl } : null;
  } catch {
    return null;
  }
}

function modernLegacyContext(candidate) {
  try {
    const pageContext = candidate.spModuleLoader
      ?._bundledComponents
      ?.[MODERN_SITE_PAGES_FEATURE_ID]
      ?.PageManager
      ?._instance
      ?.pageContext
      ?.legacyPageContext;
    const webAbsoluteUrl = safeSameOriginUrl(pageContext?.webAbsoluteUrl);
    return webAbsoluteUrl ? { raw: pageContext, pageContext, webAbsoluteUrl } : null;
  } catch {
    return null;
  }
}

function findContext() {
  const windows = candidateWindows();
  for (const [source, reader] of [
    ['host', hostContext],
    ['global', globalContext],
    ['modern-legacy', modernLegacyContext],
  ]) {
    for (const candidate of windows) {
      const found = reader(candidate);
      if (found) return { ...found, source, ownerWindow: candidate };
    }
  }
  return null;
}

function copyPageContext(found) {
  let pageContext;
  try {
    pageContext = JSON.parse(JSON.stringify(found.pageContext));
  } catch {
    pageContext = {};
    for (const key of SERIALIZABLE_FIELDS) {
      if (found.pageContext[key] !== undefined) {
        pageContext[key] = found.pageContext[key];
      }
    }
  }

  // Host adapters commonly put optional display fields beside pageContext.
  for (const key of SERIALIZABLE_FIELDS) {
    if (pageContext[key] === undefined && found.raw[key] !== undefined) {
      pageContext[key] = found.raw[key];
    }
  }

  pageContext.webAbsoluteUrl = found.webAbsoluteUrl;
  pageContext.webServerRelativeUrl ||= serverRelativeUrl(found.webAbsoluteUrl);
  pageContext.siteAbsoluteUrl ||= found.webAbsoluteUrl;
  pageContext.siteServerRelativeUrl ||=
    serverRelativeUrl(pageContext.siteAbsoluteUrl);

  // Classic pages keep a fresher digest in the form; prefer it when the
  // context came from the same accessible document.
  try {
    const digest = found.ownerWindow.document.getElementById('__REQUESTDIGEST')?.value;
    if (digest) pageContext.formDigestValue = digest;
  } catch {
    // The candidate was same-origin when captured but may have navigated.
  }

  return pageContext;
}

// Pass { refresh: true } to re-capture from the host page. run() does this on
// every preview run so the injected context follows SharePoint SPA navigation
// and classic-page digest refreshes.
export function getSpContext({ refresh = false } = {}) {
  if (cached && !refresh) return cached;

  const found = findContext();
  if (found) {
    const pageContext = copyPageContext(found);
    cached = {
      live: true,
      source: found.source,
      capturedAt: Date.now(),
      pageContext,
      baseHref: `${pageContext.webAbsoluteUrl.replace(/\/$/, '')}/`,
      label: pageContext.webAbsoluteUrl,
      user: pageContext.userDisplayName || pageContext.userLoginName || '',
    };
    return cached;
  }

  cached = {
    live: false,
    source: 'mock',
    capturedAt: Date.now(),
    pageContext: {
      isDcsPadMock: true,
      webAbsoluteUrl: location.origin,
      webServerRelativeUrl: '/',
      siteAbsoluteUrl: location.origin,
      siteServerRelativeUrl: '/',
      webTitle: 'DCSPad Mock Web',
      userId: 1,
      userLoginName: 'i:0#.f|membership|dev@mock.local',
      userDisplayName: 'Mock Developer',
      currentLanguage: 1033,
      currentCultureName: 'en-US',
      layoutsUrl: '_layouts/15',
      formDigestValue: 'MOCK-DIGEST-0x0000',
      formDigestTimeoutSeconds: 1800,
    },
    baseHref: null,
    label: 'mock (not in SharePoint)',
    user: 'Mock Developer',
  };
  return cached;
}

export function applyContextIndicators() {
  const ctx = getSpContext();
  const chip = document.getElementById('sp-chip');
  const chipText = document.getElementById('sp-chip-text');
  const statusCtx = document.getElementById('status-context');

  chip.classList.toggle('sp-chip-live', ctx.live);
  chip.classList.toggle('sp-chip-mock', !ctx.live);
  chipText.textContent = ctx.live ? 'SP: Live' : 'SP: Mock';
  chip.title = ctx.live
    ? `Connected to ${ctx.label}${ctx.user ? ` as ${ctx.user}` : ''} · context: ${ctx.source}`
    : 'Not connected to a SharePoint web — SharePoint file actions are unavailable';
  statusCtx.textContent = ctx.live
    ? `SP: ${ctx.label}${ctx.user ? ` · ${ctx.user}` : ''}`
    : 'SP: mock context (deploy to SharePoint for live APIs)';
  return ctx;
}
