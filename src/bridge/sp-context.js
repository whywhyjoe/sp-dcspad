// SharePoint context bridge.
//
// When DCSPad is hosted on a real SharePoint page, the host's
// _spPageContextInfo is captured and injected into every preview run,
// and <base href> is pointed at the web — so PnPjs v2 and raw REST
// resolve URLs exactly as they would on a normal page in that web.
// Outside SharePoint (local dev), a clearly-flagged mock keeps the same
// shape so code doesn't need guards, but _api calls will 404.

let cached = null;

export function getSpContext() {
  if (cached) return cached;

  const real = window._spPageContextInfo;
  if (real && real.webAbsoluteUrl) {
    let pageContext;
    try {
      pageContext = JSON.parse(JSON.stringify(real));
    } catch {
      // Fall back to hand-picking the serializable essentials.
      pageContext = {};
      for (const k of ['webAbsoluteUrl', 'webServerRelativeUrl', 'siteAbsoluteUrl', 'siteServerRelativeUrl', 'webTitle', 'userId', 'userLoginName', 'userDisplayName', 'currentLanguage', 'currentCultureName', 'layoutsUrl', 'webUIVersion', 'siteClientTag', 'formDigestValue', 'formDigestTimeoutSeconds']) {
        if (real[k] !== undefined) pageContext[k] = real[k];
      }
    }
    // Classic pages keep a fresher digest in the form; prefer it.
    const digestEl = document.getElementById('__REQUESTDIGEST');
    if (digestEl?.value) pageContext.formDigestValue = digestEl.value;

    cached = {
      live: true,
      pageContext,
      baseHref: real.webAbsoluteUrl.replace(/\/$/, '') + '/',
      label: real.webAbsoluteUrl,
      user: real.userDisplayName || real.userLoginName || '',
    };
    return cached;
  }

  cached = {
    live: false,
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
    baseHref: null,   // keep relative URLs pointed at the local server
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
    ? `Connected to ${ctx.label} as ${ctx.user} — _spPageContextInfo is injected into every run`
    : 'Not hosted in SharePoint — a mock _spPageContextInfo (correct shape) is injected; _api calls will fail here';
  statusCtx.textContent = ctx.live
    ? `SP: ${ctx.label} · ${ctx.user}`
    : 'SP: mock context (deploy to SharePoint for live APIs)';
  return ctx;
}
