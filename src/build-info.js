// Application identity shown in the UI. Update APP_VERSION for a release.
// Hosted bundles replace the build/revision constants in build-app.mjs;
// unbundled local development deliberately identifies itself as Build #dev.

export const APP_VERSION = '1.0.0';

const injectedBuild = typeof __DCSPAD_BUILD_NUMBER__ === 'string'
  ? __DCSPAD_BUILD_NUMBER__
  : 'dev';
const injectedRevision = typeof __DCSPAD_BUILD_REVISION__ === 'string'
  ? __DCSPAD_BUILD_REVISION__
  : '';

export const APP_BUILD_INFO = Object.freeze({
  version: APP_VERSION,
  build: injectedBuild,
  revision: injectedRevision,
});

export function buildTooltipFor(appName, info = APP_BUILD_INFO) {
  const revision = info.revision ? ` (${info.revision})` : '';
  return `${appName} — version ${info.version} — Build #${info.build}${revision}`;
}

export function buildTooltip(info = APP_BUILD_INFO) {
  return buildTooltipFor('DCSPad', info);
}

// One console line at boot naming the exact running build — the fastest way
// to tell whether a tenant is serving a stale cached bundle.
export function logBuildInfo(appName, info = APP_BUILD_INFO) {
  console.info(`[${appName}] version ${info.version} — Build #${info.build}${info.revision ? ` (${info.revision})` : ''}`);
}

export function applyBuildMarker(root = document) {
  logBuildInfo('DCSPad');
  const logo = root.querySelector('.logo');
  if (!logo) return;

  const tooltip = buildTooltip();
  logo.title = tooltip;
  logo.setAttribute('aria-label', tooltip);
  document.documentElement.dataset.dcspadVersion = APP_BUILD_INFO.version;
  document.documentElement.dataset.dcspadBuild = APP_BUILD_INFO.build;
  window.__DCSPAD_BUILD_INFO__ = APP_BUILD_INFO;
}

// Workbench flavor: same contract, its own logo element and global.
export function applyWorkbenchBuildMarker(root = document) {
  logBuildInfo('SP Workbench');
  const tooltip = buildTooltipFor('SP Workbench');
  const logo = root.querySelector('.wb-logo');
  if (logo) {
    logo.title = tooltip;
    logo.setAttribute('aria-label', tooltip);
  }
  document.documentElement.dataset.dcspadWbBuild = APP_BUILD_INFO.build;
  window.__DCSPAD_WB_BUILD_INFO__ = APP_BUILD_INFO;
}
