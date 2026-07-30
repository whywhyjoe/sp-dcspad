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

export function buildTooltip(info = APP_BUILD_INFO) {
  const revision = info.revision ? ` (${info.revision})` : '';
  return `DCSPad — version ${info.version} — Build #${info.build}${revision}`;
}

export function applyBuildMarker(root = document) {
  const logo = root.querySelector('.logo');
  if (!logo) return;

  const tooltip = buildTooltip();
  logo.title = tooltip;
  logo.setAttribute('aria-label', tooltip);
  document.documentElement.dataset.dcspadVersion = APP_BUILD_INFO.version;
  document.documentElement.dataset.dcspadBuild = APP_BUILD_INFO.build;
  window.__DCSPAD_BUILD_INFO__ = APP_BUILD_INFO;
}
