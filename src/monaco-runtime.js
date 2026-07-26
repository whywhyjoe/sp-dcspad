// Monaco runtime loading and SharePoint-safe asset routing.
//
// Standalone mode resolves ../vendor/monaco/ from this module. Hosted mode
// gets an absolute SiteAssets base + vendor version from boot.js. Workers are
// always classic, same-origin .js files: no blob URLs, CDN, or .mjs MIME risk.

let runtimePromise = null;

function runtimeBase() {
  if (window.__DCSPAD_ASSET_BASE__) {
    return new URL('vendor/monaco/', window.__DCSPAD_ASSET_BASE__);
  }
  return new URL('../vendor/monaco/', import.meta.url);
}

function assetUrl(name) {
  const url = new URL(name, runtimeBase());
  const version = window.__DCSPAD_MONACO_VERSION__;
  if (version) url.searchParams.set('v', version);
  return url.href;
}

function ensureStylesheet() {
  const existing = document.getElementById('dcspad-monaco-style');
  if (existing) {
    if (existing.dataset.loaded === 'true') return Promise.resolve();
    return new Promise((resolve, reject) => {
      existing.addEventListener('load', resolve, { once: true });
      existing.addEventListener('error', () => reject(new Error('Monaco stylesheet failed to load')), { once: true });
    });
  }

  return new Promise((resolve, reject) => {
    const link = document.createElement('link');
    link.id = 'dcspad-monaco-style';
    link.rel = 'stylesheet';
    link.href = assetUrl('monaco.css');
    link.addEventListener('load', () => {
      link.dataset.loaded = 'true';
      resolve();
    }, { once: true });
    link.addEventListener('error', () => reject(new Error('Monaco stylesheet failed to load')), { once: true });
    document.head.appendChild(link);
  });
}

function configureWorkers() {
  self.MonacoEnvironment = {
    getWorker(_moduleId, label) {
      let file = 'editor.worker.js';
      if (label === 'css' || label === 'scss' || label === 'less') file = 'css.worker.js';
      else if (label === 'html' || label === 'handlebars' || label === 'razor') file = 'html.worker.js';
      else if (label === 'typescript' || label === 'javascript') file = 'ts.worker.js';

      const worker = new Worker(assetUrl(file), {
        name: `dcspad-monaco-${label || 'editor'}`,
      });
      worker.addEventListener('error', () => {
        document.documentElement.dataset.monacoWorkerError = label || 'editor';
        const status = document.getElementById('status-run');
        if (status) {
          status.textContent = 'editor worker unavailable — language tools limited';
          status.className = 'status-item error';
        }
      }, { once: true });
      return worker;
    },
  };
}

export function loadMonacoRuntime() {
  if (runtimePromise) return runtimePromise;
  runtimePromise = (async () => {
    configureWorkers();
    await ensureStylesheet();
    const monaco = await import(assetUrl('monaco.js'));
    document.documentElement.dataset.monacoReady = 'true';
    return monaco;
  })();
  return runtimePromise;
}

export async function fetchPnpTypeLibraries() {
  const response = await fetch(assetUrl('pnpjs-types.json'), {
    credentials: 'same-origin',
    cache: 'force-cache',
  });
  if (!response.ok) throw new Error(`HTTP ${response.status} loading PnPjs types`);
  const payload = await response.json();
  if (payload.version !== '2.15.0' || !Array.isArray(payload.libs)) {
    throw new Error('PnPjs type payload is invalid or does not match runtime 2.15.0');
  }
  return payload.libs;
}
