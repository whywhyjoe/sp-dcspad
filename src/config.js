// Runtime configuration loader. The editable dcspad.config.json document is
// deliberately separate from the app bundle so SharePoint-hosted framework
// and design-system locations can change without rebuilding source.

const EMPTY_CONFIG = Object.freeze({
  version: 1,
  frameworks: Object.freeze({
    prefer: 'local',
    fallbackToCdn: true,
    items: Object.freeze({}),
  }),
  assets: Object.freeze({}),
});

let activeConfig = EMPTY_CONFIG;

const isRecord = (value) =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const cleanString = (value) => typeof value === 'string' ? value.trim() : '';
const sourcePreference = (value, fallback = 'local') =>
  value === 'cdn' || value === 'hosted' || value === 'local' ? value : fallback;

function resolveUrl(value, configUrl, { folder = false } = {}) {
  const source = cleanString(value);
  if (!source) return '';
  const resolved = new URL(source, configUrl).href;
  return folder && !resolved.endsWith('/') ? `${resolved}/` : resolved;
}

function normalizeFrameworks(value, configUrl, warnings) {
  const source = isRecord(value) ? value : {};
  const prefer = sourcePreference(source.prefer);
  const fallbackToCdn = source.fallbackToCdn !== false;
  const items = {};

  for (const [id, raw] of Object.entries(isRecord(source.items) ? source.items : {})) {
    if (!isRecord(raw)) {
      warnings.push(`framework config "${id}" was ignored because it is not an object`);
      continue;
    }
    const probeGlobal = cleanString(raw.probeGlobal).replace(/^window\./, '');
    if (probeGlobal && !/^[$A-Z_a-z][$\w]*(?:\.[$A-Z_a-z][$\w]*)*$/.test(probeGlobal)) {
      warnings.push(`framework config "${id}" has an invalid probeGlobal path`);
    }
    items[id] = {
      localUrl: resolveUrl(raw.localUrl, configUrl),
      cdnUrl: resolveUrl(raw.cdnUrl, configUrl),
      prefer: sourcePreference(raw.prefer, prefer),
      fallbackToCdn: typeof raw.fallbackToCdn === 'boolean'
        ? raw.fallbackToCdn
        : fallbackToCdn,
      probeGlobal: /^[$A-Z_a-z][$\w]*(?:\.[$A-Z_a-z][$\w]*)*$/.test(probeGlobal)
        ? probeGlobal
        : '',
      intelligence: Array.isArray(raw.intelligence)
        ? [...new Set(raw.intelligence.map(cleanString).filter(Boolean))]
        : [],
    };
  }

  return { prefer, fallbackToCdn, items };
}

function normalizeAssetGroup(raw, configUrl, defaultPreference) {
  if (!isRecord(raw)) return null;
  const files = {};
  for (const [name, path] of Object.entries(isRecord(raw.files) ? raw.files : {})) {
    const clean = cleanString(path);
    if (clean) files[name] = clean;
  }
  const rawRuntime = isRecord(raw.runtime) ? raw.runtime : {};
  return {
    prefer: sourcePreference(raw.prefer, defaultPreference),
    localBaseUrl: resolveUrl(raw.localBaseUrl, configUrl, { folder: true }),
    hostedBaseUrl: resolveUrl(raw.hostedBaseUrl, configUrl, { folder: true }),
    intelligence: Array.isArray(raw.intelligence)
      ? [...new Set(raw.intelligence.map(cleanString).filter(Boolean))]
      : [],
    files,
    runtime: {
      enabled: rawRuntime.enabled === true,
      cssFiles: Array.isArray(rawRuntime.cssFiles)
        ? [...new Set(rawRuntime.cssFiles.map(cleanString).filter(Boolean))]
        : [],
      fluentIconElement: rawRuntime.fluentIconElement === true,
    },
  };
}

function normalizeConfig(raw, configUrl) {
  const warnings = [];
  if (!isRecord(raw)) {
    return { config: EMPTY_CONFIG, warnings: ['configuration root must be an object'] };
  }
  if (raw.version !== 1) {
    warnings.push(`configuration version ${JSON.stringify(raw.version)} is not supported; expected 1`);
  }

  const assets = {};
  for (const [name, value] of Object.entries(isRecord(raw.assets) ? raw.assets : {})) {
    const group = normalizeAssetGroup(value, configUrl, 'local');
    if (group) assets[name] = group;
  }

  return {
    config: {
      version: 1,
      frameworks: normalizeFrameworks(raw.frameworks, configUrl, warnings),
      assets,
    },
    warnings,
  };
}

function configUrl() {
  return window.__DCSPAD_CONFIG_URL__
    || new URL('../dcspad.config.json', import.meta.url).href;
}

export async function loadAppConfig() {
  const url = configUrl();
  try {
    const response = await fetch(url, {
      credentials: 'same-origin',
      cache: 'no-cache',
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const normalized = normalizeConfig(await response.json(), url);
    activeConfig = normalized.config;
    return { ...normalized, url };
  } catch (error) {
    activeConfig = EMPTY_CONFIG;
    return {
      config: activeConfig,
      url,
      warnings: [`dcspad.config.json could not be loaded (${error.message || error}); built-in framework URLs remain active`],
    };
  }
}

export function getAppConfig() {
  return activeConfig;
}

// Apply a config item without mutating the stored framework catalog. A blank
// preferred URL naturally falls through to the populated source.
export function applyFrameworkConfig(entry, config = activeConfig) {
  const override = config?.frameworks?.items?.[entry?.id];
  if (!override) return { ...entry };

  const local = override.localUrl;
  const cdn = override.cdnUrl;
  const preferred = override.prefer === 'cdn' ? cdn : local;
  const alternate = override.prefer === 'cdn' ? local : cdn;
  const primary = preferred || alternate;
  const fallback = preferred && alternate && override.fallbackToCdn
    ? alternate
    : '';

  const effective = {
    ...entry,
    intelligence: [
      ...new Set([
        ...(Array.isArray(entry.intelligence) ? entry.intelligence : []),
        ...override.intelligence,
      ]),
    ],
    configuredSources: { local, cdn },
    probeGlobal: override.probeGlobal,
  };

  if (!primary) return effective;
  const primaryIsCss = /\.css(?:[?#]|$)/i.test(primary);
  if (primaryIsCss) {
    effective.css = primary;
    delete effective.js;
    // Automatic fallback is intentionally limited to scripts because a
    // parser-blocking global probe can preserve dependency order for JS.
    return effective;
  }

  effective.js = primary;
  delete effective.css;
  if (fallback && !/\.css(?:[?#]|$)/i.test(fallback) && override.probeGlobal) {
    effective.fallbackJs = fallback;
  }
  return effective;
}

export function selectedAssetBase(group) {
  if (!group) return '';
  const preferred = group.prefer === 'hosted'
    ? group.hostedBaseUrl
    : group.localBaseUrl;
  return preferred || group.hostedBaseUrl || group.localBaseUrl || '';
}
