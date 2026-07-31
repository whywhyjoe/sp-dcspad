// CanvasContent1 parser for modern SharePoint pages — tolerant by design.
//
// Two storage formats exist in the wild:
//   1. A JSON array string (current SPO): each entry is a control record.
//   2. Legacy HTML markup: <div data-sp-canvascontrol data-sp-controldata="…">
//      where the attribute holds HTML-entity-encoded JSON (getAttribute
//      decodes entities for us) and web-part payloads sit in a nested
//      [data-sp-webpartdata] attribute; rich text sits in [data-sp-rte].
//
// Expected control shapes (documented, not enforced):
//   controlType 3  client-side web part: { webPartId, webPartData: { title,
//                  description, properties, serverProcessedContent:
//                  { htmlStrings, searchablePlainTexts, imageSources, links } } }
//   controlType 4  rich text: { innerHTML }
//   pageSettingsSlice   page settings record (no position)
//   position, no controlType   empty section/column marker
//   position: { zoneIndex, sectionIndex, controlIndex, sectionFactor
//               (column width /12), layoutIndex (2 = vertical section),
//               zoneId }, emphasis: { zoneEmphasis 0–3 }, zoneGroupMetadata
//               (collapsible sections)
//
// NOTHING here throws on malformed input: unparseable entries become
// { kind: 'unknown', raw } and push a message onto errors[]. The raw entry is
// always preserved so the Raw tab can show what the parser saw.

// Friendly names for first-party web part ids (lower-cased GUID keys).
// Wrong or missing entries only mislabel — unknown ids render as the raw
// GUID, which is the designed degrade path. Entries below the marker are
// lower-confidence; verify on a live tenant when convenient.
export const WEBPART_NAMES = {
  'd1d91016-032f-456d-98a4-721247c305e8': 'Image',
  'daf0b71c-6de8-4ef7-b511-faae7c388708': 'Highlighted content',
  '490d7c76-1824-45b2-9de3-676421c997fa': 'Embed',
  'b7dd04e1-19ce-4b24-9132-b60a1c2b910d': 'File viewer',
  'af8be689-990e-492a-81f7-ba3e4cd3ed9c': 'Image gallery',
  '6410b3b6-d440-4663-8744-378976dc041e': 'Link',
  '0ef418ba-5d19-4ade-9db0-b339873291d0': 'News feed',
  'a5df8fdf-b508-4b66-98a6-d83bc2597f63': 'News',
  '8c88f208-6c77-4bdb-86a0-0c47b4316588': 'News reel',
  '58fcd18b-e1af-4b0a-b23b-422c2c52d5a2': 'Power BI',
  '91a50c94-865f-4f5c-8b4e-e49659e69772': 'Quick chart',
  'eb95c819-ab8f-4689-bd03-0c2d65d47b1f': 'Site activity',
  '275c0095-a77e-4f6d-a2a0-6a7626911518': 'Stream',
  '31e9537e-f9dc-40a4-8834-0e3b7df418bc': 'Yammer embed',
  '20745d7d-8581-4a6c-bf26-68279bc123fc': 'Events',
  '6676088b-e28e-4a90-b9cb-d0d0303cd2eb': 'Group calendar',
  'c4bd7b2f-7b6e-4599-8485-16504575f590': 'Hero',
  'f92bf067-bc19-489e-a556-7fe95f508720': 'List',
  'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788': 'Page title',
  '7f718435-ee4d-431c-bdbf-9c4ff326f46e': 'People',
  'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd': 'Quick links',
  'e377ea37-9047-43b9-8cdb-a761be2f8e09': 'Bing maps',
  '2161a1c6-db61-4731-b97c-3cdb303f7cbb': 'Divider',
  '8654b779-4886-46d4-8ffb-b5ed960ee986': 'Spacer',
  'b19b3b9e-8d13-4fec-a93c-401a091c0707': 'Microsoft Forms',
  // ---- verify on live tenant (lower confidence) ----
  'f6fdf4f8-4a24-437b-a127-32e66a5dd9b4': 'Twitter',
  '868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823': 'Weather',
  'cf91cf5d-ac23-4a7a-9dbc-cd9ea1a095eb': 'Saved for later',
  '7cba020c-5ccb-42e8-b6fc-75b3149aba7b': 'Document library',
  '0f087d7f-520e-42b7-89c0-496aaf979d58': 'Button',
  'df8e44e7-edd5-46d5-90da-aca1539313b8': 'Call to action',
  '62cac389-787f-495d-beca-e11786162ef4': 'Countdown timer',
  '9d7e898c-f1bb-473a-9ace-8b415036578b': 'Organization chart',
  '71c19a43-d08c-4178-8218-4df8554c0b0e': 'Country/region web part',
  'e84a8ca2-f63c-4fb9-bc0b-d8eef5ccb22b': 'Sites',
  '544dd15b-cf3c-441b-96da-004d5a8cea1d': 'YouTube',
  'a8cd4347-f996-48c1-bcfb-75373fed2a27': 'World clock',
  '46698648-fcd5-41fc-9526-c7f7b2ace919': 'Markdown',
  '1ef5ed11-ce7b-44be-bc5e-4abd55101d16': 'Code snippet',
};

export function webPartName(webPartId) {
  const key = String(webPartId || '').toLowerCase().replace(/[{}]/g, '');
  return WEBPART_NAMES[key] || String(webPartId || '');
}

const asObject = (v) => (v && typeof v === 'object' && !Array.isArray(v) ? v : {});

export function normalizeControl(entry) {
  if (entry === null || entry === undefined || typeof entry !== 'object' || Array.isArray(entry)) {
    return { kind: 'unknown', raw: entry };
  }
  const position = asObject(entry.position);
  const base = {
    id: entry.id || entry.controlId || '',
    controlType: entry.controlType,
    position: {
      zoneIndex: position.zoneIndex,
      sectionIndex: position.sectionIndex,
      controlIndex: position.controlIndex,
      sectionFactor: position.sectionFactor,
      layoutIndex: position.layoutIndex,
      zoneId: position.zoneId,
    },
    emphasis: asObject(entry.emphasis),
    zoneGroupMetadata: entry.zoneGroupMetadata || null,
    raw: entry,
  };
  if (entry.pageSettingsSlice) {
    return { ...base, kind: 'pageSettings', pageSettingsSlice: entry.pageSettingsSlice };
  }
  if (entry.controlType === 4) {
    return { ...base, kind: 'text', innerHTML: String(entry.innerHTML ?? '') };
  }
  if (entry.controlType === 3) {
    const webPartData = asObject(entry.webPartData);
    const spc = asObject(webPartData.serverProcessedContent);
    return {
      ...base,
      kind: 'webpart',
      webPartId: String(entry.webPartId || webPartData.id || ''),
      webPartData: {
        title: String(webPartData.title ?? ''),
        description: String(webPartData.description ?? ''),
        properties: asObject(webPartData.properties),
        serverProcessedContent: {
          htmlStrings: asObject(spc.htmlStrings),
          searchablePlainTexts: asObject(spc.searchablePlainTexts),
          imageSources: asObject(spc.imageSources),
          links: asObject(spc.links),
        },
      },
    };
  }
  if (entry.position && entry.controlType === undefined) {
    return { ...base, kind: 'section' };
  }
  return { ...base, kind: 'unknown' };
}

function parseHtmlFormat(raw, errors) {
  const doc = new DOMParser().parseFromString(raw, 'text/html');
  const nodes = doc.querySelectorAll('[data-sp-canvascontrol], [data-sp-controldata]');
  const entries = [];
  for (const node of nodes) {
    const data = node.getAttribute('data-sp-controldata');
    if (!data) continue;
    let entry;
    try {
      entry = JSON.parse(data);
    } catch {
      errors.push('An HTML canvas control carried unparseable control data.');
      entries.push({ __unparseable: true, raw: data });
      continue;
    }
    // Web-part payloads and rich text live in nested nodes in this format.
    if (!entry.webPartData) {
      const wpNode = node.querySelector('[data-sp-webpartdata]');
      if (wpNode) {
        try { entry.webPartData = JSON.parse(wpNode.getAttribute('data-sp-webpartdata')); }
        catch { errors.push('An HTML canvas web part carried unparseable web-part data.'); }
      }
    }
    if (entry.controlType === 4 && entry.innerHTML === undefined) {
      entry.innerHTML = node.querySelector('[data-sp-rte]')?.innerHTML ?? '';
    }
    entries.push(entry);
  }
  if (!entries.length) errors.push('No canvas controls were found in the HTML markup.');
  return entries;
}

// raw: the CanvasContent1 field value (string | null). Never throws.
export function parseCanvasContent(raw) {
  const errors = [];
  const text = String(raw ?? '').trim();
  if (!text) return { ok: true, controls: [], pageSettings: null, errors };

  let entries = null;
  if (text.startsWith('<')) {
    try { entries = parseHtmlFormat(text, errors); }
    catch { errors.push('The HTML canvas markup could not be parsed.'); entries = []; }
  } else {
    try {
      const parsed = JSON.parse(text);
      entries = Array.isArray(parsed) ? parsed : [parsed];
    } catch {
      errors.push('CanvasContent1 is neither valid JSON nor recognizable HTML markup.');
      return { ok: false, controls: [], pageSettings: null, errors };
    }
  }

  const controls = [];
  let pageSettings = null;
  for (const entry of entries) {
    let control;
    try {
      control = entry?.__unparseable
        ? { kind: 'unknown', raw: entry.raw }
        : normalizeControl(entry);
    } catch {
      control = { kind: 'unknown', raw: entry };
      errors.push('A canvas entry could not be normalized.');
    }
    if (control.kind === 'unknown' && !entry?.__unparseable) {
      errors.push('A canvas entry had an unrecognized shape — shown raw.');
    }
    if (control.kind === 'pageSettings' && !pageSettings) {
      pageSettings = control.pageSettingsSlice;
    }
    controls.push(control);
  }
  return { ok: true, controls, pageSettings, errors };
}

// Group controls into ordered sections -> columns -> controls. Controls with
// missing/odd positions land in a trailing "unplaced" bucket.
export function buildSectionTree(controls) {
  const sections = new Map();
  const unplaced = [];
  for (const control of controls || []) {
    if (control.kind === 'pageSettings') continue;
    const { zoneIndex, sectionIndex } = control.position || {};
    if (typeof zoneIndex !== 'number' || typeof sectionIndex !== 'number') {
      unplaced.push(control);
      continue;
    }
    if (!sections.has(zoneIndex)) {
      sections.set(zoneIndex, {
        zoneIndex,
        emphasis: 0,
        vertical: false,
        collapsible: null,
        columns: new Map(),
      });
    }
    const section = sections.get(zoneIndex);
    if (typeof control.emphasis?.zoneEmphasis === 'number') {
      section.emphasis = control.emphasis.zoneEmphasis;
    }
    if (control.position.layoutIndex === 2) section.vertical = true;
    if (control.zoneGroupMetadata) section.collapsible = control.zoneGroupMetadata;
    if (!section.columns.has(sectionIndex)) {
      section.columns.set(sectionIndex, {
        sectionIndex,
        sectionFactor: control.position.sectionFactor,
        controls: [],
      });
    }
    const column = section.columns.get(sectionIndex);
    if (typeof control.position.sectionFactor === 'number') {
      column.sectionFactor = control.position.sectionFactor;
    }
    if (control.kind !== 'section') column.controls.push(control);
  }

  const ordered = [...sections.values()]
    .sort((a, b) => a.zoneIndex - b.zoneIndex)
    .map((section) => ({
      ...section,
      columns: [...section.columns.values()]
        .sort((a, b) => a.sectionIndex - b.sectionIndex)
        .map((column) => ({
          ...column,
          controls: [...column.controls].sort(
            (a, b) => (a.position.controlIndex ?? 0) - (b.position.controlIndex ?? 0),
          ),
        })),
    }));
  return { sections: ordered, unplaced };
}

// Human text of a control: rich text stripped to plain text, web parts via
// their searchable plain texts, else ''.
export function textOfControl(control) {
  if (!control) return '';
  if (control.kind === 'text') {
    const doc = new DOMParser().parseFromString(control.innerHTML || '', 'text/html');
    return (doc.body.textContent || '').replace(/\s+/g, ' ').trim();
  }
  if (control.kind === 'webpart') {
    const texts = control.webPartData?.serverProcessedContent?.searchablePlainTexts || {};
    return Object.values(texts).filter((v) => typeof v === 'string').join(' · ');
  }
  return '';
}

// Rendered-view sanitizer: drop active content, keep formatting. Used for the
// Text tab preview only — never for anything that executes.
export function sanitizeHtml(html) {
  const doc = new DOMParser().parseFromString(String(html || ''), 'text/html');
  for (const node of doc.querySelectorAll('script, iframe, object, embed, form')) {
    node.remove();
  }
  for (const node of doc.body.querySelectorAll('*')) {
    for (const attr of [...node.attributes]) {
      const name = attr.name.toLowerCase();
      if (name.startsWith('on')) node.removeAttribute(attr.name);
      else if ((name === 'href' || name === 'src' || name === 'xlink:href')
        && /^\s*javascript:/i.test(attr.value)) {
        node.removeAttribute(attr.name);
      }
    }
  }
  return doc.body.innerHTML;
}
