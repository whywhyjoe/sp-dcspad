// List schema — capture (source client) and target probing.
//
// Every read here asks for Accept: application/json;odata=nometadata (the
// sp-rest.js client default) and, per the port basis, stays UNPROJECTED
// wherever the property set is heterogeneous: a $select naming a property one
// tenant lacks 400s the whole call, and fields/list entities are exactly that
// heterogeneous (LookupList, Choices, CustomFormatter… only exist on some
// rows). The Raw tab already reads the list entity this way (views/lists.js).

import {
  normalizeList, normalizeField, normalizeView, normalizeContentType, buildSchemaDoc,
  parentContentTypeId, LOOKUP_TYPES, TAXONOMY_TYPES,
} from './list-schema.js';
import { APP_BUILD_INFO } from '../build-info.js';
import { isDeniedRead, isExpiredSession } from './denied.js';

const guidPath = (listId, sub = '') => `web/lists(guid'${listId}')${sub}`;
const cleanGuid = (v) => String(v || '').replace(/[{}]/g, '').toLowerCase();

// Resolve a set of lookup-list GUIDs to their titles, one read each (cached
// by the caller across a single capture — this function itself does no
// caching beyond the one pass over `guids`). A GUID that fails to resolve
// (a hidden/inaccessible list) maps to null; the caller turns that into a
// warning, not a thrown error — a schema with an unresolvable lookup target
// is still worth exporting.
export async function resolveLookupTitles(client, guids, selfListId) {
  const map = new Map();
  const selfId = selfListId ? cleanGuid(selfListId) : null;
  for (const raw of guids) {
    const id = cleanGuid(raw);
    if (!id || map.has(id)) continue;
    if (id === selfId) { map.set(id, null); continue; }   // self-lookup — resolved by the caller
    try {
      const target = await client.get(guidPath(id), { select: 'Title' });
      map.set(id, target?.Title ?? null);
    } catch {
      map.set(id, null);
    }
  }
  return map;
}

export async function captureListSchema(client, listId, { includeHidden = false } = {}) {
  const warnings = [];

  const web = await client.get('web', { select: ['Id', 'Title', 'Url', 'ServerRelativeUrl', 'Language'] });
  // No $select: SP.List's own scalar set (ValidationFormula, versioning,
  // OnQuickLaunch, ListExperienceOptions…) is exactly the heterogeneous case
  // this capture is unprojected for.
  const list = await client.get(guidPath(listId), { expand: 'RootFolder' });
  // SharePoint leaves a few SP.List properties out of the default payload
  // (verified live): they only come back when named. One tolerant read for
  // the set; if the tenant rejects the set, try each alone and keep what
  // answers. A property that never answers stays at its doc default.
  const NAMED_ONLY = ['ValidationFormula', 'ValidationMessage', 'OnQuickLaunch', 'ReadSecurity', 'WriteSecurity'];
  const missing = NAMED_ONLY.filter((k) => !(k in list));
  if (missing.length) {
    try {
      Object.assign(list, await client.get(guidPath(listId), { select: missing }));
    } catch (err) {
      if (isDeniedRead(err) || isExpiredSession(err)) throw err;
      for (const key of missing) {
        try { Object.assign(list, await client.get(guidPath(listId), { select: [key] })); }
        catch { /* not on this tenant — keep the default */ }
      }
    }
  }

  const { items: rawFields } = await client.getAll(
    guidPath(listId, '/fields'),
    includeHidden ? {} : { filter: 'Hidden eq false' },
  );

  const lookupGuids = [...new Set(
    rawFields.filter((f) => LOOKUP_TYPES.has(f.TypeAsString)).map((f) => f.LookupList).filter(Boolean),
  )];
  const lookupTitles = await resolveLookupTitles(client, lookupGuids, list.Id);
  const listIdClean = cleanGuid(list.Id);
  const unreadable = new Set();   // one warning per unreadable target, not per column
  const lookupTitleById = (guid) => {
    const id = cleanGuid(guid);
    if (id === listIdClean) return list.Title;
    const title = lookupTitles.get(id);
    if (title == null && lookupTitles.has(id) && !unreadable.has(id)) {
      unreadable.add(id);
      warnings.push(`Lookup target list ${guid} could not be read; it will need a lookupMap entry.`);
    }
    return title ?? null;
  };

  const fields = rawFields.map((raw) => normalizeField(raw, { listId: list.Id, lookupTitleById }));
  for (const f of fields) {
    if (f.custom && TAXONOMY_TYPES.has(f.type)) {
      warnings.push(`Managed metadata column “${f.internalName}” cannot be recreated automatically (needs a term-set binding).`);
    }
  }

  // The ViewFields expansion saves a read per view, but not every tenant
  // honours it on the collection; when the expanded read is rejected, read
  // the views plain and let the per-view fallback below fetch each column set.
  // A denied read is the list's permission fact, not a shape problem — it
  // fails the capture like the fields read does.
  let rawViews;
  try {
    ({ items: rawViews } = await client.getAll(
      guidPath(listId, '/views'),
      { filter: 'PersonalView eq false', expand: 'ViewFields' },
    ));
  } catch (err) {
    if (isDeniedRead(err) || isExpiredSession(err)) throw err;
    ({ items: rawViews } = await client.getAll(
      guidPath(listId, '/views'),
      { filter: 'PersonalView eq false' },
    ));
  }
  const views = [];
  for (const raw of rawViews) {
    let names = raw.ViewFields?.Items?.results ?? raw.ViewFields?.Items ?? null;
    if (!Array.isArray(names)) {
      try {
        const vf = await client.get(guidPath(listId, `/views(guid'${raw.Id}')/viewfields`));
        names = vf?.Items?.results ?? vf?.Items ?? [];
      } catch (err) {
        warnings.push(`View “${raw.Title}”: fields could not be read (${err?.message || err}).`);
        names = [];
      }
    }
    views.push(normalizeView(raw, names));
  }

  let contentTypes = [];
  try {
    const { items: rawCts } = await client.getAll(guidPath(listId, '/contenttypes'), { expand: 'FieldLinks' });
    contentTypes = rawCts.map((raw) => {
      const links = raw.FieldLinks?.Items?.results ?? raw.FieldLinks?.Items ?? [];
      const fieldLinks = links.map((l) => ({ id: l.Id, name: l.Name, required: !!l.Required, hidden: !!l.Hidden }));
      return normalizeContentType(raw, fieldLinks);
    });
  } catch (err) {
    warnings.push(`Content types could not be read (${err?.message || err}).`);
  }

  const source = {
    siteUrl: client.webUrl(),
    listTitle: list.Title,
    listId: listIdClean,
    rootFolder: list.RootFolder?.ServerRelativeUrl ?? null,
    itemCount: list.ItemCount ?? null,
    webId: web.Id,
    webTitle: web.Title,
    language: web.Language,
    baseType: list.BaseType,
    entityTypeName: list.EntityTypeName,
    rootFolderName: list.RootFolder?.Name ?? null,
  };

  // Non-hidden content types, in the collection's own order — the order a
  // fresh list should attach and present them in.
  const normalizedList = normalizeList(list);
  normalizedList.contentTypeOrder = contentTypes.filter((ct) => !ct.hidden).map((ct) => ct.id);

  const doc = buildSchemaDoc({
    source,
    list: normalizedList,
    fields, views, contentTypes, warnings,
    generatorBuild: APP_BUILD_INFO.build,
  });

  return { doc, raw: { web, list, fields: rawFields, views: rawViews } };
}

// Read-side snapshot of the target web a Plan is built against: the list
// (if a same-title one already exists), its fields/views/content-type ids
// when it does, the target's available content types (only fetched when the
// source has content types enabled — an unnecessary $top=5000 read on every
// other list), and every list title on the target (for defaultTargetTitle
// and lookup-target resolution).
export async function probeTarget(client, { title, doc } = {}) {
  // One read of every list serves both the title collision and lookup-target
  // resolution. The match is made here, case-insensitively, rather than by
  // $filter: SharePoint titles collide case-insensitively, and the mock
  // resolver (like every fixture) ignores $filter.
  const { items: allLists } = await client.getAll('web/lists', {
    select: ['Id', 'Title', 'BaseTemplate', 'ContentTypesEnabled', 'RootFolder/ServerRelativeUrl'],
    expand: 'RootFolder',
  });
  const wanted = String(title || '').trim().toLowerCase();
  const match = wanted ? allLists.find((l) => String(l.Title || '').toLowerCase() === wanted) : null;
  const existingList = match
    ? {
      id: match.Id,
      title: match.Title,
      baseTemplate: match.BaseTemplate,
      contentTypesEnabled: !!match.ContentTypesEnabled,
      rootFolderUrl: match.RootFolder?.ServerRelativeUrl || '',
    }
    : null;

  let existingFields = [];
  let existingViews = [];
  let existingContentTypeIds = [];
  if (existingList) {
    const { items: rawFields } = await client.getAll(guidPath(existingList.id, '/fields'));
    existingFields = rawFields.map((f) => ({
      internalName: f.InternalName, title: f.Title, typeAsString: f.TypeAsString, id: f.Id,
    }));
    const { items: rawViews } = await client.getAll(guidPath(existingList.id, '/views'), { select: ['Id', 'Title', 'DefaultView'] });
    existingViews = rawViews.map((v) => ({ id: v.Id, title: v.Title, defaultView: !!v.DefaultView }));
    if (existingList.contentTypesEnabled) {
      const { items: rawCts } = await client.getAll(guidPath(existingList.id, '/contenttypes'), { select: ['StringId'] });
      existingContentTypeIds = rawCts.map((c) => parentContentTypeId(c.StringId));
    }
  }

  let availableContentTypes = null;
  if (doc?.list?.contentTypesEnabled) {
    const { items } = await client.getAll('web/availablecontenttypes', {
      select: ['StringId', 'Name', 'Group'], top: 5000,
    });
    availableContentTypes = items.map((c) => ({ id: c.StringId, name: c.Name, group: c.Group }));
  }

  return {
    existingList, existingFields, existingViews, existingContentTypeIds, availableContentTypes,
    targetLists: allLists.map((l) => ({ id: l.Id, title: l.Title })),
  };
}
