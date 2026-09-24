// SP Workbench page-copy — thin client over the sitepages/pages REST surface.
//
// This module is deliberately dumb: one request per method, exactly the
// shapes design/PAGE-COPY.md §0 extracted from the vendored PnPjs 2.15
// source (pnp-pages.js / pnp-files.js). No planning, no retries of its own
// — sp-write.js's post() already retries 403 (forced digest) and 429/503
// once, and mock mode is whatever `write`/`client` were built with. The
// orchestrator (page-copy.js / page-copy-run.js) owns every decision about
// *what* to call and *when*; this module only knows *how*.
//
// `client` is an sp-rest client (sp-rest.js) already connected to the web
// this instance talks to; `write` is the matching
// createSpWriteClient({ client, mockWriter }) for that same client — see
// list-schema-dialog.js's connect() for why the pair is always built
// together and never shared across a reconnect.

import { odataPathLiteral, SpFileError } from '../sp-odata.js';
import { readFileBytes as readFileBytesFromSpFiles } from '../sp-files.js';

export { readFileBytes } from '../sp-files.js';

export const VERBOSE = 'application/json;odata=verbose';
export const SITE_PAGE_META = Object.freeze({ __metadata: { type: 'SP.Publishing.SitePage' } });

// Mirrors sp-files.js readFileBytes's own default — kept local so this
// module never has to reach into sp-files.js internals to know the cap.
const MAX_FILE_BYTES = 50 * 1024 * 1024;

const FILE_INFO_SELECT = [
  'ListId', 'WebId', 'UniqueId', 'SiteId', 'Name', 'ServerRelativeUrl', 'Length', 'CheckOutType',
];

// AddImageFromExternalUrl's query values are OData string-literal query
// parameters, not path segments: encode first, then double any apostrophe
// the encoding left behind, then wrap in the literal's own quotes.
function queryLiteral(value) {
  return encodeURIComponent(String(value ?? '')).replaceAll("'", "''");
}

// GET helpers here (fileInfo/fileInfoById/fileItemId/featureActive) treat a
// 404 as "doesn't exist", not a failure — a copy plan probes for collisions
// and prior state, and "not found" is itself the answer. Anything else
// (permission, network, …) still rethrows.
function isNotFound(err) {
  return err?.status === 404 || err?.code === 'not-found';
}

async function catchNotFound(promise) {
  try {
    return await promise;
  } catch (err) {
    if (isNotFound(err)) return null;
    throw err;
  }
}

export function createSpPages({ client, write }) {
  async function getPage(id) {
    return client.get(`sitepages/pages(${Number(id)})`);
  }

  async function createPage({ pageLayoutType = 'Article', promotedState = 0 } = {}) {
    return write.postJson('sitepages/pages', {
      ...SITE_PAGE_META,
      PageLayoutType: pageLayoutType,
      PromotedState: promotedState,
    }, { contentType: VERBOSE, fallback: 'Could not create the page', code: 'page-create' });
  }

  async function addTemplateFile(folderServerRelativeUrl, fileServerRelativeUrl) {
    const path = `web/GetFolderByServerRelativePath(decodedUrl='${odataPathLiteral(folderServerRelativeUrl)}')`
      + `/Files/addTemplateFile(urloffile='${odataPathLiteral(fileServerRelativeUrl)}',templatefiletype=3)`;
    return write.postJson(path, {}, {
      fallback: 'Could not create the page file', code: 'page-create',
    });
  }

  async function checkoutPage(id) {
    return write.postJson(`sitepages/pages(${Number(id)})/checkoutpage`, {}, {
      fallback: 'Could not check out the page', code: 'page-checkout',
    });
  }

  // fields carries only the keys the caller supplies — savepage's body is
  // never assumed to be the full DTO, just SITE_PAGE_META plus whatever the
  // orchestrator decided to change.
  async function savePage(id, fields = {}) {
    return write.postJson(`sitepages/pages(${Number(id)})/savepage`, {
      ...SITE_PAGE_META,
      ...fields,
    }, {
      contentType: VERBOSE,
      headers: { 'If-Match': '*' },
      fallback: 'Could not save the page',
      code: 'page-save',
    });
  }

  async function publishPage(id) {
    return write.postJson(`sitepages/pages(${Number(id)})/publish`, {}, {
      fallback: 'Could not publish the page', code: 'page-publish',
    });
  }

  async function promoteToNews(id) {
    return write.postJson(`sitepages/pages(${Number(id)})/promoteToNews`, {}, {
      fallback: 'Could not promote the page to news', code: 'page-publish',
    });
  }

  // Undoes a checkout — never deletes a page. page-copy-run.js's
  // discardCopy() uses recycleFile() for that.
  async function discardPage(id) {
    return write.postJson(`sitepages/pages(${Number(id)})/discardPage`, SITE_PAGE_META, {
      contentType: VERBOSE, fallback: 'Could not discard the page checkout', code: 'page-checkout',
    });
  }

  async function recycleFile(serverRelativeUrl) {
    const path = `web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(serverRelativeUrl)}')/recycle`;
    return write.postJson(path, {}, { fallback: 'Could not recycle the file', code: 'page-recycle' });
  }

  // A page's list item, by id. Item ids are never reused within a list, so
  // this names exactly the page a run created — wherever it has since been
  // renamed or moved — which a remembered path cannot. discardCopy() uses it.
  async function itemFileRef(listId, itemId) {
    const item = await catchNotFound(client.get(
      `web/lists(guid'${listId}')/items(${Number(itemId)})`, { select: ['FileRef'] },
    ));
    return item ? String(item.FileRef || '') : null;
  }

  async function recycleItem(listId, itemId) {
    const path = `web/lists(guid'${listId}')/items(${Number(itemId)})/recycle`;
    return write.postJson(path, {}, { fallback: 'Could not recycle the page', code: 'page-recycle' });
  }

  async function recycleFolder(serverRelativeUrl) {
    const path = `web/GetFolderByServerRelativePath(decodedUrl='${odataPathLiteral(serverRelativeUrl)}')/recycle`;
    return write.postJson(path, {}, { fallback: 'Could not recycle the folder', code: 'page-recycle' });
  }

  // Shared shape for MoveFileByPath/CopyFileByPath — the two differ only in
  // the SP.MoveCopyUtil method name and ResetAuthorAndCreatedOnCopy.
  function moveCopyBody(srcAbsoluteUrl, destAbsoluteUrl, keepBoth, resetAuthorAndCreated) {
    return {
      srcPath: { __metadata: { type: 'SP.ResourcePath' }, DecodedUrl: srcAbsoluteUrl },
      destPath: { __metadata: { type: 'SP.ResourcePath' }, DecodedUrl: destAbsoluteUrl },
      options: {
        __metadata: { type: 'SP.MoveCopyOptions' },
        KeepBoth: Boolean(keepBoth),
        ResetAuthorAndCreatedOnCopy: resetAuthorAndCreated,
        ShouldBypassSharedLocks: true,
      },
    };
  }

  // srcAbsoluteUrl/destAbsoluteUrl are ABSOLUTE urls — the caller resolves
  // them (a plan's `finalAbsolute`/`stagingAbsolute`), never a server-relative
  // path, per §0's PnPjs source (extractWebUrl + hostUrl reconstruction).
  async function moveFileByPath(srcAbsoluteUrl, destAbsoluteUrl, { overwrite = false, keepBoth = false } = {}) {
    const path = `SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=${overwrite ? 'true' : 'false'}`;
    return write.postJson(path, moveCopyBody(srcAbsoluteUrl, destAbsoluteUrl, keepBoth, false), {
      contentType: VERBOSE, fallback: 'Could not move the page', code: 'page-move',
    });
  }

  async function copyFileByPath(srcAbsoluteUrl, destAbsoluteUrl, { overwrite = false, keepBoth = false } = {}) {
    const path = `SP.MoveCopyUtil.CopyFileByPath(overwrite=@a1)?@a1=${overwrite ? 'true' : 'false'}`;
    return write.postJson(path, moveCopyBody(srcAbsoluteUrl, destAbsoluteUrl, keepBoth, true), {
      contentType: VERBOSE, fallback: 'Could not copy the page', code: 'page-copy',
    });
  }

  async function addImageFromExternalUrl({ imageFileName, pageName, externalUrl } = {}) {
    const path = 'sitepages/AddImageFromExternalUrl'
      + `?imageFileName='${queryLiteral(imageFileName)}'`
      + `&pageName='${queryLiteral(pageName)}'`
      + `&externalUrl='${queryLiteral(externalUrl)}'`
      + '&$select=ServerRelativeUrl';
    return write.postJson(path, {}, { fallback: 'Could not import the image', code: 'page-image' });
  }

  async function fileInfo(serverRelativeUrl) {
    return catchNotFound(client.get(
      `web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(serverRelativeUrl)}')`,
      { select: FILE_INFO_SELECT },
    ));
  }

  async function fileInfoById(uniqueId) {
    return catchNotFound(client.get(
      `web/GetFileById('${odataPathLiteral(uniqueId)}')`,
      { select: FILE_INFO_SELECT },
    ));
  }

  async function fileItemId(serverRelativeUrl) {
    const item = await catchNotFound(client.get(
      `web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(serverRelativeUrl)}')/ListItemAllFields`,
      { select: ['Id'] },
    ));
    return item ? Number(item.Id) : null;
  }

  async function exists(serverRelativeUrl) {
    return (await fileInfo(serverRelativeUrl)) !== null;
  }

  // Folder counterpart to exists() — GetFolderByServerRelativePath answers
  // even for a path that doesn't exist (unlike the file endpoint, which
  // 404s), carrying its own Exists flag, so that flag is the source of
  // truth; a 404/not-found from the request itself still means "no".
  async function folderExists(serverRelativeUrl) {
    const folder = await catchNotFound(client.get(
      `web/GetFolderByServerRelativePath(decodedUrl='${odataPathLiteral(serverRelativeUrl)}')`,
      { select: ['Exists'] },
    ));
    return Boolean(folder?.Exists);
  }

  // Live: sp-files.js's own reader (its digest cache and 50 MB default are
  // shared with everything else that reads document bytes). Mock: the mock
  // web answers with { mockBytes: <length>, contentType } rather than real
  // bytes — the length is checked BEFORE allocating the zeroed buffer so an
  // oversized fixture never has to allocate to be rejected, mirroring
  // sp-files.js's own before-the-read content-length check.
  async function readFileBytes(serverRelativeUrl) {
    if (client.context().live) {
      return readFileBytesFromSpFiles(serverRelativeUrl, { webUrl: client.webUrl() });
    }
    const path = `web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(serverRelativeUrl)}')/$value`;
    const { mockBytes, contentType } = await client.get(path);
    const length = Number(mockBytes) || 0;
    if (length > MAX_FILE_BYTES) {
      throw new SpFileError(
        'The selected SharePoint file is larger than the transfer limit.',
        { code: 'too-large' },
      );
    }
    return {
      bytes: new Uint8Array(length).buffer,
      length,
      contentType: contentType || '',
      serverRelativeUrl,
    };
  }

  async function clientSideWebParts() {
    const { items } = await client.getAll('web/GetClientSideWebParts');
    return items;
  }

  async function setCommentsDisabled(listId, itemId, value) {
    const path = `web/lists(guid'${listId}')/items(${Number(itemId)})/SetCommentsDisabled`;
    return write.postJson(path, { value: Boolean(value) }, {
      fallback: 'Could not change the page comments setting', code: 'page-comments',
    });
  }

  // True only when the feature response carries a truthy DefinitionId.
  // {"odata.null": true}, an empty answer, or a 404 all mean "not active".
  async function featureActive(featureId) {
    const feature = await catchNotFound(client.get(
      `web/features/getbyid('${odataPathLiteral(featureId)}')`,
      { select: ['DefinitionId'] },
    ));
    return Boolean(feature?.DefinitionId);
  }

  async function sitePagesLibrary() {
    const { items } = await client.getAll('web/lists', {
      select: ['Id', 'Title', 'BaseTemplate', 'Hidden', 'RootFolder/ServerRelativeUrl'],
      expand: 'RootFolder',
    });
    const list = items.find((item) => !item.Hidden && Number(item.BaseTemplate) === 119);
    if (!list) return null;
    return {
      id: list.Id,
      title: list.Title,
      rootPath: list.RootFolder?.ServerRelativeUrl || '',
    };
  }

  async function webIdentity() {
    const [web, site] = await Promise.all([
      client.get('web', { select: ['Id', 'Title', 'Url', 'ServerRelativeUrl'] }),
      client.get('site', { select: ['Id', 'Url', 'ServerRelativeUrl'] }),
    ]);
    return {
      webId: web.Id,
      siteId: site.Id,
      webUrl: web.Url,
      webServerRelativeUrl: web.ServerRelativeUrl,
      siteUrl: site.Url,
      siteServerRelativeUrl: site.ServerRelativeUrl,
      title: web.Title,
    };
  }

  return {
    getPage,
    createPage,
    addTemplateFile,
    checkoutPage,
    savePage,
    publishPage,
    promoteToNews,
    discardPage,
    recycleFile,
    recycleFolder,
    itemFileRef,
    recycleItem,
    moveFileByPath,
    copyFileByPath,
    addImageFromExternalUrl,
    fileInfo,
    fileInfoById,
    fileItemId,
    exists,
    folderExists,
    readFileBytes,
    clientSideWebParts,
    setCommentsDisabled,
    featureActive,
    sitePagesLibrary,
    webIdentity,
  };
}
