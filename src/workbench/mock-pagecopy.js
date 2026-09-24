// Mock SharePoint webs for the page-copy feature test suite.
// Three isolated webs: source (/sites/pagesrc), same-SC subweb (/sites/pagesrc/team),
// and cross-SC destination (/sites/pagedst). Stateful mock that records writes to
// globalThis.__DCSPAD_WB_WRITES__ and rejects unrecognized endpoints for page-copy webs.

import { SpFileError } from '../sp-odata.js';

const WEB_URL = typeof location !== 'undefined' ? location.origin : 'http://mock';

// ---- Fixture GUIDs (from context pack §4) ----

export const PAGECOPY_IDS = {
  // Source web
  srcSite: 'd1510000-0000-4000-8000-000000000001',
  srcWeb: 'd1520000-0000-4000-8000-000000000001',
  srcSitePages: 'd1530000-0000-4000-8000-000000000001',
  srcSiteAssets: 'd1540000-0000-4000-8000-000000000001',
  srcDocuments: 'd1550000-0000-4000-8000-000000000001',
  srcEvents: 'd1560000-0000-4000-8000-000000000001',
  // Team subweb (same site collection)
  teamWeb: 'd1520000-0000-4000-8000-000000000002',
  teamSitePages: 'd1530000-0000-4000-8000-000000000002',
  teamSiteAssets: 'd1540000-0000-4000-8000-000000000002',
  // Destination web (different site collection)
  dstSite: 'd1510000-0000-4000-8000-000000000003',
  dstWeb: 'd1520000-0000-4000-8000-000000000003',
  dstSitePages: 'd1530000-0000-4000-8000-000000000003',
  dstSiteAssets: 'd1540000-0000-4000-8000-000000000003',
};

// ---- State (per web) ----

const webState = {};

function initWebState(webPath, siteId, webId, sitePageListId, siteAssetsListId, documentsListId, eventsListId) {
  webState[webPath] = {
    path: webPath,
    siteId,
    webId,
    lists: [
      { Id: sitePageListId, Title: 'Site Pages', BaseTemplate: 119, Hidden: false, RootFolder: { ServerRelativeUrl: `${webPath}/SitePages` } },
      { Id: siteAssetsListId, Title: 'Site Assets', BaseTemplate: 101, Hidden: false, RootFolder: { ServerRelativeUrl: `${webPath}/SiteAssets` } },
      ...(documentsListId ? [{ Id: documentsListId, Title: 'Documents', BaseTemplate: 101, Hidden: false, RootFolder: { ServerRelativeUrl: `${webPath}/Shared Documents` } }] : []),
      ...(eventsListId ? [{ Id: eventsListId, Title: 'Events', BaseTemplate: 106, Hidden: false, RootFolder: { ServerRelativeUrl: `${webPath}/Lists/Events` } }] : []),
    ],
    files: new Map(), // lowercase path -> { Name, ServerRelativeUrl, UniqueId, ListId, WebId, SiteId, Length, CheckOutType, contentType }
    folders: new Set(), // lowercase paths
    pages: new Map(), // id -> { dto, item }
    nextPageId: 100,
  };
}

// Initialize the three webs
initWebState(
  '/sites/pagesrc',
  PAGECOPY_IDS.srcSite,
  PAGECOPY_IDS.srcWeb,
  PAGECOPY_IDS.srcSitePages,
  PAGECOPY_IDS.srcSiteAssets,
  PAGECOPY_IDS.srcDocuments,
  PAGECOPY_IDS.srcEvents,
);
initWebState(
  '/sites/pagesrc/team',
  PAGECOPY_IDS.srcSite, // same site collection
  PAGECOPY_IDS.teamWeb,
  PAGECOPY_IDS.teamSitePages,
  PAGECOPY_IDS.teamSiteAssets,
  undefined,
  undefined,
);
initWebState(
  '/sites/pagedst',
  PAGECOPY_IDS.dstSite,
  PAGECOPY_IDS.dstWeb,
  PAGECOPY_IDS.dstSitePages,
  PAGECOPY_IDS.dstSiteAssets,
  undefined,
  undefined,
);

// ---- Reset and fixture population ----

export function resetPageCopyMock() {
  // Wipe all state
  for (const path of Object.keys(webState)) {
    webState[path].files.clear();
    webState[path].folders.clear();
    webState[path].pages.clear();
    webState[path].nextPageId = 100;
  }
  globalThis.__DCSPAD_WB_WRITES__ = [];

  // Populate source fixtures
  const src = webState['/sites/pagesrc'];
  const srcSitePages = src.lists.find(l => l.Id === PAGECOPY_IDS.srcSitePages);

  // Source files (fixture GUIDs d15f0000-0000-4000-8000-0000000000NN)
  const sourceFiles = [
    { name: 'banner.jpg', path: '/sites/pagesrc/SiteAssets/SitePages/Quarterly-Update/banner.jpg', id: 'd15f0000-0000-4000-8000-000000000001', length: 120000, listId: PAGECOPY_IDS.srcSiteAssets },
    { name: 'thumb.png', path: '/sites/pagesrc/SiteAssets/thumbs/thumb.png', id: 'd15f0000-0000-4000-8000-000000000002', length: 30000, listId: PAGECOPY_IDS.srcSiteAssets },
    { name: 'chart.png', path: '/sites/pagesrc/SiteAssets/img/chart.png', id: 'd15f0000-0000-4000-8000-000000000003', length: 45000, listId: PAGECOPY_IDS.srcSiteAssets },
    { name: 'hero1.jpg', path: '/sites/pagesrc/SiteAssets/hero/hero1.jpg', id: 'd15f0000-0000-4000-8000-000000000004', length: 80000, listId: PAGECOPY_IDS.srcSiteAssets },
    { name: 'logo.png', path: '/sites/pagesrc/SiteAssets/icons/logo.png', id: 'd15f0000-0000-4000-8000-000000000005', length: 5000, listId: PAGECOPY_IDS.srcSiteAssets },
    { name: 'logo-2.png', path: '/sites/pagesrc/SiteAssets/img/logo.png', id: 'd15f0000-0000-4000-8000-000000000006', length: 6000, listId: PAGECOPY_IDS.srcSiteAssets },
    { name: 'huge.mp4', path: '/sites/pagesrc/SiteAssets/big/huge.mp4', id: 'd15f0000-0000-4000-8000-000000000007', length: 60 * 1024 * 1024, listId: PAGECOPY_IDS.srcSiteAssets },
    { name: 'report.pdf', path: '/sites/pagesrc/Shared Documents/report.pdf', id: 'd15f0000-0000-4000-8000-000000000008', length: 200000, listId: PAGECOPY_IDS.srcDocuments },
  ];
  for (const f of sourceFiles) {
    src.files.set(f.path.toLowerCase(), {
      Name: f.name,
      ServerRelativeUrl: f.path,
      UniqueId: f.id,
      ListId: f.listId,
      WebId: PAGECOPY_IDS.srcWeb,
      SiteId: PAGECOPY_IDS.srcSite,
      Length: f.length,
      CheckOutType: 2,
      contentType: 'application/octet-stream',
    });
  }

  // Folders for source files
  src.folders.add('/sites/pagesrc/siteassets/sitepages');
  src.folders.add('/sites/pagesrc/siteassets/sitepages/quarterly-update');
  src.folders.add('/sites/pagesrc/siteassets/thumbs');
  src.folders.add('/sites/pagesrc/siteassets/img');
  src.folders.add('/sites/pagesrc/siteassets/hero');
  src.folders.add('/sites/pagesrc/siteassets/icons');
  src.folders.add('/sites/pagesrc/siteassets/big');

  // Source pages
  const layoutJson = JSON.stringify([{
    id: 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788',
    instanceId: 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788',
    title: 'Title area',
    serverProcessedContent: {
      htmlStrings: {},
      searchablePlainTexts: {},
      imageSources: { imageSource: '/sites/pagesrc/SiteAssets/SitePages/Quarterly-Update/banner.jpg' },
      links: {},
      customMetadata: { imageSource: { siteId: PAGECOPY_IDS.srcSite, webId: PAGECOPY_IDS.srcWeb, listId: PAGECOPY_IDS.srcSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-000000000001' } },
    },
    dataVersion: '1.4',
    properties: {
      title: 'Quarterly Update',
      imageSourceType: 2,
      layoutType: 'FullWidthImage',
      textAlignment: 'Left',
      showTopicHeader: true,
      showPublishDate: true,
      topicHeader: 'Finance',
      authorByline: ['alex@contoso.com'],
      webId: PAGECOPY_IDS.srcWeb,
      siteId: PAGECOPY_IDS.srcSite,
      listId: PAGECOPY_IDS.srcSiteAssets,
      uniqueId: 'd15f0000-0000-4000-8000-000000000001',
    },
  }]);

  const canvasJson = JSON.stringify([
    { controlType: 4, id: 'd15e0000-0000-4000-8000-000000000001', position: { zoneIndex: 1, sectionIndex: 1, controlIndex: 1, sectionFactor: 12 }, innerHTML: '<p>See <a href="/sites/pagesrc/SitePages/Policies.aspx">policies</a>, <a href="/sites/pagesrc-other/SitePages/X.aspx">other</a> and <img src="/sites/pagesrc/SiteAssets/img/chart.png"></p>' },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-000000000002', webPartId: 'd1d91016-032f-456d-98a4-721247c305e8', webPartData: { id: 'd1d91016-032f-456d-98a4-721247c305e8', title: 'Image', properties: { imageSourceType: 2, siteId: PAGECOPY_IDS.srcSite, webId: PAGECOPY_IDS.srcWeb, listId: PAGECOPY_IDS.srcSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-000000000003' }, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: { imageSource: '/sites/pagesrc/SiteAssets/img/chart.png' }, links: {}, customMetadata: { imageSource: { siteId: PAGECOPY_IDS.srcSite, webId: PAGECOPY_IDS.srcWeb, listId: PAGECOPY_IDS.srcSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-000000000003' } } } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-000000000003', webPartId: 'c4bd7b2f-7b6e-4599-8485-16504575f590', webPartData: { id: 'c4bd7b2f-7b6e-4599-8485-16504575f590', title: 'Hero', properties: { content: [{ type: 'Image', image: { siteId: PAGECOPY_IDS.srcSite, webId: PAGECOPY_IDS.srcWeb, listId: PAGECOPY_IDS.srcSiteAssets, id: 'd15f0000-0000-4000-8000-000000000004' } }] }, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: { 'content[0].image.url': '/sites/pagesrc/SiteAssets/hero/hero1.jpg' }, links: { 'content[0].link': '/sites/pagesrc/SitePages/Policies.aspx' } } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-000000000004', webPartId: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd', webPartData: { id: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd', title: 'Quick links', properties: { items: [{ siteId: PAGECOPY_IDS.srcSite, webId: PAGECOPY_IDS.srcWeb, listId: PAGECOPY_IDS.srcSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-000000000005', thumbnailType: 3 }] }, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: { 'items[0].image.url': '/sites/pagesrc/SiteAssets/icons/logo.png' }, links: { 'items[0].sourceItem.url': '/sites/pagesrc/SitePages/Policies.aspx' } } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-000000000005', webPartId: 'df8e44e7-edd5-46d5-90da-aca1539313b8', webPartData: { id: 'df8e44e7-edd5-46d5-90da-aca1539313b8', title: 'Call to action', properties: {}, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: { imageSource: '/sites/pagesrc/SiteAssets/img/logo.png' }, links: { buttonLink: '/sites/pagesrc/SitePages/Join.aspx' }, customMetadata: { imageSource: { siteId: PAGECOPY_IDS.srcSite, webId: PAGECOPY_IDS.srcWeb, listId: PAGECOPY_IDS.srcSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-000000000006' } } } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-000000000006', webPartId: 'f92bf067-bc19-489e-a556-7fe95f508720', webPartData: { id: 'f92bf067-bc19-489e-a556-7fe95f508720', title: 'List', properties: { selectedListId: PAGECOPY_IDS.srcEvents, selectedListUrl: '/sites/pagesrc/Lists/Events', webRelativeListUrl: '/Lists/Events' }, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-000000000007', webPartId: '91a50c94-865f-4f5c-8b4e-e49659e69772', webPartData: { id: '91a50c94-865f-4f5c-8b4e-e49659e69772', title: 'Quick chart', properties: { dataProviderType: 'list', listId: PAGECOPY_IDS.srcEvents, siteUrl: '/sites/pagesrc' }, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-000000000008', webPartId: '7cba020c-5ccb-42e8-b6fc-75b3149aba7b', webPartData: { id: '7cba020c-5ccb-42e8-b6fc-75b3149aba7b', title: 'Sites', properties: { sites: [{ Url: '/sites/pagesrc/team', Title: 'Team' }] }, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-000000000009', webPartId: 'b7dd04e1-19ce-4b24-9132-b60a1c2b910d', webPartData: { id: 'b7dd04e1-19ce-4b24-9132-b60a1c2b910d', title: 'File viewer', properties: { file: '/sites/pagesrc/Shared Documents/report.pdf', uniqueId: 'd15f0000-0000-4000-8000-000000000008', siteId: PAGECOPY_IDS.srcSite, webId: PAGECOPY_IDS.srcWeb, listId: PAGECOPY_IDS.srcDocuments }, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} }, dynamicDataPaths: { fileUrl: 'WebPart.d15e0000-0000-4000-8000-00000000000a.d15e0000-0000-4000-8000-00000000000a:selectedDocument' } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-00000000000a', webPartId: 'f92bf067-bc19-489e-a556-7fe95f508720', webPartData: { id: 'f92bf067-bc19-489e-a556-7fe95f508720', title: 'Document library', properties: { isDocumentLibrary: true, selectedListId: PAGECOPY_IDS.srcDocuments, selectedListUrl: '/sites/pagesrc/Shared Documents' }, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-00000000000b', webPartId: 'd15c0000-0000-4000-8000-000000000001', webPartData: { id: 'd15c0000-0000-4000-8000-000000000001', title: 'Site-scoped SPFx', properties: {}, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-00000000000c', webPartId: 'd15c0000-0000-4000-8000-0000000000ff', webPartData: { id: 'd15c0000-0000-4000-8000-0000000000ff', title: 'Unknown part', properties: {}, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} } } },
    { controlType: 3, id: 'd15e0000-0000-4000-8000-00000000000d', webPartId: '275c0095-a77e-4f6d-a2a0-6a7626911518', webPartData: { id: '275c0095-a77e-4f6d-a2a0-6a7626911518', title: 'Video', properties: {}, serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: { videoSource: '/sites/pagesrc/SiteAssets/big/huge.mp4' } } } },
    { controlType: 0, pageSettingsSlice: { isDefaultDescription: false, isDefaultThumbnail: false } },
  ]);

  const quPage = {
    Id: 7,
    Title: 'Quarterly Update',
    FileName: 'Quarterly-Update.aspx',
    FileRef: '/sites/pagesrc/SitePages/Quarterly-Update.aspx',
    FileLeafRef: 'Quarterly-Update.aspx',
    FileDirRef: '/sites/pagesrc/SitePages',
    UniqueId: 'd15d0000-0000-4000-8000-000000000001',
    PageLayoutType: 'Article',
    PromotedState: 2,
    Description: 'Q3 numbers',
    TopicHeader: 'Finance',
    AuthorByline: ['alex@contoso.com'],
    CanvasContent1: canvasJson,
    LayoutWebpartsContent: layoutJson,
    // Live shape (spike §11): a custom-thumbnail page's DTO BannerImageUrl is
    // the THUMBNAIL's plain URL, BannerThumbnailUrl a tokened afdcache CDN URL,
    // and the banner itself only in the header part's imageSources.
    BannerImageUrl: `${WEB_URL}/sites/pagesrc/SiteAssets/thumbs/thumb.png`,
    BannerThumbnailUrl: `${WEB_URL}/_vti_bin/afdcache.ashx/authitem/sites/pagesrc/SiteAssets/thumbs/thumb.png?_oat_=mock&width=400`,
    CommentsDisabled: true,
    IsPageCheckedOutToCurrentUser: false,
    _UIVersionString: '2.0',
    PageCategory: 'IT',
    ReviewDate: '2026-10-01T07:00:00Z',
    ShowInNav: true,
    RelatedLink: { Url: '/sites/pagesrc/SitePages/Policies.aspx', Description: 'Policies' },
    Owner: { Id: 12, Title: 'alex@contoso.com' },
  };
  // Register source pages as files
  src.files.set('/sites/pagesrc/sitepages/quarterly-update.aspx', {
    Name: 'Quarterly-Update.aspx',
    ServerRelativeUrl: '/sites/pagesrc/SitePages/Quarterly-Update.aspx',
    UniqueId: 'd15d0000-0000-4000-8000-000000000001',
    ListId: PAGECOPY_IDS.srcSitePages,
    WebId: PAGECOPY_IDS.srcWeb,
    SiteId: PAGECOPY_IDS.srcSite,
    Length: 0,
    CheckOutType: 2,
    contentType: 'text/html',
  });

  src.pages.set(7, {
    dto: quPage,
    item: {
      Id: 7,
      Title: 'Quarterly Update',
      FileLeafRef: 'Quarterly-Update.aspx',
      FileRef: '/sites/pagesrc/SitePages/Quarterly-Update.aspx',
      FileDirRef: '/sites/pagesrc/SitePages',
      UniqueId: 'd15d0000-0000-4000-8000-000000000001',
      PromotedState: 2,
      _UIVersionString: '2.0',
      CanvasContent1: canvasJson,
      LayoutWebpartsContent: layoutJson,
      BannerImageUrl: { Url: quPage.BannerImageUrl, Description: quPage.BannerImageUrl },
      Description: 'Q3 numbers',
      CommentsDisabled: true,
      PageCategory: 'IT',
      ReviewDate: '2026-10-01T07:00:00Z',
      ShowInNav: true,
      RelatedLink: quPage.RelatedLink,
      Owner: { Id: 12 },
      CheckoutUser: null,
      HasUniqueRoleAssignments: false,
      Modified: '2026-09-20T10:00:00Z',
      Editor: { Title: 'alex@contoso.com' },
    },
  });

  src.pages.set(8, {
    dto: { Id: 8, Title: 'Legacy-News', FileName: 'Legacy-News.aspx', FileRef: '/sites/pagesrc/SitePages/Legacy-News.aspx', FileLeafRef: 'Legacy-News.aspx', FileDirRef: '/sites/pagesrc/SitePages', UniqueId: 'd15d0000-0000-4000-8000-000000000002', PageLayoutType: 'Article', PromotedState: 0, CanvasContent1: '<div><div data-sp-canvascontrol></div></div>' },
    item: { Id: 8, Title: 'Legacy-News', FileLeafRef: 'Legacy-News.aspx', FileRef: '/sites/pagesrc/SitePages/Legacy-News.aspx', FileDirRef: '/sites/pagesrc/SitePages', UniqueId: 'd15d0000-0000-4000-8000-000000000002', PromotedState: 0, _UIVersionString: '1.0' },
  });

  src.pages.set(9, {
    dto: { Id: 9, Title: 'Repost', FileName: 'Repost.aspx', FileRef: '/sites/pagesrc/SitePages/Repost.aspx', FileLeafRef: 'Repost.aspx', FileDirRef: '/sites/pagesrc/SitePages', UniqueId: 'd15d0000-0000-4000-8000-000000000003', PageLayoutType: 'RepostPage', CanvasContent1: '[]' },
    item: { Id: 9, Title: 'Repost', FileLeafRef: 'Repost.aspx', FileRef: '/sites/pagesrc/SitePages/Repost.aspx', UniqueId: 'd15d0000-0000-4000-8000-000000000003' },
  });

  // Pages 8 and 9 are files too — a whole-file copy (the legacy-HTML path)
  // finds its source through web.files.
  for (const [name, uniqueId] of [['Legacy-News.aspx', 'd15d0000-0000-4000-8000-000000000002'], ['Repost.aspx', 'd15d0000-0000-4000-8000-000000000003']]) {
    src.files.set(`/sites/pagesrc/sitepages/${name.toLowerCase()}`, {
      Name: name,
      ServerRelativeUrl: `/sites/pagesrc/SitePages/${name}`,
      UniqueId: uniqueId,
      ListId: PAGECOPY_IDS.srcSitePages,
      WebId: PAGECOPY_IDS.srcWeb,
      SiteId: PAGECOPY_IDS.srcSite,
      Length: 0,
      CheckOutType: 2,
      contentType: 'text/html',
    });
  }

  // Destination page (collision with source)
  const dst = webState['/sites/pagedst'];
  // Live SPO (spike §11): a site that ever had a page image carries
  // SiteAssets/SitePages already, and Folders/AddUsingPath on it FAILS.
  dst.folders.add('/sites/pagedst/siteassets/sitepages');
  dst.pages.set(3, {
    dto: { Id: 3, Title: 'Quarterly-Update', FileName: 'Quarterly-Update.aspx', FileRef: '/sites/pagedst/SitePages/Quarterly-Update.aspx', FileLeafRef: 'Quarterly-Update.aspx', FileDirRef: '/sites/pagedst/SitePages', UniqueId: 'd15d0000-0000-4000-8000-000000000010', PageLayoutType: 'Article', CanvasContent1: '[]' },
    item: { Id: 3, Title: 'Quarterly-Update', FileLeafRef: 'Quarterly-Update.aspx', FileRef: '/sites/pagedst/SitePages/Quarterly-Update.aspx', UniqueId: 'd15d0000-0000-4000-8000-000000000010' },
  });
  // Register destination page 3 as a file
  dst.files.set('/sites/pagedst/sitepages/quarterly-update.aspx', {
    Name: 'Quarterly-Update.aspx',
    ServerRelativeUrl: '/sites/pagedst/SitePages/Quarterly-Update.aspx',
    UniqueId: 'd15d0000-0000-4000-8000-000000000010',
    ListId: PAGECOPY_IDS.dstSitePages,
    WebId: PAGECOPY_IDS.dstWeb,
    SiteId: PAGECOPY_IDS.dstSite,
    Length: 0,
    CheckOutType: 2,
    contentType: 'text/html',
  });

  src.nextPageId = 10;
  dst.nextPageId = 4;
}

// Initialize on first load
resetPageCopyMock();

// ---- Helpers ----

function getWebState(webBase) {
  // webBase is like "http://localhost/sites/pagesrc" or "http://localhost/sites/pagesrc/team"
  // Extract just the path part
  let path = webBase;
  try {
    const url = new URL(webBase);
    path = url.pathname.replace(/\/+$/, '');
  } catch {
    // If URL parsing fails, treat as a path directly
    path = webBase.replace(/\/+$/, '');
  }

  const pathLower = path.toLowerCase();
  for (const [key, state] of Object.entries(webState)) {
    if (key.toLowerCase() === pathLower) return state;
  }
  return null;
}

function record(url, body, contentType, headers) {
  (globalThis.__DCSPAD_WB_WRITES__ ||= []).push({ url, body, contentType, headers });
}

function nextFileId() {
  let seq = (globalThis.__PAGECOPY_MOCK_FILE_SEQ__ = (globalThis.__PAGECOPY_MOCK_FILE_SEQ__ || 0) + 1);
  return `d15a0000-0000-4000-8000-${String(seq).padStart(12, '0')}`;
}

function sanitizeFileName(name) {
  return name.replaceAll(/["*:<>?/\\|#%\s]+/g, '-');
}

// ---- Resolver (reads) ----

export function pageCopyResolver(url, path, webBase) {
  const web = getWebState(webBase);
  if (!web) return undefined; // Not a page-copy web

  // Strip query string for routing, but keep original URL for fieldvaluesastext check
  const pathNoQuery = path.split('?')[0];
  const pathLower = pathNoQuery.toLowerCase();
  const origin = WEB_URL;

  // web, site endpoints
  if (pathLower === 'web' || pathLower.startsWith('web?')) {
    return {
      Id: web.webId,
      Title: 'Mock Web',
      Url: `${origin}${web.path}`,
      ServerRelativeUrl: web.path,
    };
  }
  if (pathLower === 'site' || pathLower.startsWith('site?')) {
    return {
      Id: web.siteId,
      Url: `${origin}${web.path.split('/').slice(0, -1).join('/')}`,
      ServerRelativeUrl: web.path.split('/').slice(0, -1).join('/'),
    };
  }

  // web/lists (collection; must not match web/lists(guid...))
  if (/^web\/lists($|\?)/.test(pathLower)) {
    return { value: web.lists.map(l => ({
      Id: l.Id,
      Title: l.Title,
      BaseTemplate: l.BaseTemplate,
      BaseType: l.BaseTemplate === 119 ? 1 : 1,
      Hidden: l.Hidden,
      ItemCount: l.BaseTemplate === 119 ? web.pages.size : web.files.size,
      DefaultViewUrl: `${web.path}/${l.Title.replaceAll(' ', '')}/AllItems.aspx`,
      RootFolder: l.RootFolder,
    })) };
  }

  // web/lists(guid'…')/items(N) — a single item, optionally with
  // FieldValuesAsText. Anchored (and checked before the collection/fields/
  // plain-list matches below) so it never falls through to `return list`.
  const itemMatch = /^web\/lists\(guid'([0-9a-f-]+)'\)\/items\((\d+)\)$/i.exec(pathLower);
  if (itemMatch) {
    const listId = itemMatch[1].toLowerCase();
    const itemId = Number(itemMatch[2]);
    const list = web.lists.find(l => l.Id.toLowerCase() === listId);
    if (!list) return null;

    const page = web.pages.get(itemId);
    if (page && url.toLowerCase().includes('fieldvaluesastext')) {
      // Return item with FieldValuesAsText expansion
      return {
        ...page.item,
        FieldValuesAsText: {
          Title: page.item.Title || '',
          FileLeafRef: page.item.FileLeafRef || '',
          Editor: page.item.Editor?.Title || '',
          // DateTime display text in the web's regional format — what the
          // carry set sends to ValidateUpdateListItem (ISO is refused live).
          ...(page.item.ReviewDate ? { ReviewDate: '10/1/2026 12:00 AM' } : {}),
        },
      };
    }
    return page?.item || null;
  }

  // web/lists(guid'…')/items — the whole collection (the Pages grid's read).
  // Only the Site Pages library is modeled with page rows; anything else on
  // these webs answers empty rather than falling through to `return list`.
  const itemsCollectionMatch = /^web\/lists\(guid'([0-9a-f-]+)'\)\/items$/i.exec(pathLower);
  if (itemsCollectionMatch) {
    const listId = itemsCollectionMatch[1].toLowerCase();
    const list = web.lists.find(l => l.Id.toLowerCase() === listId);
    if (!list) return null;
    if (list.BaseTemplate === 119) {
      return { value: [...web.pages.values()].map(p => p.item) };
    }
    return { value: [] };
  }

  // web/lists(guid'…')/fields
  const fieldsMatch = /^web\/lists\(guid'([0-9a-f-]+)'\)\/fields$/i.exec(pathLower);
  if (fieldsMatch) {
    const listId = fieldsMatch[1].toLowerCase();
    const list = web.lists.find(l => l.Id.toLowerCase() === listId);
    if (!list) return null;

    const fields = [
      { Id: 'f1', Title: 'Title', InternalName: 'Title', TypeAsString: 'Text', Required: true, Hidden: false, ReadOnlyField: false, Choices: undefined },
      { Id: 'f2', Title: 'Description', InternalName: 'Description', TypeAsString: 'Note', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
      { Id: 'f3', Title: 'Banner Image URL', InternalName: 'BannerImageUrl', TypeAsString: 'URL', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
      { Id: 'f4', Title: 'Promoted State', InternalName: 'PromotedState', TypeAsString: 'Number', Required: false, Hidden: false, ReadOnlyField: true, Choices: undefined },
      { Id: 'f5', Title: 'Canvas Content', InternalName: 'CanvasContent1', TypeAsString: 'Note', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
      { Id: 'f6', Title: 'Layout Web Parts Content', InternalName: 'LayoutWebpartsContent', TypeAsString: 'Note', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
      { Id: 'f7', Title: 'File Leaf Ref', InternalName: 'FileLeafRef', TypeAsString: 'File', Required: false, Hidden: false, ReadOnlyField: true, Choices: undefined },
      { Id: 'f8', Title: 'Version', InternalName: '_UIVersionString', TypeAsString: 'Text', Required: false, Hidden: false, ReadOnlyField: true, Choices: undefined },
      { Id: 'f9', Title: 'Comments Disabled', InternalName: 'CommentsDisabled', TypeAsString: 'Boolean', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
    ];
    if (list.Id === PAGECOPY_IDS.srcSitePages || list.Id === PAGECOPY_IDS.teamSitePages || list.Id === PAGECOPY_IDS.dstSitePages) {
      // Add custom columns for source and team; different for dst
      if (web.path === '/sites/pagedst') {
        fields.push(
          { Id: 'f10', Title: 'Page Category', InternalName: 'PageCategory', TypeAsString: 'Choice', Required: false, Hidden: false, ReadOnlyField: false, Choices: ['Finance', 'HR'] },
          { Id: 'f11', Title: 'Review Date', InternalName: 'ReviewDate', TypeAsString: 'DateTime', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
          { Id: 'f12', Title: 'Show In Nav', InternalName: 'ShowInNav', TypeAsString: 'Boolean', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
          { Id: 'f13', Title: 'Related Link', InternalName: 'RelatedLink', TypeAsString: 'URL', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
          { Id: 'f14', Title: 'Audience', InternalName: 'Audience', TypeAsString: 'Text', Required: true, Hidden: false, ReadOnlyField: false, Choices: undefined },
          { Id: 'f15', Title: 'Owner', InternalName: 'Owner', TypeAsString: 'User', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
        );
      } else {
        fields.push(
          { Id: 'f10', Title: 'Page Category', InternalName: 'PageCategory', TypeAsString: 'Choice', Required: false, Hidden: false, ReadOnlyField: false, Choices: ['Finance', 'HR', 'IT'] },
          { Id: 'f11', Title: 'Review Date', InternalName: 'ReviewDate', TypeAsString: 'DateTime', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
          { Id: 'f12', Title: 'Show In Nav', InternalName: 'ShowInNav', TypeAsString: 'Boolean', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
          { Id: 'f13', Title: 'Related Link', InternalName: 'RelatedLink', TypeAsString: 'URL', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
          { Id: 'f15', Title: 'Owner', InternalName: 'Owner', TypeAsString: 'User', Required: false, Hidden: false, ReadOnlyField: false, Choices: undefined },
        );
      }
    }
    return { value: fields };
  }

  // web/lists(guid'…') — the plain list entity, checked last so it never
  // shadows /items, /items(N) or /fields above.
  const listMatch = /^web\/lists\(guid'([0-9a-f-]+)'\)$/i.exec(pathLower);
  if (listMatch) {
    const listId = listMatch[1].toLowerCase();
    const list = web.lists.find(l => l.Id.toLowerCase() === listId);
    return list || null;
  }

  // sitepages/pages(N)
  const pageMatch = /sitepages\/pages\((\d+)\)/i.exec(pathLower);
  if (pageMatch) {
    const pageId = Number(pageMatch[1]);
    const page = web.pages.get(pageId);
    return page ? page.dto : null;
  }

  // web/getfilebyserverrelativepath(decodedurl='…')
  const fileByPathMatch = /web\/getfilebyserverrelativepath\(decodedurl='([^']*)'\)/i.exec(url.toLowerCase());
  if (fileByPathMatch) {
    let decodedPath = fileByPathMatch[1];
    // Decode the URL-encoded path
    try { decodedPath = decodeURIComponent(decodedPath); } catch { /* keep raw */ }
    // Compare case-insensitively
    const file = Array.from(web.files.entries()).find(([key]) => key.toLowerCase() === decodedPath.toLowerCase())?.[1];
    if (!file) return null;

    if (pathLower.includes('/listitemallfields')) {
      // The page whose file this is; any other file's item is out of scope.
      const page = [...web.pages.values()].find((pg) => String(pg.dto.FileRef || pg.item.FileRef || '').toLowerCase() === file.ServerRelativeUrl.toLowerCase());
      return page ? { ...page.item, Id: page.dto.Id } : null;
    }
    if (pathLower.includes('/$value')) {
      return { mockBytes: file.Length, contentType: file.contentType };
    }
    return {
      Name: file.Name,
      ServerRelativeUrl: file.ServerRelativeUrl,
      UniqueId: file.UniqueId,
      ListId: file.ListId,
      WebId: file.WebId,
      SiteId: file.SiteId,
      Length: file.Length,
      CheckOutType: file.CheckOutType,
    };
  }

  // web/getfolderbyserverrelativepath(decodedurl='…') — the folder probe
  // (sp-pages folderExists). Exists for a registered folder, a list root, or
  // any path that holds a file; SPO answers Exists:false rather than 404.
  const folderMatch = /^web\/getfolderbyserverrelativepath\(decodedurl='([^']*)'\)$/i.exec(pathLower);
  if (folderMatch) {
    let folderPath = folderMatch[1];
    try { folderPath = decodeURIComponent(folderPath).replaceAll("''", "'"); } catch { /* keep raw */ }
    const lowerPath = folderPath.toLowerCase().replace(/\/+$/, '');
    const exists = web.folders.has(lowerPath)
      || web.lists.some((l) => String(l.RootFolder?.ServerRelativeUrl || '').toLowerCase() === lowerPath)
      || [...web.files.keys()].some((k) => k.startsWith(`${lowerPath}/`));
    // SP.Folder.ItemCount: files and subfolders directly inside, not deeper.
    const directChild = (k) => k.startsWith(`${lowerPath}/`) && !k.slice(lowerPath.length + 1).includes('/');
    const itemCount = [...web.files.keys()].filter(directChild).length
      + [...web.folders].filter(directChild).length;
    return { Exists: exists, ServerRelativeUrl: folderPath, ItemCount: exists ? itemCount : 0 };
  }

  // web/getfilebyid('guid')
  const fileByIdMatch = /web\/getfilebyid\('([0-9a-f-]+)'\)/i.exec(pathLower);
  if (fileByIdMatch) {
    const uniqueId = fileByIdMatch[1].toLowerCase();
    for (const file of web.files.values()) {
      if (file.UniqueId.toLowerCase() === uniqueId) {
        return {
          Name: file.Name,
          ServerRelativeUrl: file.ServerRelativeUrl,
          UniqueId: file.UniqueId,
          ListId: file.ListId,
          WebId: file.WebId,
          SiteId: file.SiteId,
          Length: file.Length,
          CheckOutType: file.CheckOutType,
        };
      }
    }
    return null;
  }

  // web/getclientsidewebparts
  if (pathLower === 'web/getclientsidewebparts') {
    const ids = [
      { Id: 'd1d91016-032f-456d-98a4-721247c305e8', Name: 'Image' },
      { Id: 'c4bd7b2f-7b6e-4599-8485-16504575f590', Name: 'Hero' },
      { Id: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd', Name: 'Quick links' },
      { Id: 'df8e44e7-edd5-46d5-90da-aca1539313b8', Name: 'Call to action' },
      { Id: 'f92bf067-bc19-489e-a556-7fe95f508720', Name: 'List' },
      { Id: 'b7dd04e1-19ce-4b24-9132-b60a1c2b910d', Name: 'File viewer' },
      { Id: '7cba020c-5ccb-42e8-b6fc-75b3149aba7b', Name: 'Sites' },
      { Id: '275c0095-a77e-4f6d-a2a0-6a7626911518', Name: 'Video' },
    ];
    // Source includes site-scoped SPFx; pagedst lacks Quick chart; nobody lists the unknown one
    // (Ids from a live GetClientSideWebParts: 7cba020c is Sites.)
    if (web.path === '/sites/pagesrc') {
      ids.push({ Id: '91a50c94-865f-4f5c-8b4e-e49659e69772', Name: 'Quick chart' });
      ids.push({ Id: 'd15c0000-0000-4000-8000-000000000001', Name: 'Site-scoped SPFx' });
    } else if (web.path === '/sites/pagesrc/team') {
      ids.push({ Id: '91a50c94-865f-4f5c-8b4e-e49659e69772', Name: 'Quick chart' });
    }
    return { value: ids };
  }

  // web/features/getbyid('…')
  if (/web\/features\/getbyid\('([0-9a-f-]+)'\)/i.test(pathLower)) {
    const featureMatch = /web\/features\/getbyid\('([0-9a-f-]+)'\)/i.exec(pathLower);
    if (featureMatch) {
      const featureId = featureMatch[1].toLowerCase();
      if (featureId === 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec') {
        return { DefinitionId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec' };
      }
      return { 'odata.null': true };
    }
  }

  // Unrecognized read endpoint for a page-copy web → return null (404)
  return null;
}

// ---- Writer (writes) ----

export function pageCopyWriter(url, body, contentType, headers) {
  const webBase = url.slice(0, url.indexOf('/_api/')).replace(/\/+$/, '');
  const web = getWebState(webBase);
  if (!web) return undefined; // Not a page-copy web

  // Strip query string for routing
  const pathFull = url.slice(url.indexOf('/_api/') + 6);
  const pathNoQuery = pathFull.split('?')[0];
  const pathLower = pathNoQuery.toLowerCase();
  const record = () => { globalThis.__DCSPAD_WB_WRITES__ ||= []; globalThis.__DCSPAD_WB_WRITES__.push({ url, body, contentType, headers }); };

  // Check for failure injection. Every code other than 'network' throws
  // BEFORE the write's state change — dispatchWrite() below never runs, and
  // the attempted write is still recorded for the audit trail (matching
  // real writes, which are recorded whether or not the server accepts
  // them). 'network' means the request reached the server and the state
  // change happened, but the response was lost — SpFileError(code:
  // 'network') is the ONLY "response lost" signal (sp-write.js), so
  // dispatchWrite() below still runs to completion (it records and mutates
  // state exactly like an unfailed write) and only then do we throw instead
  // of returning its result.
  let networkFailure = null;
  if (globalThis.__PAGECOPY_MOCK_FAIL__) {
    for (let i = 0; i < globalThis.__PAGECOPY_MOCK_FAIL__.length; i++) {
      const fail = globalThis.__PAGECOPY_MOCK_FAIL__[i];
      // Case-insensitive: every URL match elsewhere in this mock works
      // against a lowercased path, and PnPjs's own casing (promoteToNews,
      // CheckIn, …) shouldn't be a trap for a caller's `match` string.
      if (url.toLowerCase().includes(String(fail.match).toLowerCase())) {
        const err = new SpFileError(fail.message || 'mock failure', { code: fail.code || 'write', status: fail.status || 500 });
        if (fail.code === 'network') {
          networkFailure = err;
        } else {
          record();
          if (fail.once) globalThis.__PAGECOPY_MOCK_FAIL__.splice(i, 1);
          throw err;
        }
        if (fail.once) globalThis.__PAGECOPY_MOCK_FAIL__.splice(i, 1);
        break;
      }
    }
  }

  let bodyData = {};
  try { bodyData = JSON.parse(body || '{}'); } catch { /* keep {} */ }

  const dispatched = dispatchWrite();
  if (networkFailure) throw networkFailure;
  return dispatched;

  // All recognized write endpoints for a page-copy web. Declared as a
  // function (hoisted, called above) rather than run inline, so the failure
  // injection above can decide — from the matched entry's code — whether
  // this runs at all before it does.
  function dispatchWrite() {

  // sitepages/pages POST (create new page)
  if (pathLower === 'sitepages/pages') {
    record();
    const pageId = web.nextPageId++;
    const dto = {
      Id: pageId,
      Title: bodyData.Title || '',
      FileName: `Untitled_${pageId}.aspx`,
      FileRef: `${web.path}/SitePages/Untitled_${pageId}.aspx`,
      FileLeafRef: `Untitled_${pageId}.aspx`,
      FileDirRef: `${web.path}/SitePages`,
      UniqueId: nextFileId(),
      CanvasContent1: '[]',
      LayoutWebpartsContent: '[]',
      PromotedState: bodyData.PromotedState ?? 0,
      PageLayoutType: bodyData.PageLayoutType || 'Article',
      IsPageCheckedOutToCurrentUser: true,
      _UIVersionString: '0.1',
    };
    web.pages.set(pageId, {
      dto,
      item: {
        Id: pageId,
        Title: '',
        FileLeafRef: `Untitled_${pageId}.aspx`,
        FileRef: `${web.path}/SitePages/Untitled_${pageId}.aspx`,
        UniqueId: dto.UniqueId,
        _UIVersionString: '0.1',
        CanvasContent1: '[]',
        LayoutWebpartsContent: '[]',
      },
    });
    return dto;
  }

  // sitepages/pages(N)/checkoutpage
  if (/sitepages\/pages\(\d+\)\/checkoutpage/i.test(pathLower)) {
    record();
    const match = /sitepages\/pages\((\d+)\)\/checkoutpage/i.exec(pathLower);
    if (match) {
      const pageId = Number(match[1]);
      const page = web.pages.get(pageId);
      if (page) {
        page.dto.IsPageCheckedOutToCurrentUser = true;
      }
    }
    return {};
  }

  // sitepages/pages(N)/savepage
  if (/sitepages\/pages\(\d+\)\/savepage/i.test(pathLower)) {
    record();
    const match = /sitepages\/pages\((\d+)\)\/savepage/i.exec(pathLower);
    if (match) {
      const pageId = Number(match[1]);
      const page = web.pages.get(pageId);
      if (page) {
        const dto = page.dto;
        const item = page.item;

        // Rename file if Title is first being set and FileName starts with "Untitled_"
        if (bodyData.Title && dto.FileName.startsWith('Untitled_')) {
          const stem = sanitizeFileName(bodyData.Title);
          let newName = `${stem}.aspx`;
          let counter = 1;
          // Check for collisions in this web's SitePages folder
          for (const existingPage of web.pages.values()) {
            if (existingPage.dto.FileLeafRef && existingPage.dto.FileLeafRef.toLowerCase() === newName.toLowerCase() && existingPage.dto.Id !== pageId) {
              newName = `${stem}${counter++}.aspx`;
            }
          }
          dto.FileName = newName;
          dto.FileLeafRef = newName;
          dto.FileRef = `${web.path}/SitePages/${newName}`;
          item.FileLeafRef = newName;
          item.FileRef = `${web.path}/SitePages/${newName}`;
        }

        // Register or update file entry
        const filePath = `${web.path}/SitePages/${dto.FileLeafRef}`.toLowerCase();
        const sitePagesList = web.lists.find(l => l.BaseTemplate === 119);
        web.files.set(filePath, {
          Name: dto.FileLeafRef,
          ServerRelativeUrl: `${web.path}/SitePages/${dto.FileLeafRef}`,
          UniqueId: dto.UniqueId,
          ListId: sitePagesList?.Id,
          WebId: web.webId,
          SiteId: web.siteId,
          Length: 0,
          CheckOutType: 2,
          contentType: 'text/html',
        });

        // Merge body fields
        if ('Title' in bodyData) { dto.Title = bodyData.Title; item.Title = bodyData.Title; }
        if ('Description' in bodyData) { dto.Description = bodyData.Description; item.Description = bodyData.Description; }
        if ('BannerImageUrl' in bodyData) { dto.BannerImageUrl = bodyData.BannerImageUrl; item.BannerImageUrl = bodyData.BannerImageUrl; }
        if ('CanvasContent1' in bodyData) { dto.CanvasContent1 = bodyData.CanvasContent1; item.CanvasContent1 = bodyData.CanvasContent1; }
        if ('LayoutWebpartsContent' in bodyData) { dto.LayoutWebpartsContent = bodyData.LayoutWebpartsContent; item.LayoutWebpartsContent = bodyData.LayoutWebpartsContent; }
        if ('TopicHeader' in bodyData) dto.TopicHeader = bodyData.TopicHeader;
        if ('AuthorByline' in bodyData) dto.AuthorByline = bodyData.AuthorByline;
      }
    }
    return {};
  }

  // sitepages/pages(N)/publish
  if (/sitepages\/pages\(\d+\)\/publish/i.test(pathLower)) {
    record();
    const match = /sitepages\/pages\((\d+)\)\/publish/i.exec(pathLower);
    if (match) {
      const pageId = Number(match[1]);
      const page = web.pages.get(pageId);
      if (page) {
        const label = page.dto._UIVersionString || '0.1';
        const [major, minor] = label.split('.').map(Number);
        const newMajor = (major || 0) + 1;
        page.dto._UIVersionString = `${newMajor}.0`;
        page.item._UIVersionString = `${newMajor}.0`;
        page.dto.IsPageCheckedOutToCurrentUser = false;
      }
    }
    return {};
  }

  // sitepages/pages(N)/promotetonews
  if (/sitepages\/pages\(\d+\)\/promotetonews/i.test(pathLower)) {
    record();
    const match = /sitepages\/pages\((\d+)\)\/promotetonews/i.exec(pathLower);
    if (match) {
      const pageId = Number(match[1]);
      const page = web.pages.get(pageId);
      if (page) {
        page.dto.PromotedState = 1;
      }
    }
    return {};
  }

  // sitepages/pages(N)/discardpage
  if (/sitepages\/pages\(\d+\)\/discardpage/i.test(pathLower)) {
    record();
    const match = /sitepages\/pages\((\d+)\)\/discardpage/i.exec(pathLower);
    if (match) {
      const pageId = Number(match[1]);
      const page = web.pages.get(pageId);
      if (page) {
        page.dto.IsPageCheckedOutToCurrentUser = false;
      }
    }
    return {};
  }

  // web/getfilebyserverrelativepath(decodedurl='…')/recycle
  if (/web\/getfilebyserverrelativepath\(decodedurl='([^']*)'\)\/recycle/i.test(pathLower)) {
    record();
    const match = /web\/getfilebyserverrelativepath\(decodedurl='([^']*)'\)\/recycle/i.exec(pathLower);
    if (match) {
      let decodedPath = match[1];
      try { decodedPath = decodeURIComponent(decodedPath); } catch { /* keep raw */ }
      const srLower = decodedPath.toLowerCase();

      web.files.delete(srLower);
      // Also remove any associated page
      for (const [id, page] of web.pages) {
        if (page.dto.FileRef && page.dto.FileRef.toLowerCase() === srLower) {
          web.pages.delete(id);
        }
      }
    }
    return {};
  }

  // web/getfolderbyserverrelativepath(decodedurl='…')/recycle
  if (/web\/getfolderbyserverrelativepath\(decodedurl='([^']*)'\)\/recycle/i.test(pathLower)) {
    record();
    const match = /web\/getfolderbyserverrelativepath\(decodedurl='([^']*)'\)\/recycle/i.exec(pathLower);
    if (match) {
      let decodedPath = match[1];
      try { decodedPath = decodeURIComponent(decodedPath); } catch { /* keep raw */ }
      web.folders.delete(decodedPath.toLowerCase());
    }
    return {};
  }

  // sp.movecopyutil.movefilebypath / copyfilebypath
  if (/sp\.movecopyutil\.(move|copy)filebypath/i.test(pathLower)) {
    record();
    const srcAbsUrl = bodyData.srcPath?.DecodedUrl;
    const dstAbsUrl = bodyData.destPath?.DecodedUrl;
    // The request line is `…MoveFileByPath(overwrite=@a1)?@a1=false` — the
    // literal substring "overwrite=false" never appears, so the @a1 query
    // parameter has to be read instead.
    const overwriteMatch = /[?&]@a1=(true|false)/i.exec(url);
    const overwrite = overwriteMatch ? overwriteMatch[1].toLowerCase() === 'true' : false;

    if (srcAbsUrl && dstAbsUrl) {
      // DecodedUrl values are absolute URLs; take the pathname and decode it
      // so a percent-encoded segment compares the same as a plain one.
      let srcPath = srcAbsUrl;
      let dstPath = dstAbsUrl;
      try {
        srcPath = decodeURIComponent(new URL(srcAbsUrl).pathname);
        dstPath = decodeURIComponent(new URL(dstAbsUrl).pathname);
      } catch { /* keep as-is */ }

      const srcFile = Array.from(web.files.entries()).find(([key]) => key.toLowerCase() === srcPath.toLowerCase())?.[1];
      if (!srcFile) throw new SpFileError('Source file not found', { code: 'not-found', status: 404 });

      const isMove = /movefilebypath/i.test(pathLower);
      const isCopy = /copyfilebypath/i.test(pathLower);

      const dstExists = Array.from(web.files.entries()).some(([key]) => key.toLowerCase() === dstPath.toLowerCase());
      if (dstExists && !overwrite) {
        throw new SpFileError('Destination exists', { code: 'conflict', status: 409 });
      }

      if (isCopy) {
        // Copy: create a new page on destination
        const dstWeb = getWebState(url.substring(0, url.indexOf('/_api/')).replace(/\/+$/, ''));
        if (dstWeb && srcFile.ServerRelativeUrl.includes('/SitePages/')) {
          // This is a page copy within the mock
          const sourcePage = [...web.pages.values()].find(p => p.dto.FileRef && p.dto.FileRef.toLowerCase() === srcFile.ServerRelativeUrl.toLowerCase());
          if (sourcePage) {
            const newPageId = dstWeb.nextPageId++;
            const newName = dstPath.split('/').pop();
            const newDir = dstPath.slice(0, dstPath.lastIndexOf('/'));
            const newDto = JSON.parse(JSON.stringify(sourcePage.dto));
            newDto.Id = newPageId;
            newDto.FileName = newName;
            newDto.FileRef = dstPath;
            newDto.FileLeafRef = newName;
            newDto.FileDirRef = newDir;
            newDto.UniqueId = nextFileId();
            // Live SPO (spike §11 Q5): the copy lands as a 0.1 draft, not
            // checked out, still carrying PromotedState — the §4.2 hazard.
            newDto._UIVersionString = '0.1';
            newDto.IsPageCheckedOutToCurrentUser = false;
            const newItem = JSON.parse(JSON.stringify(sourcePage.item));
            newItem.Id = newPageId;
            newItem.FileLeafRef = newName;
            newItem.FileRef = dstPath;
            newItem.FileDirRef = newDir;
            newItem.UniqueId = newDto.UniqueId;
            newItem._UIVersionString = '0.1';
            dstWeb.pages.set(newPageId, { dto: newDto, item: newItem });
            dstWeb.files.set(dstPath.toLowerCase(), {
              ...srcFile, Name: newName, ServerRelativeUrl: dstPath, UniqueId: newDto.UniqueId, CheckOutType: 2,
            });
          }
        }
      }

      if (isMove) {
        // Move: update file location. The old path has to be captured
        // BEFORE srcFile is mutated — matching page.dto.FileRef against
        // srcFile.ServerRelativeUrl after reassigning it would compare the
        // old FileRef against the NEW path and never match.
        const oldServerRelativeUrl = srcFile.ServerRelativeUrl;
        const newName = dstPath.split('/').pop();
        web.files.delete(srcPath.toLowerCase());
        srcFile.ServerRelativeUrl = dstPath;
        srcFile.Name = newName;
        web.files.set(dstPath.toLowerCase(), srcFile);

        // Update any associated page (FileName too — savepage's own rename
        // keeps it in step with FileLeafRef, so a move must as well).
        for (const page of web.pages.values()) {
          if (page.dto.FileRef && page.dto.FileRef.toLowerCase() === oldServerRelativeUrl.toLowerCase()) {
            page.dto.FileRef = dstPath;
            page.dto.FileLeafRef = newName;
            page.dto.FileName = newName;
            page.item.FileRef = dstPath;
            page.item.FileLeafRef = newName;
          }
        }
      }
    }
    return {};
  }

  // web/getfolderbyserverrelativepath(…)/files/addtemplatefile(…)
  if (/web\/getfolderbyserverrelativepath\(decodedurl='([^']*)'\)\/files\/addtemplatefile/i.test(pathLower)) {
    record();
    const match = /web\/getfolderbyserverrelativepath\(decodedurl='([^']*)'\)\/files\/addtemplatefile\(urloffile='([^']*)'/i.exec(pathLower);
    if (match) {
      let folderPath = match[1];
      let filePath = match[2];
      try { folderPath = decodeURIComponent(folderPath); folderPath = folderPath.replaceAll("''", "'"); } catch { /* keep raw */ }
      try { filePath = decodeURIComponent(filePath); filePath = filePath.replaceAll("''", "'"); } catch { /* keep raw */ }

      const baseName = filePath.split('/').pop();
      const fullPath = `${folderPath}/${baseName}`;

      web.files.set(fullPath.toLowerCase(), {
        Name: baseName,
        ServerRelativeUrl: fullPath,
        UniqueId: nextFileId(),
        ListId: PAGECOPY_IDS.srcSitePages,
        WebId: web.webId,
        SiteId: web.siteId,
        Length: 0,
        CheckOutType: 2,
        contentType: 'text/html',
      });

      return {
        Name: baseName,
        ServerRelativeUrl: fullPath,
        UniqueId: nextFileId(),
      };
    }
  }

  // web/getfolderbyserverrelativepath(…)/files/addusingpath(…)
  if (/web\/getfolderbyserverrelativepath\(decodedurl='([^']*)'\)\/files\/addusingpath/i.test(pathLower)) {
    record();
    const match = /web\/getfolderbyserverrelativepath\(decodedurl='([^']*)'\)\/files\/addusingpath\(decodedurl='([^']*)'/i.exec(pathLower);
    if (match) {
      let folderPath = match[1];
      let fileName = match[2];
      try { folderPath = decodeURIComponent(folderPath); folderPath = folderPath.replaceAll("''", "'"); } catch { /* keep raw */ }
      try { fileName = decodeURIComponent(fileName); fileName = fileName.replaceAll("''", "'"); } catch { /* keep raw */ }

      const fullPath = `${folderPath}${folderPath.endsWith('/') ? '' : '/'}${fileName}`;

      if (web.files.has(fullPath.toLowerCase())) {
        throw new SpFileError('File exists', { code: 'conflict', status: 409 });
      }

      const uniqueId = nextFileId();
      web.files.set(fullPath.toLowerCase(), {
        Name: fileName,
        ServerRelativeUrl: fullPath,
        UniqueId: uniqueId,
        ListId: PAGECOPY_IDS.srcSiteAssets,
        WebId: web.webId,
        SiteId: web.siteId,
        Length: bodyData?.length || body?.byteLength || 0,
        CheckOutType: 2,
        contentType: 'application/octet-stream',
      });

      return {
        ServerRelativeUrl: fullPath,
        Name: fileName,
        UniqueId: uniqueId,
        CheckOutType: 2,
      };
    }
  }

  // web/folders/addusingpath(decodedurl='…')
  if (/web\/folders\/addusingpath/i.test(pathLower)) {
    record();
    const match = /web\/folders\/addusingpath\(decodedurl='([^']*)'/i.exec(pathLower);
    if (match) {
      let folderPath = match[1];
      try { folderPath = decodeURIComponent(folderPath); folderPath = folderPath.replaceAll("''", "'"); } catch { /* keep raw */ }
      // Not idempotent on SPO: an existing folder is an error (spike §11).
      if (web.folders.has(folderPath.toLowerCase())) {
        throw new SpFileError(`A file or folder with the name ${folderPath.replace(/^\/+/, '')} already exists.`, { code: 'write', status: 500 });
      }
      web.folders.add(folderPath.toLowerCase());
      return { ServerRelativeUrl: folderPath };
    }
  }

  // sitepages/addimagefromexternalurl?…
  if (/sitepages\/addimagefromexternalurl/i.test(pathLower)) {
    record();
    const pageNameMatch = /pagename='([^']*)'/.exec(pathLower);
    const imageFileNameMatch = /imagefilename='([^']*)'/.exec(pathLower);
    if (pageNameMatch && imageFileNameMatch) {
      let pageName = pageNameMatch[1];
      let imageFileName = imageFileNameMatch[1];
      try { pageName = decodeURIComponent(pageName); imageFileName = decodeURIComponent(imageFileName); } catch { /* keep raw */ }

      const imagePath = `${web.path}/SiteAssets/SitePages/${pageName}/${imageFileName}`;
      web.files.set(imagePath.toLowerCase(), {
        Name: imageFileName,
        ServerRelativeUrl: imagePath,
        UniqueId: nextFileId(),
        ListId: PAGECOPY_IDS.srcSiteAssets,
        WebId: web.webId,
        SiteId: web.siteId,
        Length: 0,
        CheckOutType: 2,
        contentType: 'image/jpeg',
      });

      return { ServerRelativeUrl: imagePath };
    }
  }

  // web/lists(guid'L')/items(N)/validateupdatelistitem
  if (/web\/lists\(guid'([0-9a-f-]+)'\)\/items\((\d+)\)\/validateupdatelistitem/i.test(pathLower)) {
    record();
    const match = /web\/lists\(guid'([0-9a-f-]+)'\)\/items\((\d+)\)\/validateupdatelistitem/i.exec(pathLower);
    if (match) {
      const listId = match[1].toLowerCase();
      const itemId = Number(match[2]);

      const result = { value: [] };
      const formValues = bodyData.formValues || [];
      const list = web.lists.find((l) => l.Id.toLowerCase() === listId);
      const known = new Set((web.fields?.[list?.Id] || web.fields?.[listId] || []).map((f) => f.InternalName));
      const page = web.pages.get(itemId);
      for (const fv of formValues) {
        const row = { FieldName: fv.FieldName, FieldValue: fv.FieldValue, HasException: false, ErrorMessage: '' };
        // Audience required on pagedst only
        if (fv.FieldName === 'Audience' && web.path === '/sites/pagedst' && !fv.FieldValue) {
          row.HasException = true;
          row.ErrorMessage = 'Audience is required.';
        } else if (known.size && !known.has(fv.FieldName) && fv.FieldName !== 'PromotedState') {
          row.HasException = true;
          row.ErrorMessage = `Column '${fv.FieldName}' does not exist.`;
        } else if (page) {
          // Applied like SPO: onto the item, and the DTO mirrors the fields it
          // carries (Description is only settable this way — spike §11).
          page.item[fv.FieldName] = fv.FieldValue;
          if (fv.FieldName === 'Description') page.dto.Description = fv.FieldValue;
          if (fv.FieldName === 'Title') page.dto.Title = fv.FieldValue;
          if (fv.FieldName === 'PromotedState') { page.dto.PromotedState = Number(fv.FieldValue); page.item.PromotedState = Number(fv.FieldValue); }
        }
        result.value.push(row);
      }
      return result;
    }
  }

  // web/lists(guid'L')/items(N)/setcommentsdisabled
  if (/web\/lists\(guid'([0-9a-f-]+)'\)\/items\((\d+)\)\/setcommentsdisabled/i.test(pathLower)) {
    record();
    const match = /web\/lists\(guid'([0-9a-f-]+)'\)\/items\((\d+)\)\/setcommentsdisabled/i.exec(pathLower);
    if (match) {
      const pageId = Number(match[2]);
      const page = web.pages.get(pageId);
      if (page) {
        page.item.CommentsDisabled = bodyData.value === true;
        page.dto.CommentsDisabled = bodyData.value === true;
      }
    }
    return {};
  }

  // web/getfilebyserverrelativepath(decodedurl='…')/checkout()
  if (/web\/getfilebyserverrelativepath\(decodedurl='([^']*)'\)\/checkout\(\)/i.test(pathLower)) {
    record();
    const match = /web\/getfilebyserverrelativepath\(decodedurl='([^']*)'\)\/checkout\(\)/i.exec(pathLower);
    if (match) {
      let decodedPath = match[1];
      try { decodedPath = decodeURIComponent(decodedPath); } catch { /* keep raw */ }
      const srLower = decodedPath.toLowerCase();
      const file = Array.from(web.files.entries()).find(([key]) => key === srLower)?.[1];
      if (!file) throw new SpFileError('File not found', { code: 'not-found', status: 404 });
      file.CheckOutType = 0; // 0 = checked out online
      for (const page of web.pages.values()) {
        if (page.dto.FileRef && page.dto.FileRef.toLowerCase() === srLower) {
          page.dto.IsPageCheckedOutToCurrentUser = true;
        }
      }
    }
    return {};
  }

  // web/getfilebyserverrelativepath(decodedurl='…')/checkin(comment='…',checkintype=N)
  if (/web\/getfilebyserverrelativepath\(decodedurl='([^']*)'\)\/checkin\(/i.test(pathLower)) {
    record();
    const match = /web\/getfilebyserverrelativepath\(decodedurl='([^']*)'\)\/checkin\(/i.exec(pathLower);
    if (match) {
      let decodedPath = match[1];
      try { decodedPath = decodeURIComponent(decodedPath); } catch { /* keep raw */ }
      const srLower = decodedPath.toLowerCase();
      const file = Array.from(web.files.entries()).find(([key]) => key === srLower)?.[1];
      if (file) {
        file.CheckOutType = 2; // 2 = none — checked in
        for (const page of web.pages.values()) {
          if (page.dto.FileRef && page.dto.FileRef.toLowerCase() === srLower) {
            page.dto.IsPageCheckedOutToCurrentUser = false;
            if (page.item) page.item.CheckoutUser = null;
            // Bump a minor version, keeping the page item's dto in step.
            const label = page.dto._UIVersionString || '0.1';
            const [major, minor] = label.split('.').map(Number);
            const newLabel = `${major || 0}.${(minor || 0) + 1}`;
            page.dto._UIVersionString = newLabel;
            if (page.item) page.item._UIVersionString = newLabel;
          }
        }
      }
    }
    return {};
  }

  // Unrecognized write endpoint for a page-copy web
  throw new SpFileError(`mock: no handler for ${url}`, { code: 'not-found', status: 404 });
  } // end dispatchWrite()
}

// ---- Global state export ----

globalThis.__PAGECOPY_MOCK__ = { reset: resetPageCopyMock, state: webState };
