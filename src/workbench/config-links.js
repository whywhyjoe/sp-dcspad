// Curated one-click jumps to SharePoint configuration panels, expressed as
// data so wrong or tenant-variant URLs are one-line fixes. Paths are
// web-relative (they start with '/', appended to the inspected web's URL).
//
// Deliberately SHORT: this is a quick-jump panel, not a mirror of the SP
// backend. Joe curated the set on 2026-07-31 — don't grow it without asking.
// Entries whose exact path varies by tenant vintage carry a `hint`.

export const LIST_SETTINGS = {
  label: 'List settings',
  path: '/_layouts/15/listedit.aspx?List={guid}',
};

export const LINK_GROUPS = [
  {
    title: 'General',
    links: [
      { label: 'Site settings', path: '/_layouts/15/settings.aspx' },
      { label: 'Site contents', path: '/_layouts/15/viewlsts.aspx' },
      { label: 'Site information', path: '/_layouts/15/prjsetng.aspx', hint: 'Classic title/description/logo page' },
      { label: 'Site usage', path: '/_layouts/15/siteanalytics.aspx' },
    ],
  },
  {
    title: 'Permissions & people',
    links: [
      { label: 'Site permissions', path: '/_layouts/15/user.aspx' },
      { label: 'People and groups', path: '/_layouts/15/people.aspx?MembershipGroupId=0' },
      { label: 'Groups', path: '/_layouts/15/groups.aspx' },
      { label: 'Access requests', path: '/Access%20Requests/pendingreq.aspx', hint: 'Only exists once an access request has been made' },
    ],
  },
  {
    title: 'Recycle bins',
    links: [
      { label: 'Recycle bin', path: '/_layouts/15/RecycleBin.aspx' },
      { label: 'Site collection recycle bin', path: '/_layouts/15/AdminRecycleBin.aspx', hint: 'Site-collection scope — needs admin rights' },
      { label: 'Second-stage recycle bin', path: '/_layouts/15/AdminRecycleBin.aspx?View=5', hint: 'Deleted-from-end-user-bin view; needs admin rights' },
    ],
  },
  {
    title: 'Galleries',
    links: [
      { label: 'Site columns', path: '/_layouts/15/mngfield.aspx' },
      { label: 'Site content types', path: '/_layouts/15/mngctype.aspx' },
      { label: 'Themes gallery', path: '/_catalogs/theme/Forms/AllItems.aspx' },
    ],
  },
  {
    title: 'Search',
    links: [
      { label: 'Search settings', path: '/_layouts/15/enhancedSearch.aspx?level=site', hint: 'Verify level param on your tenant' },
      { label: 'Search schema (managed properties)', path: '/_layouts/15/listmanagedproperties.aspx?level=site', hint: 'Verify level param on your tenant' },
      { label: 'Result sources', path: '/_layouts/15/manageresultsources.aspx?level=site', hint: 'Verify level param on your tenant' },
      { label: 'Query rules', path: '/_layouts/15/listqueryrules.aspx?level=site', hint: 'Verify level param on your tenant' },
      { label: 'Searchable columns', path: '/_layouts/15/NoCrawlSettings.aspx' },
    ],
  },
];

// Resolve a link entry against the inspected web. `params` substitutes
// {placeholders} in the path; values are URL-encoded.
export function linkUrl(webUrl, link, params = {}) {
  let path = String(link?.path || '');
  for (const [key, value] of Object.entries(params)) {
    path = path.replaceAll(`{${key}}`, encodeURIComponent(String(value)));
  }
  return `${String(webUrl || '').replace(/\/+$/, '')}${path}`;
}
