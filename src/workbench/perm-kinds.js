// SPBasePermissions decoding. REST returns BasePermissions as two unsigned
// 32-bit halves serialized as strings ({ High, Low }); the real enum is a
// 64-bit mask, so recombine with BigInt before testing bits.
//
// Flag values are the documented SP.PermissionKind / SPBasePermissions
// constants (bit positions, not the 1-based enum ordinals).

const FLAGS = [
  // [name, bit] — bit as BigInt exponent in the combined 64-bit mask.
  ['ViewListItems', 0n],
  ['AddListItems', 1n],
  ['EditListItems', 2n],
  ['DeleteListItems', 3n],
  ['ApproveItems', 4n],
  ['OpenItems', 5n],
  ['ViewVersions', 6n],
  ['DeleteVersions', 7n],
  ['CancelCheckout', 8n],
  ['ManagePersonalViews', 9n],
  ['ManageLists', 11n],
  ['ViewFormPages', 12n],
  ['AnonymousSearchAccessList', 13n],
  ['Open', 16n],
  ['ViewPages', 17n],
  ['AddAndCustomizePages', 18n],
  ['ApplyThemeAndBorder', 19n],
  ['ApplyStyleSheets', 20n],
  ['ViewUsageData', 21n],
  ['CreateSSCSite', 22n],
  ['ManageSubwebs', 23n],
  ['CreateGroups', 24n],
  ['ManagePermissions', 25n],
  ['BrowseDirectories', 26n],
  ['BrowseUserInfo', 27n],
  ['AddDelPrivateWebParts', 28n],
  ['UpdatePersonalWebParts', 29n],
  ['ManageWeb', 30n],
  ['AnonymousSearchAccessWebLists', 32n],
  ['UseClientIntegration', 36n],
  ['UseRemoteAPIs', 37n],
  ['ManageAlerts', 38n],
  ['CreateAlerts', 39n],
  ['EditMyUserInfo', 40n],
  ['EnumeratePermissions', 62n],
];

const FULL_MASK = 0x7FFFFFFFFFFFFFFFn;

export function combineBasePermissions(basePermissions) {
  const high = BigInt(String(basePermissions?.High ?? '0'));
  const low = BigInt(String(basePermissions?.Low ?? '0'));
  return (high << 32n) | low;
}

// -> { flags: string[], isFullControl, isEmpty }
export function decodeBasePermissions(basePermissions) {
  const mask = combineBasePermissions(basePermissions);
  if ((mask & FULL_MASK) === FULL_MASK) {
    return { flags: ['FullMask (all permissions)'], isFullControl: true, isEmpty: false };
  }
  const flags = FLAGS.filter(([, bit]) => (mask & (1n << bit)) !== 0n).map(([name]) => name);
  return { flags, isFullControl: false, isEmpty: flags.length === 0 };
}

export const PRINCIPAL_TYPE_NAMES = {
  1: 'User',
  2: 'Distribution list',
  4: 'Security group',
  8: 'SharePoint group',
};

export const principalTypeName = (v) => PRINCIPAL_TYPE_NAMES[v] || String(v ?? '');
