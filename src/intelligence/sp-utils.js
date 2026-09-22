// DCSPad SP Utilities (window.SPUtils) intelligence for Monaco's JavaScript
// surface. The runtime is utilities/dcspad-sp-utilities.js, a console-first
// PnPjs 2 script that DCSPad injects from its own framework catalog entry.
//
// Keep this declaration in step with the script's USAGE registry and export
// block. tests/monaco.mjs compares SP_UTILS_FUNCTION_NAMES against the keys
// the runtime actually exposes, so a drift fails the suite rather than
// shipping stale completions.

export const SP_UTILS_PACK_ID = 'dcspad-sp-utilities';

// Every public function the runtime exports, in the script's export order.
export const SP_UTILS_FUNCTION_NAMES = [
  'help',
  'clearStatus',
  'setupContext',
  'previewListItems',
  'deleteListItems',
  'getAllLists',
  'getAllSecurityGroups',
  'getSiteMembersWithGroups',
  'getSiteGroupMembers',
  'getSharingLinkForItem',
  'removeAllSharingLinksForItem',
  'addUsersFromCSVToGroups',
  'syncUsersFromCSVToGroups',
  'addFieldsToList',
  'exportListToExcel',
  'getListFields',
  'getListSchemaForMigration',
];

export const SP_UTILS_JS_LIBRARIES = [{
  filePath: 'file:///node_modules/@types/dcspad-sp-utilities/index.d.ts',
  content: `/**
 * DCSPad SP Utilities — console-first SharePoint admin helpers on PnPjs 2.
 * Exposed as the global \`SPUtils\` by utilities/dcspad-sp-utilities.js.
 * Requires PnPjs 2 on the page as \`pnp2\`. Call setupContext() first.
 */

/** One row of the usage registry that drives help() and the on-page panel. */
interface SPUtilsUsageEntry {
  /** Function name, e.g. "getAllLists". */
  readonly name: string;
  /** Signature as shown in the panel, e.g. "(listTitle)". */
  readonly sig: string;
  /** One-line purpose. */
  readonly what: string;
  /** A copyable example call. */
  readonly example: string;
}

/** A list item as returned by the paging helpers (Id and Title plus any other field). */
interface SPUtilsListItem {
  Id: number;
  Title?: string;
  [field: string]: any;
}

/** Field metadata as returned by getListFields(). */
interface SPUtilsFieldInfo {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Required: boolean;
  ReadOnlyField: boolean;
  Description?: string;
  DefaultValue?: string | null;
  Choices?: string[] | { results: string[] };
  MaxLength?: number;
}

/** The migration schema produced by getListSchemaForMigration(). */
interface SPUtilsListSchema {
  /** ISO timestamp of the export. */
  exported: string;
  list: {
    title: string;
    description: string;
    baseTemplate: number;
    enableVersioning: boolean;
    enableMinorVersions: boolean;
    forceCheckout: boolean;
    hidden: boolean;
    contentTypesEnabled: boolean;
    enableAttachments: boolean;
  };
  fields: Array<{
    internalName: string;
    displayName: string;
    type: string;
    required: boolean;
    readOnly: boolean;
    description?: string;
    defaultValue?: string | null;
    choices: string[];
    maxLength?: number;
    /** The field's SchemaXml, usable with createFieldAsXml after scrubbing ids. */
    schemaXml: string;
  }>;
  contentTypes: Array<{ Name: string; StringId: string; Description?: string; Group?: string; Hidden?: boolean; ReadOnly?: boolean }>;
  views: Array<{ Title: string; DefaultView: boolean; Paged: boolean; RowLimit: number; ViewQuery: string; Fields: string[] }>;
}

/** A column definition accepted by addFieldsToList(). */
interface SPUtilsFieldSpec {
  /** Display name; the internal name is derived from it unless \`name\` is given. */
  displayName?: string;
  /** Explicit internal name. */
  name?: string;
  /** One of: text, note, multiline, boolean, yes/no, choice, url, user, person. Defaults to text. */
  type?: 'text' | 'note' | 'multiline' | 'boolean' | 'yes/no' | 'choice' | 'url' | 'user' | 'person';
  /** Choice values; required when type is "choice". */
  choices?: string[];
}

interface SPUtilsGroupFailure {
  user: string;
  group: string;
  error: string;
}

interface SPUtilsPerGroupCounts {
  attempted: number;
  added: number;
  skippedExisting: number;
  failed: number;
}

/** Result of addUsersFromCSVToGroups(). */
interface SPUtilsAddResult {
  attempted: number;
  added: number;
  skippedExisting: number;
  failed: SPUtilsGroupFailure[];
  perGroup: Record<string, SPUtilsPerGroupCounts>;
  dryRun: boolean;
}

/** Result of syncUsersFromCSVToGroups(). */
interface SPUtilsSyncResult {
  dryRun: boolean;
  groups: number;
  totals: {
    addPlanned: number;
    added: number;
    removePlanned: number;
    removed: number;
    skippedExisting: number;
  };
  perGroup: Record<string, {
    addPlanned: number;
    added: number;
    removePlanned: number;
    removed: number;
    skippedExisting: number;
    errors: number;
  }>;
  failed: Array<{ email: string; group: string; action: 'add' | 'remove' | 'parse'; error: string }>;
  attempted: number;
  added: number;
  skippedExisting: number;
}

/** Result of addFieldsToList(). */
interface SPUtilsAddFieldsResult {
  added: number;
  skipped: number;
  failed: Array<{ name: string; error: string }>;
}

interface SPUtilsApi {
  /**
   * Print every function to the console as a table and open the on-page
   * usage panel. Returns the usage registry.
   * @example SPUtils.help()
   */
  help(): SPUtilsUsageEntry[];

  /** Empty the on-page status log. */
  clearStatus(): void;

  /**
   * **Call first.** Point PnPjs at a site. Accepts a site path relative to the
   * current tenant ("sites/Project", "teams/Marketing") or an absolute
   * same-tenant URL. Every later call uses this context.
   * @example SPUtils.setupContext("sites/Project")
   */
  setupContext(siteSubUrl: string): void;

  /**
   * Fetch every item in a list and print Id and Title as a table.
   * @param listTitle Display name of the list.
   * @param batchSize Items per request (default 100).
   * @example await SPUtils.previewListItems("Requests")
   */
  previewListItems(listTitle: string, batchSize?: number): Promise<void>;

  /**
   * Batch-delete list items, optionally filtered by a predicate.
   * Run with \`dryRun: true\` first: it prints what would be deleted and changes nothing.
   * @example await SPUtils.deleteListItems("Requests", { dryRun: true, filterFn: i => /temp/i.test(i.Title) })
   */
  deleteListItems(listTitle: string, options?: {
    /** Items per batch request (default 100). */
    batchSize?: number;
    /** Preview only; nothing is deleted (default false). */
    dryRun?: boolean;
    /** Keep only the items this returns true for. */
    filterFn?: ((item: SPUtilsListItem) => boolean) | null;
  }): Promise<void>;

  /**
   * Print every list on the site (Id, Title) as a table.
   * @example await SPUtils.getAllLists()
   */
  getAllLists(): Promise<void>;

  /**
   * Print every site group (Id, Title) as a table.
   * @example await SPUtils.getAllSecurityGroups()
   */
  getAllSecurityGroups(): Promise<void>;

  /**
   * Print each site member with the groups they belong to.
   * @example await SPUtils.getSiteMembersWithGroups()
   */
  getSiteMembersWithGroups(): Promise<void>;

  /**
   * Print each group with its members, including job title and department
   * from the user profile. System accounts are skipped.
   * @example await SPUtils.getSiteGroupMembers()
   */
  getSiteGroupMembers(): Promise<void>;

  /**
   * Create a sharing link for a file and return its URL.
   * @param filePath Server-relative path, e.g. "/sites/X/Shared Documents/a.pdf".
   * @param canEdit true for an edit link, false for view (default false).
   * @param expireInDays Days until expiry; 0 means no expiry (default 1).
   * @example await SPUtils.getSharingLinkForItem("/sites/X/Shared Documents/a.pdf", false, 7)
   */
  getSharingLinkForItem(filePath: string, canEdit?: boolean, expireInDays?: number): Promise<string>;

  /**
   * Revoke every sharing link on a file. Does nothing unless \`areYouSure\` is true.
   * @returns true on success; undefined when not confirmed.
   * @example await SPUtils.removeAllSharingLinksForItem("/sites/X/Shared Documents/a.pdf", true)
   */
  removeAllSharingLinksForItem(filePath: string, areYouSure?: boolean): Promise<boolean | undefined>;

  /**
   * Add users to groups from a CSV stored in a library. The CSV has either
   * "Email Address" and "Group" header columns, or two headerless columns
   * (email, group). Lines starting with # are ignored.
   * @param filePath Server-relative path to the CSV.
   * @example await SPUtils.addUsersFromCSVToGroups("/sites/X/Shared Documents/users.csv", { dryRun: true })
   */
  addUsersFromCSVToGroups(filePath: string, options?: {
    /** Preview only (default false). */
    dryRun?: boolean;
    /** Parallel adds (default 6). */
    concurrency?: number;
  }): Promise<SPUtilsAddResult>;

  /**
   * Make group membership match a master CSV: users in the CSV are added,
   * members not in the CSV are **removed**. Rows whose Group is "ALL" join
   * every group and are never removed. \`adminEmail\` is required and is
   * always added, never removed. \`dryRun\` defaults to true.
   * @example await SPUtils.syncUsersFromCSVToGroups("/sites/X/Shared Documents/master.csv", { dryRun: true, adminEmail: "me@contoso.com" })
   */
  syncUsersFromCSVToGroups(filePath: string, options: {
    /** Email/UPN that is always added and never removed. Required. */
    adminEmail: string;
    /** Preview only (default true). */
    dryRun?: boolean;
    /** Parallelism per group (default 6). */
    concurrency?: number;
    /** Skip removals for obvious system accounts (default true). */
    excludeSystem?: boolean;
    /** A group synced to the union of every other group's members. */
    umbrellaGroup?: string;
  }): Promise<SPUtilsSyncResult>;

  /**
   * Create list columns from a JSON spec. Existing columns are skipped, so
   * the call is safe to repeat.
   * @param listServerRelativeUrl e.g. "/sites/X/Lists/Requests".
   * @example await SPUtils.addFieldsToList("/sites/X/Lists/Requests", [{ displayName: "Priority", type: "choice", choices: ["Low", "High"] }], { dryRun: true })
   */
  addFieldsToList(listServerRelativeUrl: string, specs: SPUtilsFieldSpec[], options?: {
    /** Log planned columns without creating them (default false). */
    dryRun?: boolean;
  }): Promise<SPUtilsAddFieldsResult>;

  /**
   * Download every item in a list as an .xlsx workbook with a field-metadata
   * sheet. Requires ExcelJS on the page.
   * @example await SPUtils.exportListToExcel("Requests")
   */
  exportListToExcel(listTitle: string, options?: {
    /** Output file name (default: list title plus date). */
    fileName?: string;
  }): Promise<void>;

  /**
   * Fetch the visible fields of a list with type, required and read-only
   * flags. Prints a table and returns the array.
   * @example const fields = await SPUtils.getListFields("Requests")
   */
  getListFields(listTitle: string): Promise<SPUtilsFieldInfo[]>;

  /**
   * Capture list settings, field definitions (with SchemaXml), content types
   * and views as one object, for rebuilding the list on another site.
   * Item data is not included.
   * @example const schema = await SPUtils.getListSchemaForMigration("Requests")
   */
  getListSchemaForMigration(listTitle: string): Promise<SPUtilsListSchema>;
}

/** DCSPad SP Utilities. Run \`SPUtils.help()\` for the function list. */
declare const SPUtils: SPUtilsApi;

interface Window {
  /** DCSPad SP Utilities. Run \`SPUtils.help()\` for the function list. */
  SPUtils: SPUtilsApi;
}
`,
}];
