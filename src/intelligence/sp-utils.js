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
  'getListSchema',
  'exportListData',
  'createListFromSchema',
  'importListData',
  'copyList',
  'downloadJson',
  'readJsonFile',
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

/** Structure document produced by getListSchema(); input to createListFromSchema(). */
interface SPUtilsSchemaDoc {
  kind: 'dcspad-sputils-list-schema';
  version: 1;
  exported: string;
  source: { siteUrl: string; listTitle: string; listId: string; rootFolder: string | null; itemCount: number };
  list: {
    title: string; description: string; baseTemplate: number;
    enableVersioning: boolean; majorVersionLimit: number | null; enableMinorVersions: boolean;
    majorWithMinorVersionsLimit: number | null; draftVersionVisibility: number; forceCheckout: boolean;
    hidden: boolean; contentTypesEnabled: boolean; enableAttachments: boolean;
    enableFolderCreation: boolean; enableModeration: boolean;
  };
  fields: Array<{
    /** Source field id; dependent lookups are rebound through it. */
    id: string;
    internalName: string; staticName: string; displayName: string; type: string;
    required: boolean; readOnly: boolean; fromBaseType: boolean;
    /** true when the list author added it (not inherited, deletable). Only custom fields are recreated. */
    custom: boolean;
    description: string; defaultValue: string | null; choices: string[]; maxLength: number | null;
    indexed: boolean; enforceUniqueValues: boolean; allowMultipleValues: boolean; customFormatter: string;
    lookupListId: string | null;
    /** Title of the lookup's target list (resolved from its GUID). */
    lookupList: string | null;
    lookupField: string | null; isSelfLookup: boolean;
    /** A secondary lookup column that depends on a primary lookup. */
    isDependentLookup: boolean; primaryFieldId: string | null;
    schemaXml: string;
  }>;
  views: Array<{
    id: string; title: string; defaultView: boolean; hidden: boolean; viewType: string;
    viewQuery: string; rowLimit: number; paged: boolean; customFormatter: string; jsLink: string;
    fields: string[];
  }>;
  contentTypes: Array<{ name: string; id: string; description?: string; group?: string; hidden?: boolean; readOnly?: boolean }>;
  warnings: string[];
}

/** Item document produced by exportListData(); input to importListData(). */
interface SPUtilsDataDoc {
  kind: 'dcspad-sputils-list-data';
  version: 1;
  exported: string;
  source: { siteUrl: string; listTitle: string; listId: string };
  /** Compact per-field metadata keyed by internal name. */
  fields: Record<string, {
    type: string; custom: boolean; readOnly: boolean;
    lookupList: string | null; lookupListId: string | null; lookupField: string | null; isSelfLookup: boolean; allowMultipleValues: boolean;
  }>;
  /** Raw REST rows plus resolved people/lookups and optional attachments. */
  items: Array<SPUtilsListItem & {
    _resolved: Record<string, Array<{ Id: number; Email?: string; LoginName?: string; Title?: string; value?: string | number | boolean | null }>>;
    _attachments?: Array<{ name: string; base64: string }>;
    /** Folder path (relative to the list root) the item lives in; empty at the root. */
    _dir?: string;
    /** true for a folder row; _folderPath is its own relative path. */
    _folder?: boolean; _folderPath?: string;
  }>;
  warnings: string[];
}

/** Report returned by createListFromSchema(). */
interface SPUtilsCreateReport {
  dryRun: boolean; title: string; listId: string | null; created: boolean;
  fields: { added: number; skipped: number; failed: Array<{ internalName: string; displayName: string; error: string }> };
  views: { added: number; updated: number; failed: Array<{ title: string; error: string }> };
  warnings: string[];
}

/** Report returned by importListData(). */
interface SPUtilsImportReport {
  dryRun: boolean; listTitle: string; attempted: number; created: number;
  /** Source items skipped because a seeded idMap already had them. */
  skipped: number;
  failed: Array<{ sourceId: number; error: string }>;
  fieldErrors: Array<{ sourceId: number; field: string; message: string }>;
  /** Source item Id → new item Id. */
  idMap: Record<number, number>;
  attachments: { added: number; failed: number };
  folders: { created: number; failed: number };
  warnings: string[];
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

  /**
   * **Copy step 1.** Capture settings, every visible field (scrub-ready
   * SchemaXml, lookup targets as list titles, column formatting), views and
   * content types. No item data.
   * @example const schema = await SPUtils.getListSchema("Requests")
   */
  getListSchema(listTitle: string): Promise<SPUtilsSchemaDoc>;

  /**
   * **Copy step 2.** Export every item as raw REST values plus resolved
   * people (email/login) and lookups (shown value), so they can be re-resolved
   * on another site. Attachments are embedded only when asked.
   * @example const data = await SPUtils.exportListData("Requests", { includeAttachments: true })
   */
  exportListData(listTitle: string, options?: {
    /** Embed attachment bytes as base64 (default false). */
    includeAttachments?: boolean;
    /** Reuse a getListSchema() result instead of reading the schema again. */
    schema?: SPUtilsSchemaDoc;
  }): Promise<SPUtilsDataDoc>;

  /**
   * **Copy step 3.** Create the list on the current site from a schema
   * document, or add missing columns and views to an existing list of that
   * title. Columns are created plain → lookup → calculated.
   * @example await SPUtils.createListFromSchema(schema, { title: "Requests", dryRun: true })
   */
  createListFromSchema(schema: SPUtilsSchemaDoc, options?: {
    /** Target list title (default: the source title). */
    title?: string;
    description?: string;
    /** Log the plan and change nothing (default false). */
    dryRun?: boolean;
    /** Source lookup-list title → target list title, for lookups whose target has a different name here. */
    lookupMap?: Record<string, string>;
  }): Promise<SPUtilsCreateReport>;

  /**
   * **Copy step 4.** Create the items in the current site's list. Values use
   * the ValidateUpdateListItem conventions; people are re-resolved by email,
   * lookups by shown value, self-referencing lookups in a second pass.
   * @example await SPUtils.importListData(data, { listTitle: "Requests", dryRun: true })
   */
  importListData(data: SPUtilsDataDoc, options?: {
    /** Target list (default: the source title). */
    listTitle?: string;
    /** Show sample payloads and counts, create nothing (default false). */
    dryRun?: boolean;
    /** Parallel creates (default 1, which keeps source order; higher is faster but scrambles new ids). */
    concurrency?: number;
    /** Keep Created/Modified/Author/Editor from the source (default false). Seconds are not preserved; app/system authors fall back to the importing user. */
    preserveAuthorship?: boolean;
    /** Re-upload embedded attachments (default true). */
    includeAttachments?: boolean;
    /** A previous report's idMap: source items already in it are skipped, so a partial import resumes without duplicates. */
    idMap?: Record<number, number>;
  }): Promise<SPUtilsImportReport>;

  /**
   * Steps 1–4 chained. On the same site the copy is named "<title> Copy";
   * with targetSite the context switches before create/import and the title
   * is kept. Returns each step's output so a partial run can be resumed.
   * @example await SPUtils.copyList("Requests", { targetSite: "sites/Archive", dryRun: true })
   */
  copyList(sourceTitle: string, options?: {
    targetTitle?: string;
    /** Site path or absolute same-tenant URL; passed to setupContext(). */
    targetSite?: string;
    dryRun?: boolean;
    includeAttachments?: boolean;
    preserveAuthorship?: boolean;
    lookupMap?: Record<string, string>;
  }): Promise<{ schema: SPUtilsSchemaDoc; data: SPUtilsDataDoc; created: SPUtilsCreateReport; imported: SPUtilsImportReport; targetSite: string }>;

  /**
   * Save any object as a pretty-printed .json download.
   * @example SPUtils.downloadJson(schema, "Requests.schema.json")
   */
  downloadJson(object: unknown, fileName?: string): void;

  /**
   * Read and parse a .json file stored in a library on the current site.
   * @example const data = await SPUtils.readJsonFile("/sites/X/Shared Documents/Requests.data.json")
   */
  readJsonFile<T = any>(filePath: string): Promise<T>;
}

/** DCSPad SP Utilities. Run \`SPUtils.help()\` for the function list. */
declare const SPUtils: SPUtilsApi;

interface Window {
  /** DCSPad SP Utilities. Run \`SPUtils.help()\` for the function list. */
  SPUtils: SPUtilsApi;
}
`,
}];
