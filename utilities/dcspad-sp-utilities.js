const lastModified = `Last modified: 2026/09/21 17:05:00
`;
console.log(`[SPUtils] ${lastModified}
  ` 
);

/**
 * @fileoverview  SharePoint Utility Functions (PnPjs v2 Integration)
 *
 * @author Joseph Zapert; FCU Creative & Digital Services
 * @created Created: 2025/09/19 15:36:18
 * @lastmodified Last modified: 2026/09/21 17:05:00
 * 
 * 
 * A comprehensive collection of SharePoint management utilities using PnPjs (v2 branch),
 * exposed globally as the `pnp2` object. Designed for embedding in a SharePoint page
 * and interactive use via browser console.
 * 
 * **Core Capabilities:**
 * - SharePoint site context management and initialization
 * - List and item operations (preview, batch delete with filtering, dry-run support)
 * - Security group and membership discovery & reporting
 * - Sharing link generation and removal with expiration control
 * - Bulk user-to-group assignments from CSV (add-only or full sync modes)
 * - Concurrent request handling with exponential backoff retry for throttling (429 errors)
 * - Flexible CSV parsing (header detection, two-column format, comments, quoted fields)
 * 
 * **PnPjs (v2 Branch) Usage:**
 * The library uses PnPjs v2 (a JavaScript abstraction over SharePoint REST APIs) exposed
 * as the global `pnp2` object. All SharePoint operations are performed through this object:
* - Context setup via `pnp2.sp.setup()`
 * 
 * **Setup & Requirements:**
 * - No page markup required: the script draws its own status panel (with usage
 *   instructions) on first use. If the page already has <p id="status"></p>, that
 *   element is reused instead.
 * - Load this script after PnPjs is loaded on the page: <script src="pnpjs-bundle.js"></script>
 * - Call setupContext(siteSubUrl) before any list/group operations
 * - Example: window.SPUtils.setupContext("teams/MyTeamSite")
 * 
 * **Testing/Console Access:**
 * In DCSPad, enable "DCSPad SP Utilities (SPUtils)" in the framework catalog
 * (below PnPjs) and call window.SPUtils.functionName(params...) from the JS pane.
 * On any tenant page, load pnp2 then this file and use the browser console.
 * 
 * @example
 * // Initialize the context for a specific SharePoint site
 * window.SPUtils.setupContext("teams/Marketing");
 * 
 * @example
 * // Preview all items in a list
 * window.SPUtils.previewListItems("My List", 100);
 * 
 * @example
 * // Add users from a CSV file to SharePoint groups
 * const result = await window.SPUtils.addUsersFromCSVToGroups(
 *   "/sites/Marketing/Shared Documents/users.csv",
 *   { dryRun: true, concurrency: 6 }
 * );
 * console.table(result.failed); // See any failed assignments
 * 
 * @example
 * // Full sync: add/remove users from groups based on a master CSV
 * const syncResult = await window.SPUtils.syncUsersFromCSVToGroups(
 *   "/sites/Marketing/Shared Documents/master-list.csv",
 *   { dryRun: false, adminEmail: "admin@contoso.com", umbrellaGroup: "Marketing Team" }
 * );
 * 
 *  * =============================
 * @section Table of Contents
 * =============================
 *
 * @function setupContext(siteSubUrl)
 *   Initializes PnPjs context for a specific SharePoint site. Must be called before any other operation.
 *
 * @function help()
 *   Lists every function (console.table) and opens the on-page usage panel.
 *
 * @function clearStatus()
 *   Empties the on-page status log.
 *
 * @function previewListItems(listTitle, batchSize)
 *   Fetches and displays all items from a SharePoint list in a paginated table (console.table).
 *
 * @function deleteListItems(listTitle, options)
 *   Batch-deletes items from a SharePoint list with optional dry-run and filtering support.
 *
 * @function getAllLists()
 *   Retrieves and displays all lists on the current SharePoint site (console.table).
 *
 * @function getListSchemaForMigration(listTitle)
 * Retrieves SharePoint list schema information including list settings, field
 * definitions, content types, and views. Useful for migration, backup, and
 * recreating list structures on another SharePoint site.
 * 
 * @function getAllSecurityGroups()
 *   Retrieves and displays all SharePoint security groups on the current site (console.table).
 *
 * @function getSiteMembersWithGroups()
 *   Displays site members mapped to their group memberships (user → groups).
 *
 * @function getSiteGroupMembers()
 *   Displays all SharePoint groups with their member details, including profile info.
 *
 * @function getSharingLinkForItem(filePath, canEdit, expireInDays)
 *   Generates an expiring sharing link (read or edit) for a file.
 *
 * @function removeAllSharingLinksForItem(filePath, areYouSure)
 *   Revokes ALL sharing links for a file. Requires explicit confirmation.
 *
 * @function addUsersFromCSVToGroups(filePath, options)
 *   Bulk-adds users to groups from a CSV file. Supports dry-run and concurrency.
 * 
 * @function addFieldsToList(listTitle, fields)
 *  Adds new fields to a SharePoint list based on a provided configuration array.
 *
 * @function syncUsersFromCSVToGroups(filePath, options)
 *   Synchronizes group membership to match a master CSV (add/remove users, supports umbrella group).
 *
 * (Internal) writeStatus, retryWithBackoff
 *   Internal helpers for status logging and retrying throttled requests.
 */


let siteBaseUrl = null;
let contextInitialized = false;

window.SPUtils = (() => {
    // The tenant is whatever page this script runs on; nothing is hardcoded.
    const SPTenantBase = `${location.origin}/`;

    // ---------------------------------------------------------------
    // USAGE registry — one entry per public function. This drives the
    // on-page usage panel and SPUtils.help(); keep it in sync with the
    // export block at the bottom of this file.
    // ---------------------------------------------------------------
    const USAGE = [
      { name: "setupContext", sig: "(siteSubUrl)", what: "Call FIRST. Points PnPjs at a site on this tenant (path or absolute URL).",
        example: `SPUtils.setupContext("sites/MyProject")` },
      { name: "help", sig: "()", what: "Print this function list to the console and open the usage panel.",
        example: `SPUtils.help()` },
      { name: "clearStatus", sig: "()", what: "Empty the status log on the page.",
        example: `SPUtils.clearStatus()` },
      { name: "getAllLists", sig: "()", what: "Table of every list on the site (Id, Title).",
        example: `await SPUtils.getAllLists()` },
      { name: "getListFields", sig: "(listTitle)", what: "Visible fields with type, required, read-only. Returns the array.",
        example: `const f = await SPUtils.getListFields("Requests")` },
      { name: "getListSchemaForMigration", sig: "(listTitle)", what: "Settings + fields + views + content types as one object, for rebuilding the list elsewhere.",
        example: `const schema = await SPUtils.getListSchemaForMigration("Requests")` },
      { name: "previewListItems", sig: "(listTitle, batchSize = 100)", what: "Table of all items (Id, Title).",
        example: `await SPUtils.previewListItems("Requests")` },
      { name: "deleteListItems", sig: "(listTitle, { dryRun, filterFn, batchSize })", what: "Batch-delete items. Use dryRun: true first.",
        example: `await SPUtils.deleteListItems("Requests", { dryRun: true, filterFn: i => /temp/i.test(i.Title) })` },
      { name: "addFieldsToList", sig: "(listServerRelativeUrl, specs, { dryRun })", what: "Create columns from a JSON spec. Skips existing columns.",
        example: `await SPUtils.addFieldsToList("/teams/MySite/Lists/Requests", [{ displayName: "Priority", type: "choice", choices: ["Low","High"] }], { dryRun: true })` },
      { name: "exportListToExcel", sig: "(listTitle, { fileName })", what: "Download every item as .xlsx. Needs ExcelJS on the page.",
        example: `await SPUtils.exportListToExcel("Requests")` },
      { name: "getAllSecurityGroups", sig: "()", what: "Table of site groups (Id, Title).",
        example: `await SPUtils.getAllSecurityGroups()` },
      { name: "getSiteMembersWithGroups", sig: "()", what: "Each member and the groups they belong to.",
        example: `await SPUtils.getSiteMembersWithGroups()` },
      { name: "getSiteGroupMembers", sig: "()", what: "Each group with its members, job title and department.",
        example: `await SPUtils.getSiteGroupMembers()` },
      { name: "addUsersFromCSVToGroups", sig: "(filePath, { dryRun, concurrency })", what: "Add users to groups from a CSV in a library (Email Address, Group columns).",
        example: `await SPUtils.addUsersFromCSVToGroups("/sites/X/Shared Documents/users.csv", { dryRun: true })` },
      { name: "syncUsersFromCSVToGroups", sig: "(filePath, { dryRun, adminEmail, umbrellaGroup })", what: "Make group membership match a master CSV: adds AND removes. dryRun defaults to true.",
        example: `await SPUtils.syncUsersFromCSVToGroups("/sites/X/Shared Documents/master.csv", { dryRun: true, adminEmail: "me@x.com" })` },
      { name: "getSharingLinkForItem", sig: "(filePath, canEdit = false, expireInDays = 1)", what: "Create a view or edit sharing link for a file (0 days = no expiry).",
        example: `await SPUtils.getSharingLinkForItem("/sites/X/Shared Documents/a.pdf", false, 7)` },
      { name: "removeAllSharingLinksForItem", sig: "(filePath, areYouSure)", what: "Revoke every sharing link on a file. Second argument must be true.",
        example: `await SPUtils.removeAllSharingLinksForItem("/sites/X/Shared Documents/a.pdf", true)` }
    ];

    // ---------------------------------------------------------------
    // Status panel — drawn by the script itself on first use, so the
    // hosting page (a DCSPad preview, a Script Editor web part, a bare
    // test page) needs no markup. If the page already has <p id="status">
    // (the legacy web part HTML), that element is reused as-is.
    // Optional: put data-sputils-host on any element to mount inside it.
    // ---------------------------------------------------------------
    const FONT = "font-family: Segoe UI, system-ui, sans-serif;";
    const MONO = "font-family: Consolas, ui-monospace, monospace;";
    let statusEl = null;
    let pendingLines = [];

    const escapeHtml = (v) => String(v)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

    const buildUsagePanel = () => {
      const rows = USAGE.map(u => `
        <tr>
          <td style="padding:2px 8px 2px 0; white-space:nowrap; vertical-align:top;"><code style="${MONO}">${escapeHtml(u.name)}${escapeHtml(u.sig)}</code></td>
          <td style="padding:2px 8px 2px 0; vertical-align:top;">${escapeHtml(u.what)}<br><code style="${MONO} color:#555;">${escapeHtml(u.example)}</code></td>
        </tr>`).join("");
      return `
        <details id="sputils-usage" style="margin:6px 0 8px;">
          <summary style="cursor:pointer; font-size:12px; color:#333;">Usage &mdash; ${USAGE.length} functions (or run <code style="${MONO}">SPUtils.help()</code>)</summary>
          <div style="font-size:12px; color:#333; margin-top:6px; line-height:1.4;">
            <p style="margin:0 0 6px;">Everything hangs off <code style="${MONO}">window.SPUtils</code>. Requires PnPjs v2 on the page as <code style="${MONO}">pnp2</code>.
            Call <code style="${MONO}">setupContext()</code> first. Most functions are async: prefix with <code style="${MONO}">await</code>.
            Anything that writes has a <code style="${MONO}">dryRun</code> option; use it first.</p>
            <table style="border-collapse:collapse; width:100%;">${rows}</table>
          </div>
        </details>`;
    };

    const ensureStatusEl = () => {
      if (statusEl && statusEl.isConnected) return statusEl;
      if (typeof document === "undefined" || !document.body) return null;

      const existing = document.getElementById("status");
      if (existing) { statusEl = existing; return statusEl; }

      const host = document.querySelector("[data-sputils-host]") || document.body;
      const box = document.createElement("div");
      box.id = "statusbox";
      box.setAttribute("style", `${FONT} margin:10px; padding:10px; background-color:#eee; color:#222; border-radius:4px;`);
      box.innerHTML = `
        <div style="display:flex; align-items:baseline; gap:10px;">
          <p style="color:#666; margin:0 0 3px; font-size:10px; letter-spacing:.05em;">SPUtils STATUS</p>
          <button type="button" id="sputils-clear" style="font-size:10px; padding:1px 6px; cursor:pointer;">Clear</button>
        </div>
        ${buildUsagePanel()}
        <div id="status" style="${MONO} font-size:12px; margin:0; max-height:320px; overflow:auto;"></div>`;
      host.appendChild(box);
      box.querySelector("#sputils-clear").addEventListener("click", () => clearStatus());
      statusEl = box.querySelector("#status");
      return statusEl;
    };

    const appendLine = (source, message) => {
      const el = ensureStatusEl();
      if (!el) { pendingLines.push([source, message]); return; }
      if (pendingLines.length) {
        const queued = pendingLines; pendingLines = [];
        queued.forEach(([s, m]) => appendLine(s, m));
      }
      // Message is inserted as HTML on purpose (callers pass simple markup); source is escaped.
      el.insertAdjacentHTML("beforeend", `<div><strong>${escapeHtml(source)}:</strong> ${message}</div>`);
      el.scrollTop = el.scrollHeight;
    };

    /**
     * writeStatus (Internal)
     * Writes a status message to the on-page status panel and the browser console.
     * The panel is created on first use; if the document body is not ready yet,
     * lines are buffered and flushed on DOMContentLoaded.
     *
     * @param {string} source - Function name or source identifier for the message
     * @param {string} message - Status message to display (simple HTML allowed)
     */
    const writeStatus = (source, message) => {
      appendLine(source, message);
      console.log(`${source}: ${message}`);
    };

    if (typeof document !== "undefined" && !document.body) {
      document.addEventListener("DOMContentLoaded", () => {
        const queued = pendingLines; pendingLines = [];
        queued.forEach(([s, m]) => appendLine(s, m));
      }, { once: true });
    }

    /**
     * clearStatus
     * Empties the on-page status log.
     */
    const clearStatus = () => {
      const el = ensureStatusEl();
      if (el) el.innerHTML = "";
    };

    /**
     * help
     * Prints the function list to the console (console.table) and opens the
     * usage panel on the page.
     *
     * @returns {Array<{name:string, sig:string, what:string, example:string}>}
     */
    const help = () => {
      console.table(USAGE.map(u => ({ Function: u.name + u.sig, What: u.what, Example: u.example })));
      ensureStatusEl();
      const panel = document.getElementById("sputils-usage");
      if (panel) { panel.open = true; panel.scrollIntoView({ block: "nearest" }); }
      writeStatus("help", `${USAGE.length} functions listed in the console. Start with <code>SPUtils.setupContext("sites/YourSite")</code>.`);
      return USAGE;
    };
    // Initial check to confirm PnPjs is available

    writeStatus("Init", `SPUtils loaded. ${lastModified.trim()}`);

    if (typeof pnp2 !== "undefined") {
      writeStatus("Init", "PnPjs (pnp2) loaded successfully.");
    } else {
      writeStatus("Init", "PnPjs (pnp2) not found.");
    }

    /**
     * setupContext
     * **MUST BE CALLED FIRST** — Initializes PnPjs context for a specific SharePoint site.
     * All subsequent list, group, and user operations use this context.
     * 
     * @param {string} siteSubUrl - Site path relative to the current tenant (e.g., "teams/Marketing"
     *                               or "sites/MyProject"), or an absolute same-tenant URL
     * @throws {Error} If context setup fails
     * @example
     * // Initialize for a team site
     * window.SPUtils.setupContext("teams/FCUCommunication");
     * // Initialize for a managed path site
     * window.SPUtils.setupContext("sites/ProjectX");
     */
    const setupContext = (siteSubUrl) => {
      try {
        const raw = String(siteSubUrl || "").trim();
        const baseUrl = (/^https?:\/\//i.test(raw) ? raw : SPTenantBase + raw.replace(/^\/+/, ""))
          .replace(/\/+$/, "");
        pnp2.sp.setup({ sp: { baseUrl } });
        writeStatus("setupContext", `Context set to: ${baseUrl}`);
        siteBaseUrl = baseUrl;
        contextInitialized = true;
      } catch (error) {
        writeStatus("setupContext", `Failed to set context: ${error.message}`);
        contextInitialized = false;
      }
    };
    
      /**
     * retryWithBackoff (Internal)
     * Retries an async operation with exponential backoff when throttled (429 Too Many Requests).
     * Implements retry logic: 500ms, 1s, 2s, 4s, 8s, 16s intervals (doubling each attempt).
     * Immediately re-throws non-429 errors or fails after maxRetries.
     * 
     * @param {Function} fn - Async function to execute
     * @param {number} maxRetries - Maximum retry attempts (default: 5)
     * @param {number} baseDelayMs - Initial delay in milliseconds (default: 500)
     * @returns {Promise<any>} Result from the async function on success
     * @throws {Error} If max retries exceeded or non-429 error encountered
     */
    const retryWithBackoff = async (fn, maxRetries = 5, baseDelayMs = 500) => {
      for (let attempt = 0; attempt <= maxRetries; attempt++) {
          try {
              return await fn();
          } catch (err) {
              const is429 = err?.status === 429 || (err?.message || "").includes("Too Many Requests");
              if (!is429 || attempt === maxRetries) {
                  throw err;
              }
              const delay = baseDelayMs * Math.pow(2, attempt);
              writeStatus("retryWithBackoff", `429 detected. Retrying in ${delay}ms (attempt ${attempt + 1}/${maxRetries})...`);
              await new Promise(res => setTimeout(res, delay));
          }
      }
    };


    /**
     * previewListItems
     * Fetches and displays all items from a SharePoint list in a paginated table.
     * Useful for auditing list contents and item counts before bulk operations.
     * Output: console.table() with Id and Title columns.
     * 
     * @param {string} listTitle - Display name of the list (e.g., "My Documents" or "Tasks")
     * @param {number} batchSize - Items per request (default: 100); adjust for large lists
     * @returns {Promise<void>} Logs table to console; does not return data
     * @example
     * // Preview all items in a list
     * await window.SPUtils.previewListItems("Announcements", 100);
     */
    const previewListItems = async (listTitle, batchSize = 100) => {
      let items = [];
      let skip = 0;
  
      writeStatus("previewListItems", `Fetching items from list: ${listTitle}`);
  
      while (true) {
        const page = await pnp2.sp.web.lists.getByTitle(listTitle).items
          .select("Id", "Title")
          .top(batchSize)
          .skip(skip)();
  
        if (page.length === 0) break;
  
        items = items.concat(page);
        skip += batchSize;
      }
  
      writeStatus("previewListItems", `Total items found: ${items.length}`);
      console.table(items.map(item => ({ Id: item.Id, Title: item.Title })));
    };
  
    
  /**
   * getListFields
   * Retrieves all visible fields for a SharePoint list with configuration details.
   *
   * @param {string} listTitle
   * Display name of the list
   *
   * @returns {Promise<Array>}
   * Returns array of field metadata objects
   *
   * @example
   * const fields = await window.SPUtils.getListFields("My List");
   * console.table(fields);
   */
  const getListFields = async (listTitle) => {

      const FN = "getListFields";

      if (!contextInitialized) {
          writeStatus(FN, "⚠️ Call setupContext() first.");
          throw new Error("Context not initialized");
      }

      writeStatus(FN, `Fetching fields for list: ${listTitle}`);

      try {

          const fields = await pnp2.sp.web.lists
              .getByTitle(listTitle)
              .fields
              .filter("Hidden eq false")
              .select(
                  "InternalName",
                  "Title",
                  "TypeAsString",
                  "Required",
                  "ReadOnlyField",
                  "Description",
                  "DefaultValue",
                  "Choices",
                  "MaxLength"
              )();

          writeStatus(FN, `Fields found: ${fields.length}`);

          console.table(fields.map(f => ({
              InternalName: f.InternalName,
              Title: f.Title,
              Type: f.TypeAsString,
              Required: f.Required,
              ReadOnly: f.ReadOnlyField
          })));

          return fields;

      } catch (err) {
          writeStatus(FN, `Error: ${err.message}`);
          console.error(err);
          throw err;
      }
  };
      

    /**
     * deleteListItems
     * Batch-deletes items from a SharePoint list with optional dry-run and selective filtering.
     * Uses PnPjs batch API for efficiency. Fetches all items, applies optional filter,
     * then deletes in configurable batch sizes (default 100 per request).
     * 
     * @param {string} listTitle - Display name of the list
     * @param {object} [options] - Optional configuration object:
     *   @param {number} [options.batchSize=100] - Items per batch request
     *   @param {boolean} [options.dryRun=false] - If true, logs what would be deleted (no action)
     *   @param {Function} [options.filterFn=null] - Filter predicate to select items for deletion
     *                                              Example: item => item.Title.includes("Archive")
     * @returns {Promise<void>} Logs results to console
     * @example
     * // Delete all items with "temp" in title (dry-run only)
     * await window.SPUtils.deleteListItems("MyList", {
     *   dryRun: true,
     *   filterFn: item => item.Title && item.Title.includes("temp")
     * });
     * 
     * @example
     * // Actually delete items created before Jan 1, 2024
     * await window.SPUtils.deleteListItems("Archive", {
     *   dryRun: false,
     *   filterFn: item => new Date(item.Created) < new Date("2024-01-01")
     * });
     */
    const deleteListItems = async (listTitle, options = {}) => {
      const {
        batchSize = 100,
        dryRun = false,
        filterFn = null
      } = options;
  
      let items = [];
      let skip = 0;
  
      writeStatus("deleteListItems", `Fetching items from list: ${listTitle}`);
  
      while (true) {
        const page = await pnp2.sp.web.lists.getByTitle(listTitle).items
          .select("Id", "Title")
          .top(batchSize)
          .skip(skip)();
  
        if (page.length === 0) break;
  
        items = items.concat(page);
        skip += batchSize;
      }
  
      writeStatus("deleteListItems", `Total items fetched: ${items.length}`);
  
      // Apply optional filter function
      if (typeof filterFn === "function") {
        items = items.filter(filterFn);
        writeStatus("deleteListItems", `Items after filtering: ${items.length}`);
      }
  
      // Dry run: log what would be deleted
      if (dryRun) {
        writeStatus("deleteListItems", "Dry run enabled — no items will be deleted.");
        console.table(items.map(item => ({ Id: item.Id, Title: item.Title })));
        return;
      }
  
      writeStatus("deleteListItems", "Starting deletion...");
  
      try {
        for (let i = 0; i < items.length; i += batchSize) {
          const batch = pnp2.sp.web.createBatch();
          const chunk = items.slice(i, i + batchSize);
  
          for (const item of chunk) {
            try {
              pnp2.sp.web.lists.getByTitle(listTitle).items.getById(item.Id).inBatch(batch).delete();
            } catch (err) {
              writeStatus("deleteListItems", `Skipped item ${item.Id}: ${err.message}`);
            }
          }
          await batch.execute();
          
        }
  
        writeStatus("deleteListItems", "Deletion finished.");
      } catch (err) {
        writeStatus("deleteListItems", `Error during deletion: ${err.message}`);
        console.error(err);
      }
    };
  

  /**
   * getAllLists
   * Retrieves and displays all lists on the current SharePoint site.
   * Output: console.table() with list Title and Id columns, useful for discovering list names.
   * 
   * @returns {Promise<void>} Logs table to console; total count in status element
   * @example
   * // List all available lists on the site
   * await window.SPUtils.getAllLists();
   */
  const getAllLists = async () => {
    writeStatus("getAllLists", "Fetching all lists from site...");
    try {
      const lists = await pnp2.sp.web.lists.select("Id", "Title")();
      writeStatus("getAllLists", `Total lists found: ${lists.length}`);
      console.table(lists.map(list => ({ Id: list.Id, Title: list.Title })));
    } catch (err) {
      writeStatus("getAllLists", `Error fetching lists: ${err.message}`);
      console.error(err);
    }
  };

  /**
 * getListSchemaForMigration
 *
 * Retrieves SharePoint list configuration metadata useful for recreating
 * a list structure on another site. Does NOT export list item data.
 *
 * Intended to be paired with a CSV export of the data itself.
 *
 * Captures:
 * - List settings
 * - Field definitions
 * - Views
 * - Content types
 *
 * @param {string} listTitle
 * Display name of the list.
 *
 * @returns {Promise<object>}
 *
 * @example
 * const schema = await SPUtils.getListSchemaForMigration("Incidents");
 * console.log(JSON.stringify(schema, null, 2));
 */
const getListSchemaForMigration = async (listTitle) => {
    const FN = "getListSchemaForMigration";

    if (!contextInitialized) {
        writeStatus(FN, "⚠️ Call setupContext() first.");
        throw new Error("Context not initialized");
    }

    writeStatus(FN, `Analyzing list "${listTitle}"...`);

    try {

        const list = pnp2.sp.web.lists.getByTitle(listTitle);

        // ---------------------------------------------------
        // LIST SETTINGS
        // ---------------------------------------------------

        const settings = await list
            .select(
                "Id",
                "Title",
                "Description",
                "BaseTemplate",
                "EnableVersioning",
                "EnableMinorVersions",
                "ForceCheckout",
                "Hidden",
                "ContentTypesEnabled",
                "EnableAttachments",
                "ItemCount"
            )();

        // ---------------------------------------------------
        // FIELDS
        // ---------------------------------------------------

        const fields = await list.fields
            .filter("Hidden eq false")
            .select(
                "InternalName",
                "Title",
                "TypeAsString",
                "Required",
                "ReadOnlyField",
                "Description",
                "DefaultValue",
                "Choices",
                "MaxLength",
                "SchemaXml"
            )();

        // ---------------------------------------------------
        // CONTENT TYPES
        // ---------------------------------------------------

        let contentTypes = [];

        try {
            contentTypes = await list.contentTypes
                .select(
                    "Name",
                    "StringId",
                    "Description",
                    "Group",
                    "Hidden",
                    "ReadOnly"
                )();
        }
        catch (err) {
            writeStatus(FN, `Content Type read failed: ${err.message}`);
        }

        // ---------------------------------------------------
        // VIEWS
        // ---------------------------------------------------

        let views = [];

        try {

            const rawViews = await list.views
                .select(
                    "Title",
                    "DefaultView",
                    "Paged",
                    "RowLimit",
                    "ViewQuery"
                )();

            views = [];

            for (const view of rawViews) {

                let fieldsForView = [];

                try {
                    const viewFields =
                        await list.views
                            .getByTitle(view.Title)
                            .fields();

                    fieldsForView = viewFields.Items || [];
                }
                catch {}

                views.push({
                    ...view,
                    Fields: fieldsForView
                });
            }
        }
        catch (err) {
            writeStatus(FN, `View read failed: ${err.message}`);
        }

        // ---------------------------------------------------
        // BUILD RESULT
        // ---------------------------------------------------

        const schema = {
            exported: new Date().toISOString(),
            list: {
                title: settings.Title,
                description: settings.Description,
                baseTemplate: settings.BaseTemplate,
                enableVersioning: settings.EnableVersioning,
                enableMinorVersions: settings.EnableMinorVersions,
                forceCheckout: settings.ForceCheckout,
                hidden: settings.Hidden,
                contentTypesEnabled: settings.ContentTypesEnabled,
                enableAttachments: settings.EnableAttachments
            },
            fields: fields.map(f => ({
                internalName: f.InternalName,
                displayName: f.Title,
                type: f.TypeAsString,
                required: f.Required,
                readOnly: f.ReadOnlyField,
                description: f.Description,
                defaultValue: f.DefaultValue,
                choices: f.Choices || [],
                maxLength: f.MaxLength,
                schemaXml: f.SchemaXml
            })),
            contentTypes,
            views
        };

        writeStatus(
            FN,
            `Completed. Fields=${schema.fields.length}, Views=${views.length}, ContentTypes=${contentTypes.length}`
        );

        console.log(schema);

        return schema;

    }
    catch (err) {
        writeStatus(FN, `Error: ${err.message}`);
        console.error(err);
        throw err;
    }
};

  /**
   * getAllSecurityGroups
   * Retrieves and displays all SharePoint security groups on the current site.
   * Output: console.table() with group Title and Id columns.
   * Useful for discovering group names before adding/removing members.
   * 
   * @returns {Promise<void>} Logs table to console; total count in status element
   * @example
   * // List all security groups available
   * await window.SPUtils.getAllSecurityGroups();
   */
  const getAllSecurityGroups = async () => {
    writeStatus("getAllSecurityGroups", "Fetching all security groups...");
    try {
      const groups = await pnp2.sp.web.siteGroups.select("Id", "Title")();
      writeStatus("getAllSecurityGroups", `Total groups found: ${groups.length}`);
      console.table(groups.map(group => ({ Id: group.Id, Title: group.Title })));
    } catch (err) {
      writeStatus("getAllSecurityGroups", `Error fetching groups: ${err.message}`);
      console.error(err);
    }
  };


    /**
   * getSiteMembersWithGroups
   * Retrieves and displays site members mapped to their group memberships (user → groups).
   * Output: console.table() with User, LoginName, and comma-separated Groups columns.
   * Useful for auditing who is in which groups and identifying orphaned members.
   * 
   * @returns {Promise<void>} Logs table to console; total unique members in status element
   * @example
   * // See which groups each site member belongs to
   * await window.SPUtils.getSiteMembersWithGroups();
   */
  const getSiteMembersWithGroups = async () => {
    writeStatus("getSiteMembersWithGroups", "Fetching site groups and members...");
    try {
      const groups = await pnp2.sp.web.siteGroups.select("Id", "Title")();
      const memberMap = {};

      for (const group of groups) {
        const users = await pnp2.sp.web.siteGroups.getById(group.Id).users.select("LoginName", "Title")();
        users.forEach(user => {
          if (!memberMap[user.LoginName]) {
            memberMap[user.LoginName] = {
              User: user.Title,
              LoginName: user.LoginName,
              Groups: []
            };
          }
          memberMap[user.LoginName].Groups.push(group.Title);
        });
      }

      const memberList = Object.values(memberMap).map(member => ({
        User: member.User,
        LoginName: member.LoginName,
        Groups: member.Groups.join(", ")
      }));

      writeStatus("getSiteMembersWithGroups", `Total unique members found: ${memberList.length}`);
      console.table(memberList);
    } catch (err) {
      writeStatus("getSiteMembersWithGroups", `Error fetching members: ${err.message}`);
      console.error(err);
    }
  };

  /**
 * getSharingLinkForDocumentOld (Deprecated)
 * Legacy function for getting sharing links by library and filename.
 * @deprecated Use getSharingLinkForItem() instead (requires full server-relative path)
 */
const getSharingLinkForDocumentOld = async (libraryTitle, fileName) => {
    writeStatus("getSharingLinkForDocument", `Searching for '${fileName}' in '${libraryTitle}'...`);
    try {
        const files = await pnp2.sp.web.lists.getByTitle(libraryTitle).items
            .select("Id", "FileLeafRef", "FileRef")
            .filter(`FileLeafRef eq '${fileName}'`)();

        if (files.length === 0) {
            writeStatus("getSharingLinkForDocument", `No file named '${fileName}' found.`);
            return;
        }

        if (files.length > 1) {
            writeStatus("getSharingLinkForDocument", `Multiple files found with name '${fileName}'.`);
            console.table(files.map(f => ({ Id: f.Id, FileRef: f.FileRef })));
        }

        const fileRef = files[0].FileRef;
        const link = await pnp2.sp.web.getFileByServerRelativeUrl(fileRef)
            .getShareLink("view", 7);

        writeStatus("getSharingLinkForDocument", `Sharing link for '${fileName}' generated.`);
        console.log(`Link: ${link}`);
    } catch (err) {
        writeStatus("getSharingLinkForDocument", `Error: ${err.message}`);
        console.error(err);
    }
};



/**
 * getSharingLinkForItem
 * Generates an expiring sharing link for a file using PnPjs.
 * Returns a read-only or editable link with optional expiration.
 * 
 * @param {string} filePath - Server-relative path to file (e.g., "/sites/MyTeam/Shared Documents/Report.pdf")
 * @param {boolean} [canEdit=false] - If true, generates edit link (3); if false, read-only link (2)
 * @param {number} [expireInDays=1] - Expiration in days from now (0 = no expiration)
 * @returns {Promise<string>} URL of the generated sharing link
 * @throws {Error} If file not found or link generation fails
 * @example
 * // Generate a read-only link that expires in 7 days
 * const link = await window.SPUtils.getSharingLinkForItem(
 *   "/sites/Marketing/Shared Documents/Q4-Report.docx",
 *   false,
 *   7
 * );
 * console.log(link);
 * 
 * @example
 * // Generate an editable link with no expiration
 * const editLink = await window.SPUtils.getSharingLinkForItem(
 *   "/sites/Marketing/Shared Documents/budget.xlsx",
 *   true,
 *   0
 * );
 */
const getSharingLinkForItem = async (filePath, canEdit = false, expireInDays = 1) => {
  try {
    const expirationDate = expireInDays ===0 ? null : new Date(Date.now() + expireInDays * 24 * 60 * 60 * 1000);
    writeStatus("getSharingLinkForItem", `Getting link for '${filePath}'...`);
    // Get the file object via PnPjs using the full server-relative path
    const file = pnp2.sp.web.getFileByServerRelativePath(filePath);

    // Generate a view link (2 = view, 1 = edit)
    const link = await file.getShareLink(canEdit ? 3 : 2, expirationDate); // 2 = view link
    return link.sharingLinkInfo.Url;

  } catch (err) {
    writeStatus("getSharingLinkForItem", `Error generating link: ${err}`);
    throw err;
  }
};

/**
 * removeAllSharingLinksForItem
 * Revokes ALL sharing links for a file (both internal and external recipients lose access).
 * Requires explicit areYouSure=true confirmation to protect against accidental revocation.
 * 
 * @param {string} filePath - Server-relative path to file (e.g., "/sites/MyTeam/Shared Documents/Report.pdf")
 * @param {boolean} [areYouSure=false] - MUST be true to execute; prevents accidental removals
 * @returns {Promise<boolean>} True if successful; throws on error
 * @throws {Error} If file not found or unshare operation fails
 * @example
 * // Remove all sharing for a sensitive document
 * await window.SPUtils.removeAllSharingLinksForItem(
 *   "/sites/HR/Shared Documents/Confidential.pdf",
 *   true  // Must explicitly confirm
 * );
 */
const removeAllSharingLinksForItem = async (filePath, areYouSure = false) => {
  if (areYouSure) {
    try {
      writeStatus("removeAllSharingLinksForItem", `Removing all sharing links for '${filePath}'...`);

      const file = pnp2.sp.web.getFileByServerRelativePath(filePath);

      // Remove all sharing links
      await file.unshare();
      writeStatus("removeAllSharingLinksForItem", `All sharing links removed for '${filePath}'.`);
      return true;
    } catch (err) {
      writeStatus("removeAllSharingLinksForItem", `Error removing sharing links: ${err}`);
      throw err;
    }
  }
};

  /**
   * getSiteGroupMembers
   * Displays all SharePoint groups with their member details (groups → users).
   * Fetches user profile properties (Title, Department) for each member.
   * Output: console.table() with Group, User, Email, Title, Department columns.
   * Skips system accounts and entries without valid email addresses.
   * 
   * @returns {Promise<void>} Logs table to console; sorted by group name; total entries in status
   * @example
   * // See all groups and their members with job titles
   * await window.SPUtils.getSiteGroupMembers();
   */
  const getSiteGroupMembers = async () => {
    writeStatus("getSiteGroupMembers", "Fetching site groups and their members...");
    try {
      const groups = await pnp2.sp.web.siteGroups.select("Id", "Title")();
      const output = [];

      for (const group of groups) {
        const users = await pnp2.sp.web.siteGroups.getById(group.Id).users.select("Title", "LoginName", "Email")();

        for (const user of users) {
          const rawLogin = user.LoginName;
          const emailMatch = rawLogin.match(/[^|]+@[^|]+$/);
          const email = emailMatch ? emailMatch[0] : "";

          // Skip system accounts or entries without valid email
          if (!email || email.toLowerCase().includes("sharepoint") || email.toLowerCase().includes("system")) {
            continue;
          }

          let title = "";
          let department = "";

          try {
            title = await pnp2.sp.profiles.getUserProfilePropertyFor(rawLogin, "Title") || "";
            department = await pnp2.sp.profiles.getUserProfilePropertyFor(rawLogin, "Department") || "";
          } catch (profileErr) {
            writeStatus("getSiteGroupMembers", `Could not fetch profile for ${rawLogin}`);
          }

          output.push({
            Group: group.Title,
            User: user.Title,
            Email: email,
            Title: title,
            Department: department
          });
        }
      }

      writeStatus("getSiteGroupMembers", `Total group-user entries: ${output.length}`);
      output.sort((a, b) => a.Group.localeCompare(b.Group));
      console.table(output);
    } catch (err) {
      writeStatus("getSiteGroupMembers", `Error fetching group members: ${err.message}`);
      console.error(err);
    }
  };

/**
 * addUsersFromCSVToGroups
 * Adds users listed in a CSV to SharePoint groups.
 * Accepts either:
 *  - Two-column rows: Email/UPN, Group (no header), OR
 *  - A header row with any/all of: Name, Email Address, Title, Department, Manager, Group
 *    (Only "Email Address" and "Group" are used; others are ignored.)
 *
 * Other behaviors:
 *  - Lines beginning with '#' are ignored as comments.
 *  - Quotes are supported around values. (Escaped quotes "" inside quoted fields are handled.)
 *  - File must live in a document library on the current site/tenant.
 *
 * @param {string} filePath  Server-relative CSV path (e.g., "/sites/foo/Shared Documents/users_and_groups.csv")
 * @param {object} [options]
 * @param {boolean} [options.dryRun=false]    If true, preview only (no adds)
 * @param {number}  [options.concurrency=6]   Parallel adds
 * @returns {Promise<{attempted:number, added:number, skippedExisting:number, failed:Array<{user:string,group:string,error:string}>, perGroup:Object, dryRun:boolean}>}
 */
const addUsersFromCSVToGroups = async (filePath, { dryRun = false, concurrency = 6 } = {}) => {
  const FN = "addUsersFromCSVToGroups";
  try {
    if (!contextInitialized) {
      writeStatus(FN, "⚠️ Call setupContext() first to initialize PnPjs context.");
      throw new Error("Context not initialized");
    }
    if (!filePath) throw new Error("filePath is required.");

    // --- Helpers ---
    const stripBOM = (s) => (s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s);
    const toClaimsFromUPN = (upn) => `i:0#.f|membership|${(upn || "").toLowerCase()}`;
    const normHeader = (s) => (s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); // remove spaces/punct
    const unquote = (s) => s.replace(/^"|"$/g, "");
    // Split a single CSV line into fields, respecting quotes and escaped "" inside quotes
    const splitCsvLine = (line) => {
      const out = [];
      let cur = "";
      let inQ = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          // Handle escaped double-quote inside a quoted field
          if (inQ && line[i + 1] === '"') { cur += '"'; i++; continue; }
          inQ = !inQ;
          continue;
        }
        if (ch === "," && !inQ) { out.push(cur); cur = ""; continue; }
        cur += ch;
      }
      out.push(cur);
      return out.map(v => unquote(v.trim()));
    };

    // --- 1) Read CSV text (PnPjs respects auth/baseUrl) ---
    writeStatus(FN, `Reading CSV from ${filePath}...`);
    let csv = await pnp2.sp.web.getFileByServerRelativePath(filePath).getText();
    csv = stripBOM(csv);

    // --- 2) Preprocess lines: trim, drop empties, ignore comments ---
    const rawLines = csv.split(/\r?\n/).map(l => l.trim()).filter(Boolean).filter(l => !l.startsWith("#"));
    if (!rawLines.length) {
      writeStatus(FN, "Nothing to process.");
      return { attempted: 0, added: 0, skippedExisting: 0, failed: [], perGroup: {}, dryRun };
    }

    // --- 3) Parse first line; detect header (Email Address / Group) ---
    const firstFields = splitCsvLine(rawLines[0]);
    const headerCandidates = firstFields.map(normHeader);

    // We recognize headers like:
    //   "name", "emailaddress", "title", "department", "manager", "group" (case/spacing-insensitive)
    const isHeader =
      headerCandidates.includes("emailaddress") && headerCandidates.includes("group");

    let emailIdx = -1, groupIdx = -1;
    let dataLines = rawLines;

    if (isHeader) {
      // Find the indices we care about; ignore everything else
      emailIdx = headerCandidates.findIndex(h => h === "emailaddress" || h === "email" || h === "upn");
      groupIdx = headerCandidates.findIndex(h => h === "group" || h === "groupname");
      dataLines = rawLines.slice(1);
      writeStatus(FN, "Header detected. Using 'Email Address' and 'Group' columns.");
    } else {
      // No header: treat as two-column CSV (email, group)
      emailIdx = 0;
      groupIdx = 1;
      writeStatus(FN, "No header detected. Assuming two columns: Email, Group.");
    }

    // --- 4) Build work items (email + group), ignore other columns ---
    const rows = [];
    for (const line of dataLines) {
      const cols = splitCsvLine(line);
      if (cols.length === 0) continue;
      const email = (cols[emailIdx] || "").toLowerCase().trim();
      const group = (cols[groupIdx] || "").trim();
      if (email && group) rows.push({ email, group });
    }

    if (!rows.length) {
      writeStatus(FN, "No valid rows found (need Email Address and Group).");
      return { attempted: 0, added: 0, skippedExisting: 0, failed: [], perGroup: {}, dryRun };
    }

    // De-duplicate (email + group pair)
    const key = (r) => `${r.email}||${r.group.toLowerCase()}`;
    const uniq = new Map();
    for (const r of rows) if (!uniq.has(key(r))) uniq.set(key(r), r);
    const workItems = Array.from(uniq.values());
    writeStatus(FN, `Prepared ${workItems.length} unique (email, group) assignment(s).`);

    // --- 5) Resolve groups & prefetch existing members per group ---
    const groupNames = Array.from(new Set(workItems.map(w => w.group)));
    writeStatus(FN, `Resolving ${groupNames.length} group(s) and fetching existing members...`);

    const groupCaches = new Map(); // name -> { api, existingEmails:Set, existingLogins:Set }
    const groupErrors = new Map(); // name -> "not found" etc.

    for (const gName of groupNames) {
      try {
        const api = pnp2.sp.web.siteGroups.getByName(gName);
        const members = await api.users.select("LoginName", "Email")();
        const existingEmails = new Set(members.map(m => (m.Email || "").toLowerCase()).filter(Boolean));
        const existingLogins = new Set(members.map(m => (m.LoginName || "").toLowerCase()).filter(Boolean));
        groupCaches.set(gName, { api, existingEmails, existingLogins });
      } catch (e) {
        const msg = `Group not found or inaccessible: '${gName}'`;
        groupErrors.set(gName, msg);
        writeStatus(FN, `⚠️ ${msg}`);
      }
    }

    const validWork = workItems.filter(w => groupCaches.has(w.group));
    const invalidWork = workItems.filter(w => !groupCaches.has(w.group));

    // Per-group breakdown
    const perGroup = Object.fromEntries(groupNames.map(n => [n, { attempted: 0, added: 0, skippedExisting: 0, failed: 0 }]));

    // --- 6) DRY RUN path ---
    if (dryRun) {
      const preview = [];
      for (const w of validWork) {
        const cache = groupCaches.get(w.group);
        const claims = toClaimsFromUPN(w.email);
        const exists = cache.existingEmails.has(w.email) || cache.existingLogins.has(claims) || cache.existingLogins.has(w.email);
        if (!exists) preview.push({ Email: w.email, Group: w.group });
        perGroup[w.group].attempted += exists ? 0 : 1;
        perGroup[w.group].skippedExisting += exists ? 1 : 0;
      }

      writeStatus(FN, `DRY RUN: ${preview.length} add(s) would be attempted. ${invalidWork.length} row(s) reference missing groups.`);
      if (preview.length) console.table(preview);
      if (invalidWork.length) {
        console.table(invalidWork.map(w => ({ Email: w.email, Group: w.group, Error: groupErrors.get(w.group) || "Group not found" })));
      }

      const attempted = preview.length;
      const skippedExisting = validWork.length - preview.length;
      const failed = invalidWork.map(w => ({ user: w.email, group: w.group, error: groupErrors.get(w.group) || "Group not found" }));
      if (failed.length) console.table(failed.map(f => ({ Email: f.user, Group: f.group, Error: f.error })));

      return { attempted, added: 0, skippedExisting, failed, perGroup, dryRun };
    }

    // --- 7) Execute adds with modest concurrency ---
    const results = { attempted: 0, added: 0, skippedExisting: 0, failed: [] };
    const queue = [];

    // Build queue only for those not already in the group (by email or claims/login)
    for (const w of validWork) {
      const cache = groupCaches.get(w.group);
      const claims = toClaimsFromUPN(w.email);
      const exists = cache.existingEmails.has(w.email) || cache.existingLogins.has(claims) || cache.existingLogins.has(w.email);
      if (!exists) queue.push(w);
      else { results.skippedExisting++; perGroup[w.group].skippedExisting++; }
    }
    results.attempted = queue.length;

    // Per-group attempted counts
    const attemptedByGroup = new Map(groupNames.map(n => [n, 0]));
    for (const w of queue) attemptedByGroup.set(w.group, (attemptedByGroup.get(w.group) || 0) + 1);
    for (const [g, count] of attemptedByGroup.entries()) if (perGroup[g]) perGroup[g].attempted += count;

    // Invalid work → failures
    for (const w of invalidWork) {
      results.failed.push({ user: w.email, group: w.group, error: groupErrors.get(w.group) || "Group not found" });
      perGroup[w.group] = perGroup[w.group] || { attempted: 0, added: 0, skippedExisting: 0, failed: 0 };
      perGroup[w.group].failed++;
    }

    const workers = new Array(Math.max(1, Math.min(concurrency, queue.length))).fill(0).map(async () => {
      while (queue.length) {
        const w = queue.shift();
        const cache = groupCaches.get(w.group);
        const { api, existingLogins, existingEmails } = cache;
        const claims = toClaimsFromUPN(w.email);
        try {
          // Fast path: add by claims login
          try {
            await api.users.add(claims);
          } catch {
            // Second chance: try raw UPN as login
            try {
              await api.users.add(w.email);
            } catch {
              // Last resort: ensureUser then add by returned LoginName
              const ensured = await pnp2.sp.web.ensureUser(w.email);
              await api.users.add(ensured.data.LoginName);
            }
          }
          results.added++;
          perGroup[w.group].added++;
          existingLogins.add(claims);
          existingEmails.add(w.email);
          writeStatus(FN, `✔️ Added ${w.email} → ${w.group}`);
        } catch (err) {
          const msg = err?.message || String(err);
          results.failed.push({ user: w.email, group: w.group, error: msg });
          perGroup[w.group].failed++;
          writeStatus(FN, `⚠️ Failed to add ${w.email} → ${w.group}: ${msg}`);
        }
      }
    });

    await Promise.all(workers);

    writeStatus(FN, `Finished. Added=${results.added}, SkippedExisting=${results.skippedExisting}, Failed=${results.failed.length}.`);

    // Log a table of users with errors (as requested)
    if (results.failed.length) {
      console.table(results.failed.map(f => ({ Email: f.user, Group: f.group, Error: f.error })));
    }

    return { ...results, perGroup, dryRun };

  } catch (err) {
    writeStatus("addUsersFromCSVToGroups", `Error: ${err.message}`);
    console.error(err);
    throw err;
  }
};

/**
 * addFieldsToList
 *
 * Creates SharePoint list columns from a simple JSON specification.
 * Designed for admin / migration / schema automation scenarios.
 *
 * Behavior:
 * - Resolves list via server-relative URL (safe across sites)
 * - Skips existing fields (idempotent)
 * - Supports dry-run preview mode
 * - Uses retryWithBackoff() for throttling protection
 *
 * Supported Field Types:
 * - Text (single line)
 * - Note / Multiline
 * - Boolean / Yes-No
 * - Choice (requires choices[])
 * - URL
 * - User / Person (single)
 *
 * @param {string} listServerRelativeUrl
 * Server-relative path to the list
 * Example: "/teams/MySite/Lists/MyList"
 *
 * @param {Array<Object>} specs
 * Array of field definitions:
 * [
 *   { displayName: "Title", type: "text" },
 *   { displayName: "Priority", type: "choice", choices: ["Low","Medium","High"] }
 * ]
 *
 * @param {object} [options]
 * @param {boolean} [options.dryRun=false]
 * If true, logs planned changes without creating fields
 *
 * @returns {Promise<{added:number, skipped:number, failed:Array}>}
 *
 * @example
 * await SPUtils.addFieldsToList(
 *   "/teams/MySite/Lists/Requests",
 *   [
 *     { displayName: "Request ID", type: "text" },
 *     { displayName: "Details", type: "note" },
 *     { displayName: "Priority", type: "choice", choices: ["Low","High"] },
 *     { displayName: "Owner", type: "user" }
 *   ],
 *   { dryRun: true }
 * );
 */

const addFieldsToList = async (listServerRelativeUrl, specs, { dryRun = false } = {}) => {
  const FN = "addFieldsToList";

  if (!contextInitialized) {
    writeStatus(FN, "⚠️ Call setupContext() first.");
    throw new Error("Context not initialized");
  }

  if (!listServerRelativeUrl) throw new Error("listServerRelativeUrl is required.");
  if (!Array.isArray(specs) || !specs.length) throw new Error("specs[] is required.");

  const esc = s => String(s).replace(/'/g, "''");

  // Resolve list safely (no GetList(@list) issues)
  const listArr = await pnp2.sp.web.lists
    .select("Id", "Title", "RootFolder/ServerRelativeUrl")
    .expand("RootFolder")
    .filter(`RootFolder/ServerRelativeUrl eq '${esc(listServerRelativeUrl)}'`)
    .top(1)();

  if (!listArr.length) throw new Error(`List not found: ${listServerRelativeUrl}`);
  const list = pnp2.sp.web.lists.getById(listArr[0].Id);

  const fieldExists = async (internalName) => {
    const res = await list.fields
      .select("Id")
      .filter(`InternalName eq '${esc(internalName)}'`)
      .top(1)();
    return res.length > 0;
  };

  const toInternalName = name =>
    String(name)
      .normalize("NFKD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^A-Za-z0-9]/g, "")
      .replace(/^[0-9]/, "_$&") || "Field";

  const results = { added: 0, skipped: 0, failed: [] };

  for (const spec of specs) {
    try {
      const displayName = spec.displayName || spec.name || "Field";
      const type = (spec.type || "Text").toLowerCase();
      const internalName = spec.name || toInternalName(displayName);

      if (await fieldExists(internalName)) {
        writeStatus(FN, `SKIP (exists): ${internalName}`);
        results.skipped++;
        continue;
      }

      if (dryRun) {
        writeStatus(FN, `DRY RUN: ${type} → ${displayName}`);
        continue;
      }

      let xml = "";

      switch (type) {
        case "text":
          xml = `<Field Type="Text" Name="${internalName}" DisplayName="${displayName}" />`;
          break;

        case "note":
        case "multiline":
          xml = `<Field Type="Note" Name="${internalName}" DisplayName="${displayName}" NumLines="6" />`;
          break;

        case "boolean":
        case "yes/no":
          xml = `<Field Type="Boolean" Name="${internalName}" DisplayName="${displayName}" />`;
          break;

        case "choice":
          const choices = (spec.choices || []).map(c => `<CHOICE>${c}</CHOICE>`).join("");
          xml = `<Field Type="Choice" Name="${internalName}" DisplayName="${displayName}"><CHOICES>${choices}</CHOICES></Field>`;
          break;

        case "url":
          xml = `<Field Type="URL" Name="${internalName}" DisplayName="${displayName}" />`;
          break;

        case "user":
        case "person":
          xml = `<Field Type="User" Name="${internalName}" DisplayName="${displayName}" UserSelectionMode="PeopleOnly" />`;
          break;

        default:
          throw new Error(`Unsupported type: ${type}`);
      }

      await retryWithBackoff(() => list.fields.createFieldAsXml(xml));

      writeStatus(FN, `ADDED: ${internalName}`);
      results.added++;

    } catch (err) {
      results.failed.push({
        name: spec.displayName || spec.name,
        error: err.message
      });

      writeStatus(FN, `⚠️ Failed: ${spec.displayName || spec.name}`);
    }
  }

  writeStatus(FN, `Done. Added=${results.added}, Skipped=${results.skipped}`);
  return results;
}; 

/**
 * exportListToExcel
 *
 * Exports ALL items from a SharePoint list to an Excel (.xlsx) file.
 * Automatically includes all fields present in the list.
 *
 * Behavior:
 * - Fetches list field definitions dynamically
 * - Retrieves all items (paged)
 * - Normalizes complex field types (user, lookup, etc.)
 * - Generates and downloads Excel file via SheetJS (XLSX)
 *
 * @param {string} listTitle
 * Display name of the list
 *
 * @param {object} [options]
 * @param {string} [options.fileName]
 * Optional output file name (default = listTitle + timestamp)
 *
 * @returns {Promise<void>}
 *
 * @example
 * await SPUtils.exportListToExcel("Requests");
 */

const exportListToExcel = async (listTitle, { fileName } = {}) => {
  const FN = "exportListToExcel";

  if (!contextInitialized) {
    writeStatus(FN, "⚠️ Call setupContext() first.");
    throw new Error("Context not initialized");
  }

  if (typeof ExcelJS === "undefined") {
    throw new Error("ExcelJS not found on page.");
  }

  writeStatus(FN, `Fetching fields...`);

  let fields = await pnp2.sp.web.lists
    .getByTitle(listTitle)
    .fields
    .filter("Hidden eq false")
    .select("InternalName", "Title", "TypeAsString", "Required", "ReadOnlyField")();

    
 fields = fields.filter(f => f.InternalName !== "ID");


  fields.unshift({
      InternalName: "ID",
      Title: "ID",
      TypeAsString: "Counter",
      Required: true,
      ReadOnlyField: true
  });


  const fieldNames = fields.map(f => f.InternalName);

writeStatus(FN, "Fetching items (no $select)...");

let allItems = [];
let nextLink = null;

do {
  const url = nextLink ||
    `${siteBaseUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$top=200`;

  const res = await fetch(url, {
    headers: {
      Accept: "application/json;odata=nometadata"
    },
    credentials: "same-origin"
  });

  const json = await res.json();

  allItems = allItems.concat(json.value);

  nextLink = json['@odata.nextLink'];

} while (nextLink);

allItems.sort((a, b) => a.ID - b.ID);

writeStatus(FN, `Items retrieved: ${allItems.length}`);

const normalize = (val) => {
  if (val == null) return "";

  // multi-value objects (lookup/user arrays)
  if (Array.isArray(val)) {
    return val.map(v =>
      v?.Title ||
      v?.LookupValue ||
      v?.Email ||
      ""
    ).join("; ");
  }

  // SharePoint returns objects as nested metadata sometimes
  if (typeof val === "object") {
    return (
      val.Title ||
      val.LookupValue ||
      val.Email ||
      val.Description ||
      val.Url ||
      JSON.stringify(val)
    );
  }

  return val;
};

  writeStatus(FN, "Building workbook...");

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(listTitle);
  const metaSheet = workbook.addWorksheet("Field Metadata");

  // header row
  worksheet.columns = fields.map(f => ({
    header: `${f.Title} [${f.TypeAsString}]`,
    key: f.InternalName,
    width: 30
  }));

  metaSheet.columns = [
    { header: "Display Name", key: "title", width: 30 },
    { header: "Internal Name", key: "internal", width: 30 },
    { header: "Type", key: "type", width: 20 },
    { header: "Required", key: "required", width: 12 },
    { header: "Read Only", key: "readonly", width: 12 }
  ];

  fields.forEach(f => {
    metaSheet.addRow({
      title: f.Title,
      internal: f.InternalName,
      type: f.TypeAsString,
      required: f.Required ? "Yes" : "No",
      readonly: f.ReadOnlyField ? "Yes" : "No"
    });
  });

  // rows
  for (const item of allItems) {
    const row = {};
    for (const f of fieldNames) {
      row[f] = normalize(item[f]);
    }
    worksheet.addRow(row);
  }

  const finalName =
    fileName ||
    `${listTitle.replace(/\s+/g, "_")}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  writeStatus(FN, "Downloading file...");

  const buffer = await workbook.xlsx.writeBuffer();

  const blob = new Blob([buffer], {
    type:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = finalName;
  link.click();

  writeStatus(FN, `✅ Export complete: ${finalName}`);
};


/**
* syncUsersFromCSVToGroups
* Reads a master CSV and synchronizes SharePoint group membership:
*  - Adds users in CSV who are not yet in their Group
*  - Removes users found in the Group who are not present in the CSV for that Group
*  - Honors special "ALL" rows: users in ALL are added to every group and never removed
*  - Requires an adminEmail: always added, never removed (all groups + umbrella)
*  - Optional umbrellaGroup: synced to the union of all groups (incl. ALL + admin)
*
* CSV header supported (only "Email Address" and "Group" are used; others ignored):
*   Name, Email Address, Title, Department, Manager, Group
*
* Also supports no header with two columns: Email, Group
* - Comma-delimited ONLY (quotes supported; lines starting with '#' are ignored)
*
* @param {string} filePath  Server-relative path to the CSV (e.g. "/sites/foo/Shared Documents/master.csv")
* @param {object} [options]
* @param {boolean} [options.dryRun=true]        Preview only; no changes are made
* @param {number}  [options.concurrency=6]       Parallelism for add/remove operations inside each group
* @param {boolean} [options.excludeSystem=true]  Skip removals for obvious system/service accounts
* @param {string}  [options.adminEmail]          REQUIRED. Email/UPN of admin user (always added; never removed)
* @param {string}  [options.umbrellaGroup]       OPTIONAL. Group to be synced to the union of all members
* @returns {Promise<{
*   dryRun:boolean,
*   groups:number,
*   totals:{ addPlanned:number, added:number, removePlanned:number, removed:number, skippedExisting:number },
*   perGroup:Record<string,{ addPlanned:number, added:number, removePlanned:number, removed:number, skippedExisting:number, errors:number }>,
*   failed:Array<{ email:string, group:string, action:'add'|'remove'|'parse', error:string }>,
*   // add-only style mirrors:
*   attempted:number, added:number, skippedExisting:number
* }>}
*/
const syncUsersFromCSVToGroups = async (filePath, {
 dryRun = true,
 concurrency = 6,
 excludeSystem = true,
 adminEmail,
 umbrellaGroup
} = {}) => {
 const FN = "syncUsersFromCSVToGroups";

 // ---------- helpers ----------
 const stripBOM = (s) => (s && s.charCodeAt(0) === 0xFEFF ? s.slice(1) : s);
 const normHeader = (s) => (s || "").toLowerCase().replace(/[^a-z0-9]/g, ""); // case/spacing/punct-insensitive
 const toClaims = (upn) => `i:0#.f|membership|${String(upn || "").toLowerCase()}`;
 const looksSystem = (email, login) => {
   if (!excludeSystem) return false;
   const e = String(email || "").toLowerCase();
   const l = String(login || "").toLowerCase();
   return e.includes("sharepoint") || e.includes("system") || l.includes("sharepoint") || l.includes("system");
 };
 const isAllGroup = (g) => String(g || "").trim().toLowerCase() === "all";

 // CSV splitter (comma-delimited; quote-aware; supports escaped "" inside quotes)
 const splitCommaLine = (line) => {
   const out = [];
   let cur = "";
   let inQ = false;
   for (let i = 0; i < line.length; i++) {
     const ch = line[i];
     if (ch === '"') {
       if (inQ && line[i + 1] === '"') { cur += '"'; i++; continue; } // escaped quote
       inQ = !inQ; continue;
     }
     if (ch === "," && !inQ) { out.push(cur); cur = ""; continue; }
     cur += ch;
   }
   out.push(cur);
   return out.map(v => v.trim().replace(/^"|"$/g, ""));
 };

 // small promise pool
 const runPool = async (items, limit, worker) => {
   const q = [...items];
   const n = Math.max(1, Math.min(limit, q.length || 1));
   const workers = new Array(n).fill(0).map(async () => {
     while (q.length) { const it = q.shift(); await worker(it); }
   });
   await Promise.all(workers);
 };

 try {
   if (!contextInitialized) {
     writeStatus(FN, "⚠️ Call setupContext() first to initialize PnPjs context.");
     throw new Error("Context not initialized");
   }
   if (!filePath) throw new Error("filePath is required.");
   const admin = String(adminEmail || "").trim().toLowerCase();
   if (!admin) {
     throw new Error("adminEmail is required (always added; never removed).");
   }

   // 1) Read CSV (via PnPjs to honor auth/context)
   writeStatus(FN, `Reading CSV from ${filePath}...`);
   let text = await pnp2.sp.web.getFileByServerRelativePath(filePath).getText();
   text = stripBOM(text);

   // 2) Preprocess lines: trim, drop empties, ignore comments
   const rawLines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean).filter(l => !l.startsWith("#"));
   if (!rawLines.length) {
     writeStatus(FN, "Nothing to process.");
     return {
       dryRun,
       groups: 0,
       totals: { addPlanned:0, added:0, removePlanned:0, removed:0, skippedExisting:0 },
       perGroup:{},
       failed:[],
       attempted: 0, added: 0, skippedExisting: 0
     };
   }

   // 3) Detect header (comma-delimited only)
   const firstFields = splitCommaLine(rawLines[0]);
   const headerKeys = firstFields.map(normHeader);
   const headerDetected = headerKeys.includes("emailaddress") && headerKeys.includes("group");

   let emailIdx = 0, groupIdx = 1, dataLines = rawLines;
   if (headerDetected) {
     emailIdx = headerKeys.findIndex(h => h === "emailaddress" || h === "email" || h === "upn");
     groupIdx = headerKeys.findIndex(h => h === "group" || h === "groupname");
     dataLines = rawLines.slice(1);
     writeStatus(FN, "Header detected. Using 'Email Address' and 'Group' columns; others ignored.");
   } else {
     writeStatus(FN, "No recognized header. Assuming two columns: Email, Group.");
   }

   // 4) Parse rows → separate desired groups and ALL rows
   const parsed = [];
   for (const line of dataLines) {
     const cols = splitCommaLine(line);
     if (!cols.length) continue;
     const email = (cols[emailIdx] || "").toLowerCase().trim();
     const group = (cols[groupIdx] || "").trim();
     if (email && group) parsed.push({ email, group });
   }
   if (!parsed.length) {
     writeStatus(FN, "No valid rows (need Email Address and Group).");
     return {
       dryRun,
       groups: 0,
       totals: { addPlanned:0, added:0, removePlanned:0, removed:0, skippedExisting:0 },
       perGroup:{},
       failed:[],
       attempted: 0, added: 0, skippedExisting: 0
     };
   }

   // 5) Build desired membership per group and collect ALL users
   const key = (r) => `${r.email}||${r.group.toLowerCase()}`;
   const uniq = new Map();
   for (const r of parsed) if (!uniq.has(key(r))) uniq.set(key(r), r);
   const rows = Array.from(uniq.values());

   const allUsers = new Set();           // users listed under "ALL"
   const desiredByGroup = new Map();     // group -> Set<email> (excluding "ALL")

   for (const { email, group } of rows) {
     if (isAllGroup(group)) {
       allUsers.add(email);
     } else {
       if (!desiredByGroup.has(group)) desiredByGroup.set(group, new Set());
       desiredByGroup.get(group).add(email);
     }
   }

   // Protect the admin and treat as if also in ALL
   allUsers.add(admin);

   // For each real group, inject ALL users + admin into desired set
   for (const [groupName, set] of desiredByGroup.entries()) {
     for (const a of allUsers) set.add(a);
   }

   // 6) Resolve groups & fetch current members; plan adds/removes
   const perGroup = {};
   const failed = [];
   let addPlanned = 0, removePlanned = 0, added = 0, removed = 0, skippedExisting = 0;

   const processOneGroup = async (groupName, desiredEmailsSet, { abortIfMissing = false } = {}) => {
     perGroup[groupName] = perGroup[groupName] || { addPlanned: 0, added: 0, removePlanned: 0, removed: 0, skippedExisting: 0, errors: 0 };

     let groupApi, members;
     try {
       groupApi = pnp2.sp.web.siteGroups.getByName(groupName);
       members = await groupApi.users.select("LoginName", "Email", "Title")(); // throws if missing
     } catch (e) {
       const msg = `Group not found or inaccessible: '${groupName}'`;
       writeStatus(FN, `⚠️ ${msg}`);
       failed.push({ email: "", group: groupName, action: "parse", error: msg });
       perGroup[groupName].errors++;
       if (abortIfMissing) {
         throw new Error(msg);
       }
       return { plannedAdds:0, plannedRemoves:0 };
     }

     // Build fast lookup of existing membership
     const existingEmails = new Set(members.map(m => (m.Email || "").toLowerCase()).filter(Boolean));
     const existingLogins = new Set(members.map(m => (m.LoginName || "").toLowerCase()).filter(Boolean));

     // Determine ADDs
     const toAdd = [];
     for (const email of desiredEmailsSet) {
       const claims = toClaims(email);
       const already = existingEmails.has(email) || existingLogins.has(claims) || existingLogins.has(email);
       if (already) { skippedExisting++; perGroup[groupName].skippedExisting++; continue; }
       toAdd.push(email);
     }
     perGroup[groupName].addPlanned += toAdd.length;
     addPlanned += toAdd.length;

     // Determine REMOVEs: present in group but not in desired, but NEVER remove ALL/admin; also skip system if configured
     const protectEmails = allUsers;                    // ALL + admin already in this set
     const protectClaims = new Set(Array.from(protectEmails).map(toClaims));

     const toRemoveMembers = members.filter(m => {
       const e = (m.Email || "").toLowerCase();
       const l = (m.LoginName || "").toLowerCase();
       if (looksSystem(e, l)) return false;            // optional system skip
       if (protectEmails.has(e) || protectClaims.has(l) || protectEmails.has(l)) return false; // NEVER remove ALL/admin
       // remove if not desired by this group
       const inDesired = (e && desiredEmailsSet.has(e)) || desiredEmailsSet.has(l) || desiredEmailsSet.has(l.replace(/^i:0#\.f\|membership\|/, ""));
       return !inDesired;
     });

     perGroup[groupName].removePlanned += toRemoveMembers.length;
     removePlanned += toRemoveMembers.length;

     // DRY RUN: show plan
     if (dryRun) {
       if (toAdd.length) {
         writeStatus(FN, `DRY RUN: ${groupName} → add ${toAdd.length}`);
         console.table(toAdd.map(e => ({ Action: "ADD", Group: groupName, Email: e })));
       }
       if (toRemoveMembers.length) {
         writeStatus(FN, `DRY RUN: ${groupName} → remove ${toRemoveMembers.length}`);
         console.table(toRemoveMembers.map(m => ({ Action: "REMOVE", Group: groupName, Login: m.LoginName, Email: m.Email || "" })));
       }
       return { plannedAdds: toAdd.length, plannedRemoves: toRemoveMembers.length };
     }

     // Execute adds (claims → upn → ensureUser)
     await runPool(toAdd, concurrency, async (email) => {
       const claims = toClaims(email);
       try {
         try {
           
            await retryWithBackoff(() => groupApi.users.add(claims));

         } catch {
           try {
            await retryWithBackoff(() => groupApi.users.add(email)); 
           } catch {
            const ensured = await retryWithBackoff(() => pnp2.sp.web.ensureUser(email));
            await retryWithBackoff(() => groupApi.users.add(ensured.data.LoginName));
           }
         }
         added++; perGroup[groupName].added++;
         existingEmails.add(email);
         existingLogins.add(claims);
         writeStatus(FN, `✔️ ADDED ${email} → ${groupName}`);
       } catch (err) {
         const msg = err?.message || String(err);
         failed.push({ email, group: groupName, action: "add", error: msg });
         perGroup[groupName].errors++;
         writeStatus(FN, `⚠️ Failed to add ${email} → ${groupName}: ${msg}`);
       }
     });

     // Execute removals (by LoginName only)
     await runPool(toRemoveMembers, concurrency, async (m) => {
       const login = m.LoginName;
       try {
         await groupApi.users.removeByLoginName(login);
         removed++; perGroup[groupName].removed++;
         writeStatus(FN, `✔️ REMOVED ${login} (${m.Email || ""}) from ${groupName}`);
       } catch (err) {
         const msg = err?.message || String(err);
         failed.push({ email: m.Email || "", group: groupName, action: "remove", error: msg });
         perGroup[groupName].errors++;
         writeStatus(FN, `⚠️ Failed to remove ${login} from ${groupName}: ${msg}`);
       }
     });

     return { plannedAdds: toAdd.length, plannedRemoves: toRemoveMembers.length };
   };

   // Process each real group
   for (const [groupName, desiredSet] of desiredByGroup.entries()) {
     await processOneGroup(groupName, desiredSet);
   }

   // 7) Umbrella group: union of all desired sets (incl. ALL + admin)
   if (umbrellaGroup && umbrellaGroup.trim()) {
     const union = new Set();
     for (const set of desiredByGroup.values()) for (const e of set) union.add(e);
     // Ensure admin + ALL are in union (already added through sets, but re-assert)
     for (const a of allUsers) union.add(a);
     union.add(admin);

     await processOneGroup(umbrellaGroup.trim(), union, { abortIfMissing: true });
   }

   // 8) Final summary + per-group table + error table
   const totals = { addPlanned, added, removePlanned, removed, skippedExisting };
   const groups = Object.keys(perGroup).length;

   // Per-group summary table
   const perGroupRows = Object.entries(perGroup).map(([g, x]) => ({
     Group: g,
     AddPlanned: x.addPlanned,
     Added: x.added,
     RemovePlanned: x.removePlanned,
     Removed: x.removed,
     SkippedExisting: x.skippedExisting,
     Errors: x.errors
   }));
   if (perGroupRows.length) {
     console.table(perGroupRows);
   }

   writeStatus(FN, `Finished. Groups=${groups} | Add: ${added}/${addPlanned} | Remove: ${removed}/${removePlanned} | SkippedExisting=${skippedExisting}`);
   if (failed.length) {
     console.table(failed.map(f => ({ Action: f.action.toUpperCase(), Group: f.group, Email: f.email, Error: f.error })));
   }

   // Add-only style mirrors for convenience
   const attempted = totals.addPlanned;
   const addedCompact = totals.added;
   const skippedExistingCompact = totals.skippedExisting;

   return {
     dryRun,
     groups,
     totals,
     perGroup,
     failed,
     attempted: attempted,
     added: addedCompact,
     skippedExisting: skippedExistingCompact
   };

 } catch (err) {
   writeStatus("syncUsersFromCSVToGroups", `Error: ${err.message}`);
   console.error(err);
   throw err;
 }
};



    // Expose public functions
    return {
      help,
      clearStatus,
      setupContext,
      previewListItems,
      deleteListItems,
      getAllLists,
      getAllSecurityGroups,
      getSiteMembersWithGroups,
      getSiteGroupMembers,
      getSharingLinkForItem,
      removeAllSharingLinksForItem,
      addUsersFromCSVToGroups,
      syncUsersFromCSVToGroups,
      addFieldsToList,
      exportListToExcel,
      getListFields,    
      getListSchemaForMigration
    };

  })();
  