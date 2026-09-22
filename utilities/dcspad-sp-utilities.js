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
 * @function getListSchema(listTitle) · exportListData(listTitle, options) ·
 *           createListFromSchema(schema, options) · importListData(data, options) ·
 *           copyList(sourceTitle, options) · downloadJson(object, fileName) · readJsonFile(filePath)
 *   The list-copy primitives: capture structure, export items, recreate, import — in
 *   memory across two setupContext() calls or through JSON files. All support dryRun.
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
      { name: "getListSchemaForMigration", sig: "(listTitle)", what: "Older read-only schema dump. For copying a list use getListSchema().",
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
        example: `await SPUtils.removeAllSharingLinksForItem("/sites/X/Shared Documents/a.pdf", true)` },
      { name: "getListSchema", sig: "(listTitle)", what: "COPY step 1: settings, fields (SchemaXml, lookups, formatting), views. Rebuildable; no item data.",
        example: `const schema = await SPUtils.getListSchema("Requests")` },
      { name: "exportListData", sig: "(listTitle, { includeAttachments, schema })", what: "COPY step 2: every item as JSON with people and lookups resolved (attachments opt-in).",
        example: `const data = await SPUtils.exportListData("Requests")` },
      { name: "createListFromSchema", sig: "(schema, { title, description, dryRun, lookupMap })", what: "COPY step 3: create the list (or add missing columns/views) on the current site.",
        example: `await SPUtils.createListFromSchema(schema, { dryRun: true })` },
      { name: "importListData", sig: "(data, { listTitle, dryRun, concurrency, preserveAuthorship, includeAttachments, idMap })", what: "COPY step 4: create the items; people by email, lookups by value, self-lookups in pass 2.",
        example: `await SPUtils.importListData(data, { listTitle: "Requests", dryRun: true })` },
      { name: "copyList", sig: "(sourceTitle, { targetTitle, targetSite, dryRun, includeAttachments, preserveAuthorship, lookupMap })", what: "Steps 1–4 chained. Same site → \"<title> Copy\"; targetSite switches context for the create/import.",
        example: `await SPUtils.copyList("Requests", { targetSite: "sites/Archive", dryRun: true })` },
      { name: "downloadJson", sig: "(object, fileName)", what: "Save a schema/data document (or anything) as a .json download.",
        example: `SPUtils.downloadJson(schema, "Requests.schema.json")` },
      { name: "readJsonFile", sig: "(filePath)", what: "Read a .json document from a library on the current site.",
        example: `const schema = await SPUtils.readJsonFile("/sites/X/Shared Documents/Requests.schema.json")` }
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
    // "sites/X", "/sites/X" or an absolute same-tenant URL → normalized absolute site URL.
    const resolveSiteUrl = (siteSubUrl) => {
      const raw = String(siteSubUrl || "").trim();
      return (/^https?:\/\//i.test(raw) ? raw : SPTenantBase + raw.replace(/^\/+/, "")).replace(/\/+$/, "");
    };
    const setupContext = (siteSubUrl) => {
      try {
        const baseUrl = resolveSiteUrl(siteSubUrl);
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
              const throttled = err?.status === 429 || err?.status === 503 || (err?.message || "").includes("Too Many Requests");
              // noRetry marks an error whose own retry budget is already spent (the digest refresh).
              if (err?.noRetry || !throttled || attempt === maxRetries) {
                  throw err;
              }
              // Retry-After (seconds) wins over the exponential schedule when the server sends one.
              const delay = err?.retryAfterMs || baseDelayMs * Math.pow(2, attempt);
              writeStatus("retryWithBackoff", `${err?.status || 429} detected. Retrying in ${delay}ms (attempt ${attempt + 1}/${maxRetries})...`);
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
  
      writeStatus("previewListItems", `Fetching items from list: ${listTitle}`);
  
      items = await fetchAllRows(listPathByTitle(listTitle), { select: "Id,Title", orderby: "ID asc", top: batchSize });
  
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
  
      writeStatus("deleteListItems", `Fetching items from list: ${listTitle}`);
  
      items = await fetchAllRows(listPathByTitle(listTitle), { select: "Id,Title", orderby: "ID asc", top: batchSize });
  
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



// =====================================================================
// LIST COPY PRIMITIVES
//
// Four composable steps plus a wrapper. Each step can be run alone,
// dry-run, and handed a JSON object produced by the previous step —
// in memory across two setupContext() calls, or through a file via
// downloadJson() / readJsonFile().
//
//   const schema = await SPUtils.getListSchema("Requests");
//   const data   = await SPUtils.exportListData("Requests");
//   SPUtils.setupContext("sites/OtherSite");
//   await SPUtils.createListFromSchema(schema, { dryRun: true });
//   await SPUtils.importListData(data, { dryRun: true });
//
// Writes go through ValidateUpdateListItem / AddValidateUpdateItemUsingPath
// with the string conventions SharePoint documents for every field type,
// so one code path covers text, choice, multi-choice, number, date,
// yes/no, URL, person, lookup and managed metadata — and the server
// reports per-field errors instead of a bare 400.
// =====================================================================

const SCHEMA_KIND = "dcspad-sputils-list-schema";
const DATA_KIND = "dcspad-sputils-list-data";

const requireContext = (FN) => {
  if (!contextInitialized) {
    writeStatus(FN, "⚠️ Call setupContext() first.");
    throw new Error("Context not initialized");
  }
};

const escOData = (s) => String(s).replace(/'/g, "''");
const cleanGuid = (g) => String(g || "").replace(/[{}]/g, "").toLowerCase();

// Site URL for raw _api calls: the setupContext() target, else the page's web.
const apiBase = () => siteBaseUrl
  || (typeof _spPageContextInfo !== "undefined" && _spPageContextInfo.webAbsoluteUrl)
  || location.origin;

// An Error carrying the HTTP status and the server's Retry-After (ms), so
// retryWithBackoff can recognise throttling from any of the raw fetch paths.
const httpError = (res, what) => {
  const err = new Error(`HTTP ${res.status} ${res.statusText} for ${what}`);
  err.status = res.status;
  const retryAfter = Number(res.headers?.get?.("Retry-After"));
  if (retryAfter > 0) err.retryAfterMs = retryAfter * 1000;
  return err;
};

// Raw REST paging (nometadata). Follows @odata.nextLink — never $skip,
// which fails past the list view threshold.
const fetchAllRows = async (listPath, { select, filter, orderby, top = 500 } = {}) => {
  const params = [`$top=${top}`];
  if (select) params.push(`$select=${select}`);
  if (filter) params.push(`$filter=${filter}`);
  if (orderby) params.push(`$orderby=${orderby}`);
  let url = `${apiBase()}/_api/web/${listPath}/items?${params.join("&")}`;
  const rows = [];
  while (url) {
    const json = await retryWithBackoff(async () => {
      const res = await fetch(url, {
        headers: { Accept: "application/json;odata=nometadata" },
        credentials: "same-origin"
      });
      if (!res.ok) throw httpError(res, url);
      return res.json();
    });
    rows.push(...(json.value || []));
    url = json["@odata.nextLink"] || null;
  }
  return rows;
};
const listPathByTitle = (title) => `lists/getByTitle('${encodeURIComponent(escOData(title))}')`;
const listPathById = (id) => `lists(guid'${cleanGuid(id)}')`;

// Request digest per site, refreshed a minute before it expires. Item writes
// go through raw REST (ValidateUpdateListItem endpoints) rather than the
// PnPjs item API: PnPjs 2.15 has no AddValidateUpdateItemUsingPath.
const digestCache = new Map(); // base -> { value, expires }
let digestInFlight = null; // single-flight: concurrent workers share one refresh
const getDigest = async (force = false) => {
  const base = apiBase();
  const cached = digestCache.get(base);
  if (!force && cached && cached.expires > Date.now()) return cached.value;
  if (digestInFlight) return digestInFlight;
  digestInFlight = retryWithBackoff(async () => {
    const res = await fetch(`${base}/_api/contextinfo`, {
      method: "POST",
      headers: { Accept: "application/json;odata=nometadata" },
      credentials: "same-origin"
    });
    if (!res.ok) throw httpError(res, "contextinfo");
    const json = await res.json();
    const value = json.FormDigestValue;
    const seconds = Number(json.FormDigestTimeoutSeconds) || 1800;
    digestCache.set(base, { value, expires: Date.now() + (seconds - 60) * 1000 });
    return value;
  }).catch((err) => {
    // The refresh already retried throttling; the write that needed the
    // digest must not restart that budget.
    err.noRetry = true;
    throw err;
  }).finally(() => { digestInFlight = null; });
  return digestInFlight;
};
// GET helper with the same error shape as spPostJson.
const spGetJson = async (path) => {
  const res = await fetch(`${apiBase()}/_api/web/${path}`, {
    headers: { Accept: "application/json;odata=nometadata" }, credentials: "same-origin"
  });
  if (!res.ok) throw httpError(res, path);
  return res.json();
};
const spPostJson = async (path, body) => {
  const send = async (digest) => fetch(`${apiBase()}/_api/web/${path}`, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": digest
    },
    credentials: "same-origin",
    body: JSON.stringify(body)
  });
  let res = await send(await getDigest());
  if (res.status === 403) res = await send(await getDigest(true)); // expired digest looks like 403
  if (!res.ok) {
    let detail = "";
    try { const j = await res.json(); detail = j?.["odata.error"]?.message?.value || j?.error?.message?.value || j?.error?.message || ""; } catch {}
    const err = httpError(res, path);
    if (detail) err.message += `: ${detail}`;
    throw err;
  }
  return res.status === 204 ? {} : res.json();
};
// Both ValidateUpdate endpoints answer with rows of { FieldName, FieldValue, HasException, ErrorMessage }.
const formResultRows = (result) => Array.isArray(result) ? result : (result?.value || result?.results || []);

// ValidateUpdateListItem wants dates the way the web's regional settings
// display them, in the WEB's time zone — never ISO. Verified live on
// en-US, en-GB, en-CA, fr-FR and ja-JP webs: SP.DateFormat 0 = month-day-year,
// 1 = day-month-year, 2 = year-month-day, and the separator varies. The
// definitive source for the date order is the server's own sample in a
// validation error ("Enter a date and time like this: 23/02/2012 02:25 PM"),
// so calibrateDateFormat() asks for it once per site with a create that a
// deliberately invalid date aborts — nothing is written; settings are the
// fallback. The time is always written as a zero-padded 24-hour clock: every
// locale tested accepts it, while AM/PM markers are locale-specific and the
// ja-JP web rejects its own sample form. Time-zone conversion uses the
// server's utcToLocalTime, cached per UTC day, with an exact per-instant
// lookup on DST-transition days.
const regionalCache = new Map(); // base -> { ...settings, tzOffsetByDay, tzOffsetExact, format }
const getRegionalSettings = async () => {
  const base = apiBase();
  if (!regionalCache.has(base)) {
    // Retried like every other call; a failure is an error, never a guessed
    // default — a wrong date order silently corrupts every date written.
    let settings;
    try {
      settings = await retryWithBackoff(() => spGetJson("RegionalSettings?$select=DateFormat,DateSeparator,TimeSeparator,Time24,AM,PM,LocaleId"));
    } catch (err) {
      throw new Error(`Regional settings of ${base} could not be read (${err.message}); dates cannot be written safely.`);
    }
    regionalCache.set(base, { ...settings, tzOffsetByDay: new Map(), tzOffsetExact: new Map(), format: null });
  }
  return regionalCache.get(base);
};

// A date format spec: { order: "mdy"|"dmy"|"ymd", sep, timeSep, source }
const formatFromSettings = (s) => ({
  order: s.DateFormat === 2 ? "ymd" : s.DateFormat === 1 ? "dmy" : "mdy",
  sep: s.DateSeparator || "/",
  timeSep: s.TimeSeparator || ":",
  source: "settings"
});
// Parse "2/23/2012 2:25 PM" / "23/02/2012 02:25 PM" / "2012-02-23 2:25 PM" / "2012/02/23 午後 2:25".
const formatFromSample = (sample, s) => {
  const m = String(sample || "").match(/(\d{1,4})([^\d\s])(\d{1,2})\2(\d{1,4})/);
  if (!m) return null;
  const parts = [m[1], m[3], m[4]];
  const order = parts.map(p => (p.length === 4 ? "y" : Number(p) === 23 ? "d" : "m")).join("");
  if (!/^(mdy|dmy|ymd)$/.test(order)) return null;
  const rest = sample.slice(m.index + m[0].length);
  const tm = rest.match(/\d{1,2}([^\d\s])\d{2}/);
  return { order, sep: m[2], timeSep: tm ? tm[1] : (s.TimeSeparator || ":"), source: "sample" };
};
// Ask the target list for its expected sample. The create is aborted by the
// invalid date (a rejected field never creates an item), so this is read-only.
const calibrateDateFormat = async (listPath, rootFolder, probeField) => {
  const s = await getRegionalSettings();
  if (s.format) return s.format;
  let format = null;
  try {
    const result = await spPostJson(`${listPath}/AddValidateUpdateItemUsingPath`, {
      listItemCreateInfo: { FolderPath: { DecodedUrl: rootFolder }, UnderlyingObjectType: 0 },
      formValues: [{ FieldName: probeField, FieldValue: "not-a-date" }],
      bNewDocumentUpdate: false
    });
    const row = formResultRows(result).find(r => r.FieldName === probeField && r.HasException);
    const sample = row?.ErrorMessage?.split(":").slice(1).join(":") || "";
    format = formatFromSample(sample, s);
  } catch {}
  s.format = format || formatFromSettings(s);
  return s.format;
};

// The web's offset from UTC for an instant, from the server. Retried; on
// failure the caller's date fails, nothing is cached.
const utcOffsetAt = async (utc) => {
  try {
    const local = (await retryWithBackoff(() => spGetJson(`RegionalSettings/TimeZone/utcToLocalTime(@d)?@d='${utc.toISOString()}'`))).value;
    return new Date(`${local}Z`).getTime() - utc.getTime();
  } catch (err) {
    throw new Error(`Web time zone could not be read (${err.message}); date not written.`);
  }
};
const offsetForDay = async (s, day) => {
  if (!s.tzOffsetByDay.has(day)) s.tzOffsetByDay.set(day, await utcOffsetAt(new Date(`${day}T12:00:00Z`)));
  return s.tzOffsetByDay.get(day);
};
const webLocalDate = async (utc) => {
  const s = await getRegionalSettings();
  const dayOf = (d) => d.toISOString().slice(0, 10);
  const day = dayOf(utc);
  const here = await offsetForDay(s, day);
  const prev = await offsetForDay(s, dayOf(new Date(utc.getTime() - 86400000)));
  const next = await offsetForDay(s, dayOf(new Date(utc.getTime() + 86400000)));
  let offsetMs = here;
  if (here !== prev || here !== next) {
    // A transition day: the noon offset may be wrong for this instant.
    const key = utc.toISOString();
    if (!s.tzOffsetExact.has(key)) s.tzOffsetExact.set(key, await utcOffsetAt(utc));
    offsetMs = s.tzOffsetExact.get(key);
  }
  return new Date(utc.getTime() + offsetMs);
};
const toWebDateString = async (value, { dateOnly = false } = {}) => {
  const utc = new Date(value);
  if (Number.isNaN(utc.getTime())) return String(value);
  const s = await getRegionalSettings();
  const f = s.format || formatFromSettings(s);
  const d = await webLocalDate(utc);
  const n2 = (n) => String(n).padStart(2, "0");
  const partsByKey = { y: String(d.getUTCFullYear()), m: n2(d.getUTCMonth() + 1), d: n2(d.getUTCDate()) };
  const dateText = f.order.split("").map(k => partsByKey[k]).join(f.sep);
  if (dateOnly) return dateText;
  return `${dateText} ${n2(d.getUTCHours())}${f.timeSep}${n2(d.getUTCMinutes())}`;
};

// Small promise pool shared by the copy steps.
const runPool = async (items, limit, worker) => {
  const q = [...items];
  const n = Math.max(1, Math.min(limit, q.length || 1));
  await Promise.all(new Array(n).fill(0).map(async () => {
    while (q.length) { await worker(q.shift()); }
  }));
};

const USER_TYPES = new Set(["User", "UserMulti"]);
const LOOKUP_TYPES = new Set(["Lookup", "LookupMulti"]);
const TAXONOMY_TYPES = new Set(["TaxonomyFieldType", "TaxonomyFieldTypeMulti"]);
// Never written on import; SharePoint owns them (authorship is handled separately).
const NEVER_WRITE = new Set(["ID", "Id", "Attachments", "ContentType", "ContentTypeId", "Author", "Editor",
  "Created", "Modified", "GUID", "FileRef", "FileDirRef", "FileLeafRef", "UniqueId", "_UIVersionString", "Order"]);
const NEVER_WRITE_TYPES = new Set(["Computed", "Counter", "Attachments", "File", "ContentTypeId", "Calculated", "Guid"]);

// A "custom" field is one the list author added: not inherited from the base
// type and deletable. Site columns added to the list qualify; system columns
// (Author, Modified, ContentType, _UIVersionString…) do not.
const isCustomField = (f) => !f.FromBaseType && f.CanBeDeleted === true && !f.Hidden
  && f.TypeAsString !== "Computed";

// Strip everything that ties SchemaXml to its source list/web so
// createFieldAsXml recreates the column cleanly on another list.
const scrubSchemaXml = (xml, { lookupListId = null } = {}) => {
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  const el = doc.documentElement;
  if (!el || el.nodeName === "parsererror") throw new Error("SchemaXml could not be parsed");
  for (const attr of ["ID", "SourceID", "ColName", "RowOrdinal", "Version", "WebId", "List", "Sealed", "Customization"]) {
    el.removeAttribute(attr);
  }
  if (lookupListId) el.setAttribute("List", `{${cleanGuid(lookupListId)}}`);
  return new XMLSerializer().serializeToString(el);
};

const base64FromBuffer = (buffer) => {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
};
const bufferFromBase64 = (b64) => {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
};

/**
 * getListSchema
 *
 * Captures everything needed to rebuild a list's structure elsewhere:
 * settings, every visible field (with scrub-ready SchemaXml, lookup targets
 * resolved to list titles, column formatting), views (query, row limit,
 * ordered fields, view formatting) and content types (informational).
 * Item data is NOT included — see exportListData().
 *
 * @param {string} listTitle - Display name of the list
 * @returns {Promise<object>} A `dcspad-sputils-list-schema` document
 * @example
 * const schema = await SPUtils.getListSchema("Requests");
 * SPUtils.downloadJson(schema, "Requests.schema.json");
 */
const getListSchema = async (listTitle) => {
  const FN = "getListSchema";
  requireContext(FN);
  writeStatus(FN, `Reading schema of "${listTitle}"...`);
  const web = pnp2.sp.web;
  const list = web.lists.getByTitle(listTitle);
  const warnings = [];

  const settings = await list
    .select("Id", "Title", "Description", "BaseTemplate", "EnableVersioning", "MajorVersionLimit",
      "EnableMinorVersions", "MajorWithMinorVersionsLimit", "DraftVersionVisibility", "ForceCheckout",
      "Hidden", "ContentTypesEnabled", "EnableAttachments", "EnableFolderCreation", "EnableModeration",
      "ItemCount", "RootFolder/ServerRelativeUrl")
    .expand("RootFolder")();

  if (settings.BaseTemplate !== 100) {
    warnings.push(`BaseTemplate ${settings.BaseTemplate}: only generic lists (100) are fully supported; libraries need a file copy, not an item import.`);
  }
  if (settings.ContentTypesEnabled) {
    warnings.push("Content types are enabled on the source; they are exported for reference but not recreated.");
  }

  // No $select on purpose: fields are heterogeneous and type-specific
  // properties (LookupList, Choices, CustomFormatter…) only exist on some.
  const rawFields = await list.fields.filter("Hidden eq false")();
  const listTitleCache = new Map();
  const lookupTitle = async (guid) => {
    const id = cleanGuid(guid);
    if (!id) return null;
    if (id === cleanGuid(settings.Id)) return settings.Title;
    if (!listTitleCache.has(id)) {
      try {
        const target = await web.lists.getById(id).select("Title")();
        listTitleCache.set(id, target.Title);
      } catch { listTitleCache.set(id, null); }
    }
    return listTitleCache.get(id);
  };

  const fields = [];
  for (const f of rawFields) {
    const type = f.TypeAsString;
    const isLookup = LOOKUP_TYPES.has(type);
    const custom = isCustomField(f);
    // Lookup targets are resolved for custom columns only; system lookup-typed
    // fields (ItemChildCount, AppAuthor, compliance columns) point at hidden lists.
    const lookupListId = isLookup && custom ? cleanGuid(f.LookupList) : null;
    const entry = {
      id: cleanGuid(f.Id),
      internalName: f.InternalName,
      staticName: f.StaticName,
      displayName: f.Title,
      type,
      required: !!f.Required,
      readOnly: !!f.ReadOnlyField,
      fromBaseType: !!f.FromBaseType,
      custom,
      description: f.Description || "",
      defaultValue: f.DefaultValue ?? null,
      choices: Array.isArray(f.Choices) ? f.Choices : (f.Choices?.results || []),
      maxLength: f.MaxLength ?? null,
      indexed: !!f.Indexed,
      enforceUniqueValues: !!f.EnforceUniqueValues,
      allowMultipleValues: !!f.AllowMultipleValues || type.endsWith("Multi"),
      customFormatter: f.CustomFormatter || "",
      lookupListId,
      lookupList: lookupListId ? await lookupTitle(lookupListId) : null,
      lookupField: isLookup ? (f.LookupField || "Title") : null,
      isSelfLookup: !!lookupListId && lookupListId === cleanGuid(settings.Id),
      // A secondary ("additional") lookup column depends on a primary lookup by field id.
      isDependentLookup: !!f.IsDependentLookup,
      primaryFieldId: f.IsDependentLookup ? cleanGuid(f.PrimaryFieldId) : null,
      schemaXml: f.SchemaXml
    };
    if (lookupListId && !entry.lookupList) warnings.push(`Lookup "${f.InternalName}" points at a list that could not be read; it will need a lookupMap entry.`);
    if (TAXONOMY_TYPES.has(type) && entry.custom) warnings.push(`Managed metadata column "${f.InternalName}" cannot be recreated automatically (needs a term-set binding).`);
    fields.push(entry);
  }

  const rawViews = await list.views.filter("PersonalView eq false")();
  const views = [];
  for (const v of rawViews) {
    let viewFields = [];
    try {
      const vf = await list.views.getById(v.Id).fields();
      viewFields = vf.Items || vf.results || [];
    } catch (err) {
      warnings.push(`View "${v.Title}": fields could not be read (${err.message}).`);
    }
    views.push({
      id: v.Id,
      title: v.Title,
      defaultView: !!v.DefaultView,
      hidden: !!v.Hidden,
      viewType: v.ViewType || "HTML",
      viewQuery: v.ViewQuery || "",
      rowLimit: v.RowLimit ?? 30,
      paged: v.Paged !== false,
      customFormatter: v.CustomFormatter || "",
      jsLink: v.JSLink || "",
      fields: viewFields
    });
  }

  let contentTypes = [];
  try {
    contentTypes = (await list.contentTypes.select("Name", "StringId", "Description", "Group", "Hidden", "ReadOnly")())
      .map(ct => ({ name: ct.Name, id: ct.StringId, description: ct.Description, group: ct.Group, hidden: ct.Hidden, readOnly: ct.ReadOnly }));
  } catch (err) {
    warnings.push(`Content types could not be read (${err.message}).`);
  }

  const schema = {
    kind: SCHEMA_KIND,
    version: 1,
    exported: new Date().toISOString(),
    source: {
      siteUrl: apiBase(),
      listTitle: settings.Title,
      listId: cleanGuid(settings.Id),
      rootFolder: settings.RootFolder?.ServerRelativeUrl || null,
      itemCount: settings.ItemCount
    },
    list: {
      title: settings.Title,
      description: settings.Description || "",
      baseTemplate: settings.BaseTemplate,
      enableVersioning: !!settings.EnableVersioning,
      majorVersionLimit: settings.MajorVersionLimit ?? null,
      enableMinorVersions: !!settings.EnableMinorVersions,
      majorWithMinorVersionsLimit: settings.MajorWithMinorVersionsLimit ?? null,
      draftVersionVisibility: settings.DraftVersionVisibility ?? 0,
      forceCheckout: !!settings.ForceCheckout,
      hidden: !!settings.Hidden,
      contentTypesEnabled: !!settings.ContentTypesEnabled,
      enableAttachments: settings.EnableAttachments !== false,
      enableFolderCreation: !!settings.EnableFolderCreation,
      enableModeration: !!settings.EnableModeration
    },
    fields,
    views,
    contentTypes,
    warnings
  };

  writeStatus(FN, `Done. Fields=${fields.length} (custom ${fields.filter(f => f.custom).length}), Views=${views.length}, Warnings=${warnings.length}`);
  warnings.forEach(w => writeStatus(FN, `⚠️ ${w}`));
  return schema;
};

// Id → user cache for the current site, filled lazily.
const siteUserCache = new Map(); // siteBase -> Map(id -> {Id, Email, LoginName, Title})
const getSiteUsers = async () => {
  const base = apiBase();
  if (!siteUserCache.has(base)) {
    const users = await pnp2.sp.web.siteUsers.select("Id", "Email", "LoginName", "Title")();
    siteUserCache.set(base, new Map(users.map(u => [u.Id, u])));
  }
  return siteUserCache.get(base);
};
const resolveUser = async (id) => {
  const users = await getSiteUsers();
  if (!users.has(id)) {
    try {
      const u = await pnp2.sp.web.getUserById(id).select("Id", "Email", "LoginName", "Title")();
      users.set(id, u);
    } catch { users.set(id, null); }
  }
  const u = users.get(id);
  return u ? { Id: u.Id, Email: u.Email || "", LoginName: u.LoginName || "", Title: u.Title || "" } : { Id: id, Email: "", LoginName: "", Title: "" };
};

/**
 * exportListData
 *
 * Exports every item as raw REST values plus resolved forms of the values
 * that do not survive a move: person fields carry email and login, lookup
 * fields carry the looked-up text. Author/Editor are resolved too so
 * importListData() can preserve authorship. Attachments are opt-in (base64).
 *
 * @param {string} listTitle - Display name of the list
 * @param {object} [options]
 * @param {boolean} [options.includeAttachments=false] - Embed attachment bytes (base64)
 * @param {object}  [options.schema] - A getListSchema() result to reuse (saves a round trip)
 * @returns {Promise<object>} A `dcspad-sputils-list-data` document
 * @example
 * const data = await SPUtils.exportListData("Requests", { includeAttachments: true });
 * SPUtils.downloadJson(data, "Requests.data.json");
 */
const exportListData = async (listTitle, { includeAttachments = false, schema = null } = {}) => {
  const FN = "exportListData";
  requireContext(FN);
  const sch = schema && schema.kind === SCHEMA_KIND ? schema : await getListSchema(listTitle);
  const warnings = [];
  writeStatus(FN, `Fetching items of "${listTitle}"...`);
  const rows = await fetchAllRows(listPathByTitle(listTitle), { select: "*,FSObjType,FileDirRef,FileRef", orderby: "ID asc" });
  const rootFolder = (sch.source.rootFolder || "").replace(/\/+$/, "");
  const relPath = (serverRelative) => {
    const p = String(serverRelative || "");
    if (!rootFolder || !p.toLowerCase().startsWith(rootFolder.toLowerCase())) return "";
    return p.slice(rootFolder.length).replace(/^\/+/, "");
  };
  const folderCount = rows.filter(r => Number(r.FSObjType) === 1).length;
  writeStatus(FN, `Items: ${rows.length}${folderCount ? ` (${folderCount} folder(s); items keep their folder path)` : ""}`);

  const userFields = sch.fields.filter(f => USER_TYPES.has(f.type) && (f.custom || f.internalName === "Author" || f.internalName === "Editor"));
  const lookupFields = sch.fields.filter(f => LOOKUP_TYPES.has(f.type) && f.custom);

  // Lookup value maps: id -> shown value, one fetch per lookup target.
  const lookupValues = new Map(); // internalName -> Map(id -> value)
  for (const f of lookupFields) {
    if (!f.lookupListId) continue;
    try {
      const targetRows = await fetchAllRows(listPathById(f.lookupListId), { select: `Id,${f.lookupField}` });
      lookupValues.set(f.internalName, new Map(targetRows.map(r => [r.Id, r[f.lookupField]])));
    } catch (err) {
      warnings.push(`Lookup "${f.internalName}": target list could not be read (${err.message}); ids exported without values.`);
    }
  }

  const toIdList = (v) => (v == null ? [] : Array.isArray(v) ? v : [v]).filter(x => x != null);
  const items = [];
  for (const raw of rows) {
    const resolved = {};
    for (const f of userFields) {
      const ids = toIdList(raw[`${f.internalName}Id`]);
      if (ids.length) resolved[f.internalName] = await Promise.all(ids.map(resolveUser));
    }
    for (const f of lookupFields) {
      const ids = toIdList(raw[`${f.internalName}Id`]);
      if (ids.length) {
        const map = lookupValues.get(f.internalName);
        resolved[f.internalName] = ids.map(id => ({ Id: id, value: map ? (map.get(id) ?? null) : null }));
      }
    }
    const isFolder = Number(raw.FSObjType) === 1;
    const item = { ...raw, _resolved: resolved, _dir: relPath(raw.FileDirRef) };
    if (isFolder) { item._folder = true; item._folderPath = relPath(raw.FileRef); }
    items.push(item);
  }

  let attachmentCount = 0;
  if (includeAttachments) {
    const list = pnp2.sp.web.lists.getByTitle(listTitle);
    const withAttachments = items.filter(i => i.Attachments === true && !i._folder);
    writeStatus(FN, `Reading attachments for ${withAttachments.length} item(s)...`);
    await runPool(withAttachments, 4, async (item) => {
      try {
        const files = await list.items.getById(item.Id).attachmentFiles();
        item._attachments = [];
        for (const a of files) {
          const buffer = await retryWithBackoff(() => pnp2.sp.web.getFileByServerRelativePath(a.ServerRelativeUrl).getBuffer());
          item._attachments.push({ name: a.FileName, base64: base64FromBuffer(buffer) });
          attachmentCount++;
        }
      } catch (err) {
        warnings.push(`Item ${item.Id}: attachments could not be read (${err.message}).`);
      }
    });
  }

  const fieldMap = {};
  for (const f of sch.fields) {
    fieldMap[f.internalName] = {
      type: f.type, custom: f.custom, readOnly: f.readOnly,
      lookupList: f.lookupList, lookupListId: f.lookupListId, lookupField: f.lookupField, isSelfLookup: f.isSelfLookup,
      allowMultipleValues: f.allowMultipleValues
    };
  }

  const data = {
    kind: DATA_KIND,
    version: 1,
    exported: new Date().toISOString(),
    source: { siteUrl: apiBase(), listTitle: sch.source.listTitle, listId: sch.source.listId },
    fields: fieldMap,
    items,
    warnings
  };
  writeStatus(FN, `Done. Items=${items.length}${includeAttachments ? `, Attachments=${attachmentCount}` : ""}, Warnings=${warnings.length}`);
  warnings.forEach(w => writeStatus(FN, `⚠️ ${w}`));
  return data;
};

/**
 * createListFromSchema
 *
 * Creates a list from a getListSchema() document on the current site, or
 * adds whatever is missing to a list of that title that already exists.
 * Custom columns are recreated from scrubbed SchemaXml in dependency order
 * (plain → lookup → calculated); the Title column's display name and
 * required flag are applied; views are created or updated with their
 * query, fields and formatting.
 *
 * @param {object} schema - A `dcspad-sputils-list-schema` document
 * @param {object} [options]
 * @param {string}  [options.title] - Target list title (default: the source title)
 * @param {string}  [options.description]
 * @param {boolean} [options.dryRun=false] - Log the plan, change nothing
 * @param {object}  [options.lookupMap] - Source lookup list title → target list title
 * @returns {Promise<object>} Report: listId, created, fields, views, warnings
 * @example
 * await SPUtils.createListFromSchema(schema, { title: "Requests", dryRun: true });
 */
const createListFromSchema = async (schema, { title, description, dryRun = false, lookupMap = {} } = {}) => {
  const FN = "createListFromSchema";
  requireContext(FN);
  if (!schema || schema.kind !== SCHEMA_KIND) throw new Error(`Expected a ${SCHEMA_KIND} document (from getListSchema).`);
  const web = pnp2.sp.web;
  const targetTitle = title || schema.list.title;
  const report = {
    dryRun, title: targetTitle, listId: null, created: false,
    fields: { added: 0, skipped: 0, failed: [] },
    views: { added: 0, updated: 0, failed: [] },
    warnings: [...(schema.warnings || [])]
  };
  const warn = (m) => { report.warnings.push(m); writeStatus(FN, `⚠️ ${m}`); };

  // ---- list ------------------------------------------------------------
  const existing = await web.lists.filter(`Title eq '${escOData(targetTitle)}'`).select("Id", "Title", "BaseTemplate")();
  let list = null;
  if (existing.length) {
    list = web.lists.getById(existing[0].Id);
    report.listId = cleanGuid(existing[0].Id);
    writeStatus(FN, `List "${targetTitle}" exists — adding missing fields and views only.`);
    // Reconcile the settings an import depends on; leave the rest alone.
    const reconcile = {};
    if (schema.list.enableAttachments) reconcile.EnableAttachments = true;
    if (schema.list.enableFolderCreation) reconcile.EnableFolderCreation = true;
    if (Object.keys(reconcile).length && !dryRun) {
      try { await retryWithBackoff(() => list.update(reconcile)); }
      catch (err) { warn(`Existing list settings not reconciled (${Object.keys(reconcile).join(", ")}): ${err.message}`); }
    }
  } else if (dryRun) {
    writeStatus(FN, `DRY RUN: would create list "${targetTitle}" (template ${schema.list.baseTemplate}).`);
  } else {
    const r = await retryWithBackoff(() => web.lists.add(
      targetTitle,
      description ?? schema.list.description ?? "",
      schema.list.baseTemplate || 100,
      !!schema.list.contentTypesEnabled
    ));
    list = r.list;
    report.listId = cleanGuid(r.data.Id);
    report.created = true;
    writeStatus(FN, `Created list "${targetTitle}".`);
    const s = schema.list;
    const settings = {
      EnableVersioning: s.enableVersioning,
      EnableAttachments: s.enableAttachments,
      EnableFolderCreation: s.enableFolderCreation,
      EnableModeration: s.enableModeration,
      ForceCheckout: !!s.forceCheckout,
      Hidden: !!s.hidden
    };
    if (s.enableVersioning && s.majorVersionLimit) settings.MajorVersionLimit = s.majorVersionLimit;
    if (s.enableVersioning && s.enableMinorVersions) {
      settings.EnableMinorVersions = true;
      if (s.majorWithMinorVersionsLimit) settings.MajorWithMinorVersionsLimit = s.majorWithMinorVersionsLimit;
    }
    if (s.enableModeration || s.enableMinorVersions) settings.DraftVersionVisibility = s.draftVersionVisibility;
    try { await retryWithBackoff(() => list.update(settings)); }
    catch (err) { warn(`List settings could not all be applied: ${err.message}`); }
  }

  // ---- fields ----------------------------------------------------------
  const existingFields = new Map(); // internalName -> field
  const existingTitles = new Set();
  if (list) {
    for (const f of await list.fields.select("Id", "InternalName", "Title", "TypeAsString")()) {
      existingFields.set(f.InternalName, f);
      existingTitles.add(f.Title.toLowerCase());
    }
  }
  const targetListIdByTitle = new Map();
  const resolveTargetList = async (sourceTitle) => {
    const mapped = lookupMap[sourceTitle] || sourceTitle;
    if (!targetListIdByTitle.has(mapped)) {
      const found = await web.lists.filter(`Title eq '${escOData(mapped)}'`).select("Id")();
      targetListIdByTitle.set(mapped, found.length ? cleanGuid(found[0].Id) : null);
    }
    return { title: mapped, id: targetListIdByTitle.get(mapped) };
  };

  const custom = schema.fields.filter(f => f.custom);
  // plain → lookup → dependent lookup (needs its primary's new id) → calculated
  const tier = (f) => (f.type === "Calculated" ? 3 : f.isDependentLookup ? 2 : LOOKUP_TYPES.has(f.type) ? 1 : 0);
  const ordered = [...custom].sort((a, b) => tier(a) - tier(b));
  const fieldIdMap = new Map(); // source field id -> target field id

  for (const f of ordered) {
    try {
      const present = existingFields.get(f.internalName);
      if (present) {
        if (present.TypeAsString !== f.type) {
          throw new Error(`exists on the target as ${present.TypeAsString}, source is ${f.type}; values will not import`);
        }
        if (present.Id && f.id) fieldIdMap.set(f.id, cleanGuid(present.Id));
        report.fields.skipped++;
        continue;
      }
      if (existingTitles.has(f.displayName.toLowerCase())) {
        throw new Error(`a different column already uses the display name "${f.displayName}"`);
      }
      if (TAXONOMY_TYPES.has(f.type)) throw new Error("managed metadata columns are not recreated automatically");

      let lookupListId = null;
      if (LOOKUP_TYPES.has(f.type)) {
        if (f.isSelfLookup) {
          lookupListId = report.listId;
          if (!lookupListId && dryRun) lookupListId = "00000000-0000-0000-0000-000000000000";
        } else {
          const target = await resolveTargetList(f.lookupList);
          if (!target.id) throw new Error(`lookup target list "${target.title}" was not found on this site (add a lookupMap entry)`);
          lookupListId = target.id;
        }
      }
      // A secondary lookup column is bound to its primary by field id (FieldRef
      // in SchemaXml), so it goes through the dependent-lookup endpoint with
      // the primary's NEW id rather than through createFieldAsXml.
      if (f.isDependentLookup) {
        const primaryFieldId = fieldIdMap.get(f.primaryFieldId) || null;
        if (!primaryFieldId && !dryRun) throw new Error("its primary lookup column was not created on the target");
        if (dryRun) {
          writeStatus(FN, `DRY RUN: add dependent lookup "${f.displayName}" (${f.internalName}) → primary ${f.primaryFieldId}`);
          report.fields.added++;
          continue;
        }
        const dep = await retryWithBackoff(() => list.fields.addDependentLookupField(f.displayName, primaryFieldId, f.lookupField || "Title"));
        if (f.id && dep?.data?.Id) fieldIdMap.set(f.id, cleanGuid(dep.data.Id));
        existingFields.set(f.internalName, { Id: dep?.data?.Id, InternalName: f.internalName, Title: f.displayName, TypeAsString: f.type });
        existingTitles.add(f.displayName.toLowerCase());
        report.fields.added++;
        writeStatus(FN, `ADDED dependent lookup "${f.displayName}" (${f.internalName})`);
        continue;
      }
      const xml = scrubSchemaXml(f.schemaXml, { lookupListId });

      if (dryRun) {
        writeStatus(FN, `DRY RUN: add ${f.type} "${f.displayName}" (${f.internalName})${lookupListId ? ` → list ${lookupListId}` : ""}`);
        report.fields.added++;
        continue;
      }
      const createdField = await retryWithBackoff(() => list.fields.createFieldAsXml(xml));
      if (f.id && createdField?.data?.Id) fieldIdMap.set(f.id, cleanGuid(createdField.data.Id));
      if (f.customFormatter && !/CustomFormatter=/.test(xml)) {
        try { await list.fields.getByInternalNameOrTitle(f.internalName).update({ CustomFormatter: f.customFormatter }); }
        catch (err) { warn(`Column formatting for "${f.internalName}" was not applied: ${err.message}`); }
      }
      existingFields.set(f.internalName, { Id: createdField?.data?.Id, InternalName: f.internalName, Title: f.displayName, TypeAsString: f.type });
      existingTitles.add(f.displayName.toLowerCase());
      report.fields.added++;
      writeStatus(FN, `ADDED ${f.type} "${f.displayName}" (${f.internalName})`);
    } catch (err) {
      report.fields.failed.push({ internalName: f.internalName, displayName: f.displayName, error: err.message });
      writeStatus(FN, `⚠️ Field "${f.displayName}" failed: ${err.message}`);
    }
  }

  // Title column: display name / required are per-list customizations.
  const titleField = schema.fields.find(f => f.internalName === "Title");
  if (titleField && (titleField.displayName !== "Title" || titleField.required === false)) {
    if (dryRun) {
      writeStatus(FN, `DRY RUN: set Title column display name "${titleField.displayName}", required=${titleField.required}`);
    } else if (list) {
      try {
        await list.fields.getByInternalNameOrTitle("Title").update({ Title: titleField.displayName, Required: titleField.required });
      } catch (err) { warn(`Title column settings not applied: ${err.message}`); }
    }
  }

  // ---- views -----------------------------------------------------------
  for (const v of schema.views.filter(v => !v.hidden)) {
    try {
      const wanted = v.fields.filter(name => existingFields.has(name) || dryRun);
      const missing = v.fields.filter(name => !existingFields.has(name));
      if (missing.length && !dryRun) warn(`View "${v.title}": columns not on the target were left out: ${missing.join(", ")}`);
      if (dryRun) {
        writeStatus(FN, `DRY RUN: view "${v.title}" [${wanted.join(", ")}]${v.customFormatter ? " + formatting" : ""}`);
        report.views.added++;
        continue;
      }
      const found = await list.views.filter(`Title eq '${escOData(v.title)}'`).select("Id")();
      const addAll = async (view, names) => {
        for (const name of names) await retryWithBackoff(() => view.fields.add(name));
      };
      const props = { ViewQuery: v.viewQuery, RowLimit: v.rowLimit, Paged: v.paged };
      if (found.length) {
        // Existing view: rebuild the columns FIRST and touch nothing else
        // until they succeed; on failure restore the previous columns so the
        // view keeps its old, consistent state. Blank source formatting and
        // JSLink are sent too, so stale target values are cleared.
        const view = list.views.getById(found[0].Id);
        let previous = [];
        try { const vf = await view.fields(); previous = vf.Items || vf.results || []; } catch {}
        await retryWithBackoff(() => view.fields.removeAll());
        try {
          await addAll(view, wanted);
        } catch (err) {
          warn(`View "${v.title}": rebuilding its columns failed (${err.message})${previous.length ? "; restoring the previous columns" : ""}`);
          try { await retryWithBackoff(() => view.fields.removeAll()); await addAll(view, previous); } catch {}
          throw err;
        }
        await retryWithBackoff(() => view.update({ ...props, CustomFormatter: v.customFormatter || "", JSLink: v.jsLink || "" }));
        if (v.defaultView) { try { await view.update({ DefaultView: true }); } catch {} }
        report.views.updated++;
      } else {
        const r = await retryWithBackoff(() => list.views.add(v.title, false, { ...props, DefaultView: v.defaultView }));
        const view = r.view;
        await retryWithBackoff(() => view.fields.removeAll());
        try {
          await addAll(view, wanted);
        } catch (err) {
          // A half-built new view is worse than none.
          try { await view.delete(); } catch {}
          throw err;
        }
        const extras = {};
        if (v.customFormatter) extras.CustomFormatter = v.customFormatter;
        if (v.jsLink) extras.JSLink = v.jsLink;
        if (Object.keys(extras).length) {
          try { await retryWithBackoff(() => view.update(extras)); }
          catch (err) { warn(`View "${v.title}": formatting/JSLink not applied: ${err.message}`); }
        }
        report.views.added++;
      }
      writeStatus(FN, `${found.length ? "UPDATED" : "ADDED"} view "${v.title}" (${wanted.length} columns)`);
    } catch (err) {
      report.views.failed.push({ title: v.title, error: err.message });
      writeStatus(FN, `⚠️ View "${v.title}" failed: ${err.message}`);
    }
  }

  writeStatus(FN, `${dryRun ? "DRY RUN complete" : "Done"}. Fields: +${report.fields.added} / skipped ${report.fields.skipped} / failed ${report.fields.failed.length}. Views: +${report.views.added} / updated ${report.views.updated} / failed ${report.views.failed.length}.`);
  return report;
};

/**
 * importListData
 *
 * Creates items in the current site's list from an exportListData()
 * document. Values are written with the ValidateUpdateListItem string
 * conventions, person fields are re-resolved by email, lookup fields by
 * their shown value against the target's lookup list, and self-referencing
 * lookups are filled in a second pass through the old→new id map.
 *
 * @param {object} data - A `dcspad-sputils-list-data` document
 * @param {object} [options]
 * @param {string}  [options.listTitle] - Target list (default: the source title)
 * @param {boolean} [options.dryRun=false] - Plan only; shows sample payloads
 * @param {number}  [options.concurrency=1] - Parallel creates. 1 keeps source order (new ids ascend like the
 *                                            old ones); higher is faster but scrambles the order.
 * @param {boolean} [options.preserveAuthorship=false] - Keep Created/Modified/Author/Editor
 * @param {boolean} [options.includeAttachments=true] - Re-upload embedded attachments
 * @param {object}  [options.idMap] - A previous report's idMap: source items already in it are skipped,
 *                                    so a partial import can be resumed without duplicating rows
 * @returns {Promise<object>} Report: created, skipped, failed, fieldErrors, idMap, attachments, folders
 * @example
 * const r = await SPUtils.importListData(data, { listTitle: "Requests", dryRun: true });
 */
const importListData = async (data, { listTitle, dryRun = false, concurrency = 1, preserveAuthorship = false, includeAttachments = true, idMap: idMapSeed = null } = {}) => {
  const FN = "importListData";
  requireContext(FN);
  if (!data || data.kind !== DATA_KIND) throw new Error(`Expected a ${DATA_KIND} document (from exportListData).`);
  const web = pnp2.sp.web;
  const targetTitle = listTitle || data.source.listTitle;
  const report = {
    dryRun, listTitle: targetTitle, attempted: data.items.filter(i => !i._folder).length, created: 0, skipped: 0,
    failed: [], fieldErrors: [], idMap: { ...(idMapSeed || {}) }, attachments: { added: 0, failed: 0 },
    folders: { created: 0, failed: 0 }, warnings: []
  };
  if (idMapSeed) writeStatus(FN, `Resuming: ${Object.keys(idMapSeed).length} source item(s) already imported will be skipped.`);
  const warn = (m) => { report.warnings.push(m); writeStatus(FN, `⚠️ ${m}`); };

  const found = await web.lists.filter(`Title eq '${escOData(targetTitle)}'`).select("Id", "RootFolder/ServerRelativeUrl").expand("RootFolder")();
  if (!found.length) {
    if (dryRun) {
      writeStatus(FN, `DRY RUN: target list "${targetTitle}" does not exist yet; ${data.items.length} item(s) would be imported after createListFromSchema().`);
      return report;
    }
    throw new Error(`Target list "${targetTitle}" was not found. Run createListFromSchema() first.`);
  }
  const list = web.lists.getById(found[0].Id);
  const rootFolder = found[0].RootFolder.ServerRelativeUrl;
  const targetFields = new Map((await list.fields.filter("Hidden eq false")()).map(f => [f.InternalName, f]));

  // Writable = present on target, not read-only, not system-owned, and
  // either custom on the source or the Title column.
  const writable = [];
  const deferredSelf = [];
  for (const [name, meta] of Object.entries(data.fields)) {
    const tf = targetFields.get(name);
    if (!tf || tf.ReadOnlyField || NEVER_WRITE.has(name) || NEVER_WRITE_TYPES.has(tf.TypeAsString)) continue;
    if (!(meta.custom || name === "Title")) continue;
    if (LOOKUP_TYPES.has(tf.TypeAsString) && meta.isSelfLookup) { deferredSelf.push({ name, tf, meta }); continue; }
    // The source-id shortcut for ambiguous lookup values is only meaningful
    // when the target column points at the very same list as the source.
    const sameLookupList = LOOKUP_TYPES.has(tf.TypeAsString) && !!meta.lookupListId && cleanGuid(tf.LookupList) === meta.lookupListId;
    writable.push({ name, tf, meta, sameLookupList });
  }
  writeStatus(FN, `Target "${targetTitle}": ${writable.length} writable column(s)${deferredSelf.length ? `, ${deferredSelf.length} self-lookup(s) in pass 2` : ""}.`);

  // Learn the web's date format from the server before writing any date.
  const dateField = writable.find(w => w.tf.TypeAsString === "DateTime");
  if (dateField || preserveAuthorship) {
    const f = await calibrateDateFormat(listPathById(found[0].Id), rootFolder, dateField ? dateField.name : "Created");
    writeStatus(FN, `Date format (${f.source}): ${f.order} sep "${f.sep}", 24-hour clock`);
  }

  // Lookup targets on the TARGET side: shown value -> [ids]. A value that
  // appears more than once is ambiguous; see lookupValue().
  const lookupIdsByValue = new Map();
  for (const w of writable) {
    if (!LOOKUP_TYPES.has(w.tf.TypeAsString)) continue;
    try {
      const showField = w.tf.LookupField || "Title";
      const rows = await fetchAllRows(listPathById(w.tf.LookupList), { select: `Id,${showField}` });
      const byValue = new Map();
      for (const r of rows) {
        const key = String(r[showField] ?? "");
        if (!byValue.has(key)) byValue.set(key, []);
        byValue.get(key).push(r.Id);
      }
      lookupIdsByValue.set(w.name, byValue);
    } catch (err) {
      warn(`Lookup "${w.name}": target lookup list could not be read (${err.message}); values will be left empty.`);
    }
  }

  // Person resolution by email/login, cached.
  const loginCache = new Map();
  const loginFor = async (user) => {
    const key = (user.Email || user.LoginName || "").toLowerCase();
    if (!key) return null;
    // App/system principals ("SharePoint App", SYSTEM) have no email and can't
    // be matched by ValidateUpdateListItem; leave the field empty instead.
    if (!user.Email && !/^i:0#\.f\|membership\|/i.test(user.LoginName || "")) return null;
    // A dry run must not touch the target: ensureUser can add a principal to
    // the site's user list, so derive the claims key from the email instead.
    if (dryRun) return user.Email ? `i:0#.f|membership|${user.Email.toLowerCase()}` : user.LoginName;
    if (!loginCache.has(key)) {
      try {
        const ensured = await retryWithBackoff(() => web.ensureUser(user.Email || user.LoginName));
        loginCache.set(key, ensured.data.LoginName);
      } catch {
        loginCache.set(key, user.Email ? `i:0#.f|membership|${user.Email.toLowerCase()}` : null);
      }
    }
    return loginCache.get(key);
  };
  const userValue = async (users) => {
    const keys = [];
    for (const u of users || []) { const l = await loginFor(u); if (l) keys.push({ Key: l }); }
    return keys.length ? JSON.stringify(keys) : "";
  };
  // Resolve a lookup by its shown value on the target. Ambiguous values are
  // only accepted when one candidate carries the source id (same-site copies);
  // missing and ambiguous references are reported per field, never guessed.
  const lookupValue = (name, resolvedList, multi, errs, sameLookupList) => {
    const byValue = lookupIdsByValue.get(name);
    if (!byValue) return "";
    const ids = [];
    for (const r of resolvedList || []) {
      if (r.value == null || r.value === "") {
        errs.push({ field: name, message: `source id ${r.Id} had no shown value on the source; reference left empty` });
        continue;
      }
      const value = String(r.value);
      const candidates = byValue.get(value) || [];
      if (candidates.length === 1) { ids.push(candidates[0]); continue; }
      if (candidates.length > 1 && sameLookupList && candidates.includes(r.Id)) { ids.push(r.Id); continue; }
      errs.push({ field: name, message: candidates.length
        ? `"${value}" matches ${candidates.length} items on the target lookup list; reference left empty`
        : `"${value}" (source id ${r.Id}) was not found on the target lookup list; reference left empty` });
    }
    if (!ids.length) return "";
    return multi ? `${ids.join(";#")};#` : String(ids[0]);
  };
  const taxonomyValue = (raw) => {
    const terms = (raw == null ? [] : Array.isArray(raw) ? raw : [raw]).filter(t => t && t.Label);
    return terms.map(t => `${t.Label}|${cleanGuid(t.TermGuid)};`).join("");
  };

  const toFormValue = async (w, item, errs) => {
    const raw = item[w.name];
    const type = w.tf.TypeAsString;
    switch (type) {
      case "Boolean": return raw == null ? "" : (raw ? "1" : "0");
      case "MultiChoice": {
        const arr = raw == null ? [] : Array.isArray(raw) ? raw : [raw];
        return arr.length ? `;#${arr.join(";#")};#` : "";
      }
      case "URL": return raw && raw.Url ? `${raw.Url}, ${raw.Description || raw.Url}` : "";
      case "DateTime": return raw ? toWebDateString(raw, { dateOnly: w.tf.DisplayFormat === 0 }) : "";
      case "User": case "UserMulti": return userValue(item._resolved?.[w.name]);
      case "Lookup": case "LookupMulti": return lookupValue(w.name, item._resolved?.[w.name], type === "LookupMulti", errs, w.sameLookupList);
      case "TaxonomyFieldType": case "TaxonomyFieldTypeMulti": return taxonomyValue(raw);
      default:
        if (raw == null) return "";
        return typeof raw === "object" ? JSON.stringify(raw) : String(raw);
    }
  };

  const authorshipValues = async (item) => {
    const values = [];
    const author = await userValue(item._resolved?.Author);
    const editor = await userValue(item._resolved?.Editor);
    if (author) values.push({ FieldName: "Author", FieldValue: author });
    if (editor) values.push({ FieldName: "Editor", FieldValue: editor });
    if (item.Created) values.push({ FieldName: "Created", FieldValue: await toWebDateString(item.Created) });
    if (item.Modified) values.push({ FieldName: "Modified", FieldValue: await toWebDateString(item.Modified) });
    return values;
  };
  const buildFormValues = async (item) => {
    const values = [];
    const errs = [];
    for (const w of writable) values.push({ FieldName: w.name, FieldValue: await toFormValue(w, item, errs) });
    if (preserveAuthorship) values.push(...await authorshipValues(item));
    return { values, errs };
  };
  const folderPathFor = (item) => (item._dir ? `${rootFolder}/${item._dir}` : rootFolder);

  const folders = data.items.filter(i => i._folder && i._folderPath)
    .sort((a, b) => a._folderPath.split("/").length - b._folderPath.split("/").length || a.Id - b.Id);
  const items = data.items.filter(i => !i._folder).sort((a, b) => a.Id - b.Id);

  if (dryRun) {
    if (folders.length) writeStatus(FN, `DRY RUN: ${folders.length} folder(s) would be created first: ${folders.map(f => f._folderPath).join(", ")}`);
    for (const item of items.slice(0, 3)) {
      const { values, errs } = await buildFormValues(item);
      writeStatus(FN, `DRY RUN sample (source Id ${item.Id}${item._dir ? `, folder ${item._dir}` : ""}): ${values.map(v => `${v.FieldName}=${JSON.stringify(v.FieldValue)}`).join(" · ")}`);
      for (const e of errs) report.fieldErrors.push({ sourceId: item.Id, ...e });
    }
    const alreadyImported = [...folders, ...items].filter(i => report.idMap[i.Id]).length;
    report.skipped = alreadyImported;
    writeStatus(FN, `DRY RUN: ${items.length - items.filter(i => report.idMap[i.Id]).length} item(s) would be created in "${targetTitle}"${alreadyImported ? `; ${alreadyImported} already imported would be skipped` : ""}.`);
    return report;
  }

  // ---- pass 0: folders --------------------------------------------------
  // A list folder must be created as a list ITEM (UnderlyingObjectType 1):
  // a folder made through web.folders.add has no item and nothing can be
  // created inside it (verified live). Parents come first (sorted by depth).
  for (const f of folders) {
    if (report.idMap[f.Id]) { report.skipped++; continue; }
    const slash = f._folderPath.lastIndexOf("/");
    const parent = slash === -1 ? rootFolder : `${rootFolder}/${f._folderPath.slice(0, slash)}`;
    const name = f._folderPath.slice(slash + 1);
    // A folder row carries metadata like any item: same conversion path, Title = folder name.
    const built = await buildFormValues(f);
    for (const e of built.errs) report.fieldErrors.push({ sourceId: f.Id, ...e });
    const formValues = [{ FieldName: "Title", FieldValue: name }, ...built.values.filter(v => v.FieldName !== "Title")];
    try {
      const result = await retryWithBackoff(() => spPostJson(`${listPathById(found[0].Id)}/AddValidateUpdateItemUsingPath`, {
        listItemCreateInfo: { FolderPath: { DecodedUrl: parent }, UnderlyingObjectType: 1, LeafName: { DecodedUrl: name } },
        formValues,
        bNewDocumentUpdate: preserveAuthorship
      }));
      const rows = formResultRows(result);
      const newId = Number(rows.find(r => /^id$/i.test(r.FieldName))?.FieldValue) || null;
      const rejected = rows.filter(r => r.HasException);
      for (const r of rejected) report.fieldErrors.push({ sourceId: f.Id, field: r.FieldName, message: r.ErrorMessage });
      if (!newId) throw new Error(rejected.map(r => `${r.FieldName}: ${r.ErrorMessage}`).join("; ") || "server did not return the folder's item Id");
      report.idMap[f.Id] = newId;
      report.folders.created++;
    } catch (err) {
      if (/already exists/i.test(err.message)) {
        // Prove it is a real list folder (one with an item) before counting it.
        try {
          const info = await spGetJson(`GetFolderByServerRelativePath(decodedUrl='${encodeURIComponent(`${parent}/${name}`)}')/ListItemAllFields?$select=Id`);
          if (info?.Id) { report.idMap[f.Id] = info.Id; report.folders.created++; continue; }
        } catch {}
        report.folders.failed++;
        warn(`Folder "${f._folderPath}" exists but has no list item; its items cannot be created inside it`);
        continue;
      }
      report.folders.failed++;
      warn(`Folder "${f._folderPath}" could not be created (${err.message}); its items will fail unless the folder exists`);
    }
  }
  if (folders.length) writeStatus(FN, `Pass 0: folders created ${report.folders.created}/${folders.length}${report.skipped ? ` (${report.skipped} already imported)` : ""}.`);

  // ---- pass 1: create -------------------------------------------------
  await runPool(items, concurrency, async (item) => {
    if (report.idMap[item.Id]) { report.skipped++; return; }
    try {
      const built = await buildFormValues(item);
      let formValues = built.values;
      for (const e of built.errs) report.fieldErrors.push({ sourceId: item.Id, ...e });
      const folderPath = folderPathFor(item);
      const addOnce = () => retryWithBackoff(() => spPostJson(`${listPathById(found[0].Id)}/AddValidateUpdateItemUsingPath`, {
        listItemCreateInfo: { FolderPath: { DecodedUrl: folderPath }, UnderlyingObjectType: 0 },
        formValues,
        bNewDocumentUpdate: preserveAuthorship
      }));
      let rows = formResultRows(await addOnce());
      let newId = Number(rows.find(r => /^id$/i.test(r.FieldName))?.FieldValue) || null;
      let rejected = rows.filter(r => r.HasException);
      for (const r of rejected) report.fieldErrors.push({ sourceId: item.Id, field: r.FieldName, message: r.ErrorMessage });
      // A rejected field aborts the whole create. Retry once without the
      // rejected fields so the item still lands; the errors stay in the report.
      if (!newId && rejected.length) {
        const drop = new Set(rejected.map(r => r.FieldName));
        formValues = formValues.filter(v => !drop.has(v.FieldName));
        rows = formResultRows(await addOnce());
        newId = Number(rows.find(r => /^id$/i.test(r.FieldName))?.FieldValue) || null;
        rejected = rows.filter(r => r.HasException);
        for (const r of rejected) report.fieldErrors.push({ sourceId: item.Id, field: r.FieldName, message: r.ErrorMessage });
        if (newId) warn(`Item ${item.Id} → ${newId}: created without ${[...drop].join(", ")} (see fieldErrors)`);
      }
      if (!newId) throw new Error(`not created: ${rejected.map(r => `${r.FieldName}: ${r.ErrorMessage}`).join("; ") || "server did not return the new item Id"}`);
      report.idMap[item.Id] = newId;
      report.created++;
      if (includeAttachments && Array.isArray(item._attachments)) {
        for (const a of item._attachments) {
          try {
            await retryWithBackoff(() => list.items.getById(newId).attachmentFiles.add(a.name, bufferFromBase64(a.base64)));
            report.attachments.added++;
          } catch (err) {
            report.attachments.failed++;
            warn(`Item ${item.Id} → ${newId}: attachment "${a.name}" failed (${err.message})`);
          }
        }
      }
    } catch (err) {
      report.failed.push({ sourceId: item.Id, error: err.message });
      writeStatus(FN, `⚠️ Item ${item.Id} failed: ${err.message}`);
    }
  });
  writeStatus(FN, `Pass 1: created ${report.created}/${items.length}${report.skipped ? `, skipped ${report.skipped} already imported` : ""}, failed ${report.failed.length}, field errors ${report.fieldErrors.length}.`);
  const allRows = [...folders, ...items];

  // ---- pass 2: self-referencing lookups -------------------------------
  if (deferredSelf.length) {
    const updates = [];
    for (const item of allRows) {
      const newId = report.idMap[item.Id];
      if (!newId) continue;
      const formValues = [];
      for (const w of deferredSelf) {
        const refs = item._resolved?.[w.name] || [];
        const targets = [];
        for (const r of refs) {
          if (report.idMap[r.Id]) targets.push(report.idMap[r.Id]);
          else report.fieldErrors.push({ sourceId: item.Id, field: w.name, message: `references source item ${r.Id}, which was not imported; reference left empty` });
        }
        if (!targets.length) continue;
        formValues.push({ FieldName: w.name, FieldValue: w.tf.TypeAsString === "LookupMulti" ? `${targets.join(";#")};#` : String(targets[0]) });
      }
      if (formValues.length) updates.push({ item, newId, formValues });
    }
    writeStatus(FN, `Pass 2: ${updates.length} item(s) with self-referencing lookups...`);
    await runPool(updates, concurrency, async ({ item, newId, formValues }) => {
      try {
        const result = await retryWithBackoff(() => spPostJson(`${listPathById(found[0].Id)}/items(${newId})/ValidateUpdateListItem`, {
          formValues,
          bNewDocumentUpdate: false
        }));
        const rows = formResultRows(result);
        for (const r of rows) if (r.HasException) report.fieldErrors.push({ sourceId: item.Id, field: r.FieldName, message: r.ErrorMessage });
      } catch (err) {
        report.fieldErrors.push({ sourceId: item.Id, field: deferredSelf.map(d => d.name).join(","), message: err.message });
      }
    });
  }

  // ---- pass 3: restore authorship ------------------------------------
  // Attachment uploads and the pass-2 update stamp Modified/Editor again, so
  // the preserved values are written last.
  if (preserveAuthorship) {
    const restores = allRows.filter(i => report.idMap[i.Id]);
    let restored = 0;
    await runPool(restores, concurrency, async (item) => {
      try {
        const formValues = await authorshipValues(item);
        if (!formValues.length) return;
        const result = await retryWithBackoff(() => spPostJson(`${listPathById(found[0].Id)}/items(${report.idMap[item.Id]})/ValidateUpdateListItem`, {
          formValues,
          bNewDocumentUpdate: true
        }));
        const bad = formResultRows(result).filter(r => r.HasException);
        for (const r of bad) report.fieldErrors.push({ sourceId: item.Id, field: r.FieldName, message: r.ErrorMessage });
        if (!bad.length) restored++;
      } catch (err) {
        report.fieldErrors.push({ sourceId: item.Id, field: "Author/Editor/Created/Modified", message: err.message });
      }
    });
    writeStatus(FN, `Pass 3: authorship restored on ${restored}/${restores.length} item(s).`);
  }

  if (report.fieldErrors.length) console.table(report.fieldErrors);
  if (report.failed.length) console.table(report.failed);
  writeStatus(FN, `Done. Created ${report.created}/${report.attempted}. Failed ${report.failed.length}. Field errors ${report.fieldErrors.length}. Attachments +${report.attachments.added}/-${report.attachments.failed}.`);
  return report;
};

/**
 * copyList
 *
 * The four primitives chained: read schema and data from the current
 * site, optionally switch site, create the target list, import the items.
 * Returns everything each step produced so a partial failure can be
 * resumed by calling the remaining steps by hand.
 *
 * @param {string} sourceTitle - List to copy (on the current context)
 * @param {object} [options]
 * @param {string}  [options.targetTitle] - Default: same title on another site, "<title> Copy" on the same site
 * @param {string}  [options.targetSite] - Site path or URL; switches setupContext() for the create/import
 * @param {boolean} [options.dryRun=false]
 * @param {boolean} [options.includeAttachments=false]
 * @param {boolean} [options.preserveAuthorship=false]
 * @param {object}  [options.lookupMap] - Source lookup list title → target list title
 * @returns {Promise<{schema:object,data:object,created:object,imported:object,targetSite:string}>}
 * @example
 * await SPUtils.copyList("Requests", { targetSite: "sites/Archive", dryRun: true });
 */
const copyList = async (sourceTitle, { targetTitle, targetSite, dryRun = false, includeAttachments = false, preserveAuthorship = false, lookupMap = {} } = {}) => {
  const FN = "copyList";
  requireContext(FN);
  const schema = await getListSchema(sourceTitle);
  const data = await exportListData(sourceTitle, { includeAttachments, schema });
  // A targetSite that resolves to the current site is a same-site copy: it
  // must not point back at the source list.
  const sameSite = !targetSite || resolveSiteUrl(targetSite).toLowerCase() === apiBase().replace(/\/+$/, "").toLowerCase();
  const title = targetTitle || (sameSite ? `${sourceTitle} Copy` : sourceTitle);
  if (sameSite && title.toLowerCase() === sourceTitle.toLowerCase()) {
    throw new Error(`Target "${title}" is the source list on the same site; choose a different targetTitle or targetSite.`);
  }
  if (targetSite && !sameSite) setupContext(targetSite);
  writeStatus(FN, `${dryRun ? "DRY RUN: " : ""}copying "${sourceTitle}" → "${title}" on ${apiBase()}`);
  const created = await createListFromSchema(schema, { title, dryRun, lookupMap });
  const imported = await importListData(data, { listTitle: title, dryRun, preserveAuthorship, includeAttachments });
  writeStatus(FN, `${dryRun ? "DRY RUN complete" : "Copy complete"}: ${imported.created}/${imported.attempted} item(s), ${created.fields.added} column(s) added.`);
  return { schema, data, created, imported, targetSite: apiBase() };
};

/**
 * downloadJson
 * Saves any object as a pretty-printed .json download (schema and data
 * documents round-trip through readJsonFile()).
 * @param {object} object
 * @param {string} [fileName="export.json"]
 */
const downloadJson = (object, fileName = "export.json") => {
  const blob = new Blob([JSON.stringify(object, null, 2)], { type: "application/json" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName.endsWith(".json") ? fileName : `${fileName}.json`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  writeStatus("downloadJson", `Downloaded ${link.download}`);
};

/**
 * readJsonFile
 * Reads and parses a .json file stored in a library on the current site.
 * @param {string} filePath - Server-relative path, e.g. "/sites/X/Shared Documents/Requests.schema.json"
 * @returns {Promise<any>}
 */
const readJsonFile = async (filePath) => {
  const FN = "readJsonFile";
  requireContext(FN);
  const text = await pnp2.sp.web.getFileByServerRelativePath(filePath).getText();
  const parsed = JSON.parse(text.charCodeAt(0) === 0xFEFF ? text.slice(1) : text);
  writeStatus(FN, `Read ${filePath}${parsed?.kind ? ` (${parsed.kind})` : ""}`);
  return parsed;
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
      getListSchemaForMigration,
      // list copy primitives
      getListSchema,
      exportListData,
      createListFromSchema,
      importListData,
      copyList,
      downloadJson,
      readJsonFile
    };

  })();
  