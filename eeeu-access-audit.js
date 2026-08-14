



// ============================================================================
// SHAREPOINT EEEU / BROAD-ACCESS AUDIT
//
// Modes:
//   "groups"
//       Checks all SharePoint groups on the specified site and reports groups
//       containing one or more target principals.
//
//   "content"
//       Checks all visible document libraries on the specified site.
//       Reports:
//         - Libraries assigned a target principal
//         - Folders with unique permissions assigned a target principal
//         - Files with unique permissions assigned a target principal
//         - Parent folders containing matching files
//
// Default target principals:
//   - Everyone except external users
//   - Everyone
//   - SharePoint EEEU Visitors
//
// Requires:
//   - The custom PnPjs v2 bundle exposed as window.pnp2
//   - The current user must have access to the target SharePoint site
//   - For complete results, the current user must be able to enumerate groups,
//     libraries, folders, files, and role assignments
//
// Usage examples:
//
//   await auditSharePointAccess(
//       "https://bmo.sharepoint.com/sites/FCUPortal",
//       "groups"
//   );
//
//   await auditSharePointAccess(
//       "https://bmo.sharepoint.com/sites/FCUPortal",
//       "content"
//   );
//
//   await auditSharePointAccess(
//       "https://bmo.sharepoint.com/sites/FCUPortal",
//       "content",
//       [
//           "Everyone except external users",
//           "Everyone",
//           "SharePoint EEEU Visitors",
//           "FCUDocumentation Viewers"
//       ]
//   );
//
// Returns:
//   {
//       siteUrl,
//       mode,
//       targets,
//       results,
//       nonMatchingGroups,
//       errors
//   }
//
// ============================================================================

async function auditSharePointAccess(
    siteUrl,
    mode,
    targetPrincipals = [
        "Everyone except external users",
        "Everyone",
        "SharePoint EEEU Visitors"
    ]
) {

    // ------------------------------------------------------------------------
    // Configuration
    // ------------------------------------------------------------------------



    const VALID_MODES = [
        "groups",
        "content"
    ];

    function normalizeSiteUrl(siteUrl) {

        let url = String(siteUrl || "")
            .trim()
            .replace(/\/+$/, "");

        if (!/^https?:\/\//i.test(url)) {

            url = "https://bmo.sharepoint.com" +
                (url.startsWith("/") ? "" : "/") +
                url;
        }

        return url;
    }

        let normalizedSiteUrl = normalizeSiteUrl(siteUrl);

        pnp2.sp.setup({
        sp: {
            baseUrl: normalizedSiteUrl
        }
        });

        const siteSp = pnp2.sp;

        const currentWeb = await siteSp.web
        .select("Title", "Url")
        .get();

    console.log(
        "Connected to:",
        currentWeb.Url
    );


    const normalizedMode = String(mode || "")
        .trim()
        .toLowerCase();

    const normalizedTargets = Array.from(
        new Set(
            (Array.isArray(targetPrincipals)
                ? targetPrincipals
                : [targetPrincipals]
            )
                .map(value => String(value || "").trim())
                .filter(Boolean)
        )
    );

    const targetLookup = new Map(
        normalizedTargets.map(target => [
            normalizePrincipalName(target),
            target
        ])
    );

    const results = [];
    const nonMatchingGroups = [];
    const errors = [];

    let outputContainer = null;
    let statusElement = null;

    // ------------------------------------------------------------------------
    // Validation
    // ------------------------------------------------------------------------

    if (!window.pnp2 || !pnp2.sp) {
        throw new Error(
            "The pnp2 object was not found. Load the custom PnPjs v2 bundle before running this audit."
        );
    }

    if (!normalizedSiteUrl) {
        throw new Error(
            "A SharePoint site URL is required."
        );
    }

    try {
        new URL(normalizedSiteUrl);
    } catch {
        throw new Error(
            `The SharePoint site URL is not valid: ${normalizedSiteUrl}`
        );
    }

    if (!VALID_MODES.includes(normalizedMode)) {
        throw new Error(
            `Invalid mode "${mode}". Use "groups" or "content".`
        );
    }

    if (!normalizedTargets.length) {
        throw new Error(
            "At least one target principal must be supplied."
        );
    }

    // ------------------------------------------------------------------------
    // Output UI
    // ------------------------------------------------------------------------

    outputContainer = createOutputContainer(
        normalizedSiteUrl,
        normalizedMode,
        normalizedTargets
    );

    statusElement = outputContainer.querySelector(
        "[data-sp-audit-status]"
    );

    setStatus(
        `Connecting to ${normalizedSiteUrl}...`
    );

    try {

        // --------------------------------------------------------------------
        // Configure PnPjs for the supplied site URL
        //
        // This assumes the pnp2 bundle exposes the standard PnPjs v2 singleton.
        // --------------------------------------------------------------------



        // Prove that the user can reach the site before performing the audit.

        const webInfo = await siteSp.web
            .select(
                "Id",
                "Title",
                "Url"
            )
            .get();


        setStatus(
            `Connected to ${webInfo.Title || normalizedSiteUrl}. Starting ${normalizedMode} audit...`
        );

        // --------------------------------------------------------------------
        // Run selected audit
        // --------------------------------------------------------------------

        if (normalizedMode === "groups") {

            await auditGroups({
                siteSp,
                targetLookup,
                results,
                nonMatchingGroups,
                errors,
                setStatus
            });

        } else if (normalizedMode === "content") {

            await auditContent({
                siteSp,
                targetLookup,
                results,
                errors,
                setStatus
            });

        }

    } catch (error) {

        addError({
            errors,
            scope: "Site",
            location: normalizedSiteUrl,
            operation: "Connect to site",
            error
        });

    }

    // ------------------------------------------------------------------------
    // Render final tables
    // ------------------------------------------------------------------------

    if (normalizedMode === "groups") {

        renderGroupAuditTables(
            outputContainer,
            results,
            nonMatchingGroups,
            errors
        );

    } else {

        renderResultsTable(
            outputContainer,
            normalizedMode,
            results
        );

        renderErrorsTable(
            outputContainer,
            errors
        );

    }

    const resultCount = results.length;
    const nonMatchCount = nonMatchingGroups.length;
    const errorCount = errors.length;

    const completionMessage =
        normalizedMode === "groups"
            ? (
                `Audit complete. ` +
                `${resultCount} matching membership${resultCount === 1 ? "" : "s"}, ` +
                `${nonMatchCount} successfully checked non-matching group${nonMatchCount === 1 ? "" : "s"}, ` +
                `and ${errorCount} group error${errorCount === 1 ? "" : "s"}.`
            )
            : (
                `Audit complete. Found ${resultCount} matching result${resultCount === 1 ? "" : "s"} ` +
                `and ${errorCount} error${errorCount === 1 ? "" : "s"}.`
            );

    setStatus(
        completionMessage,
        errorCount ? "warning" : "success"
    );

    console.log("SharePoint access audit results:", results);

    if (errors.length) {
        console.warn(
            "SharePoint access audit errors:",
            errors
        );
    }

    return {
        siteUrl: normalizedSiteUrl,
        mode: normalizedMode,
        targets: normalizedTargets,
        results,
        nonMatchingGroups,
        errors
    };

    // ========================================================================
    // GROUP AUDIT
    // ========================================================================

    async function auditGroups({
        siteSp,
        targetLookup,
        results,
        nonMatchingGroups,
        errors,
        setStatus
    }) {

        let groups = [];

        try {

            groups = await siteSp.web.siteGroups
                .select(
                    "Id",
                    "Title",
                    "Description",
                    "LoginName",
                    "OwnerTitle",
                    "OnlyAllowMembersViewMembership"
                )
                .get();

        } catch (error) {

            addError({
                errors,
                scope: "Site groups",
                location: normalizedSiteUrl,
                operation: "Enumerate SharePoint groups",
                error
            });

            return;
        }

        setStatus(
            `Checking ${groups.length} SharePoint group${groups.length === 1 ? "" : "s"}...`
        );

        let groupNumber = 0;

        for (const group of groups) {

            groupNumber++;

            setStatus(
                `Checking group ${groupNumber} of ${groups.length}: ${group.Title}`
            );

            try {

                const members = await siteSp.web.siteGroups
                    .getById(group.Id)
                    .users
                    .select(
                        "Id",
                        "Title",
                        "LoginName",
                        "PrincipalType",
                        "Email"
                    )
                    .get();

                const matches = findPrincipalMatches(
                    members,
                    targetLookup
                );

                if (matches.length) {

                    for (const match of matches) {

                        results.push({
                            Type: "SharePoint group",
                            Group: group.Title || "",
                            GroupId: group.Id || "",
                            Owner: group.OwnerTitle || "",
                            MembersOnlyVisibility:
                                group.OnlyAllowMembersViewMembership
                                    ? "Yes"
                                    : "No",
                            Principal: match.Title || "",
                            LoginName: match.LoginName || "",
                            Permission: "Group member",
                            MemberCount: members.length,
                            URL: buildGroupMembershipUrl(
                                normalizedSiteUrl,
                                group.Id
                            ),
                            Notes:
                                `Target principal is a member of the SharePoint group "${group.Title}".`
                        });

                    }

                } else {

                    nonMatchingGroups.push({
                        Type: "SharePoint group",
                        Group: group.Title || "",
                        GroupId: group.Id || "",
                        Owner: group.OwnerTitle || "",
                        MembersOnlyVisibility:
                            group.OnlyAllowMembersViewMembership
                                ? "Yes"
                                : "No",
                        MemberCount: members.length,
                        URL: buildGroupMembershipUrl(
                            normalizedSiteUrl,
                            group.Id
                        ),
                        Notes:
                            "Group membership was checked successfully. No target principals were found."
                    });

                }

            } catch (error) {

                addError({
                    errors,
                    scope: "SharePoint group",
                    location: group.Title,
                    operation: "Read group membership",
                    error,
                    metadata: {
                        GroupId: group.Id || "",
                        Owner: group.OwnerTitle || "",
                        MembersOnlyVisibility:
                            group.OnlyAllowMembersViewMembership
                                ? "Yes"
                                : "No",
                        URL: buildGroupMembershipUrl(
                            normalizedSiteUrl,
                            group.Id
                        )
                    }
                });

            }

        }

    }

    // ========================================================================
    // LIBRARY / FOLDER / FILE AUDIT
    // ========================================================================

    async function auditContent({
        siteSp,
        targetLookup,
        results,
        errors,
        setStatus
    }) {

        let libraries = [];

        try {

            libraries = await siteSp.web.lists
                .filter(
                    "BaseTemplate eq 101 and Hidden eq false"
                )
                .select(
                    "Id",
                    "Title",
                    "Description",
                    "ItemCount",
                    "HasUniqueRoleAssignments",
                    "RootFolder/Name",
                    "RootFolder/ServerRelativeUrl"
                )
                .expand("RootFolder")
                .get();

        } catch (error) {

            addError({
                errors,
                scope: "Document libraries",
                location: normalizedSiteUrl,
                operation: "Enumerate document libraries",
                error
            });

            return;
        }

        setStatus(
            `Found ${libraries.length} visible document librar${libraries.length === 1 ? "y" : "ies"}.`
        );

        let libraryNumber = 0;

        for (const library of libraries) {

            libraryNumber++;

            const libraryName =
                library.Title ||
                library.RootFolder?.Name ||
                `Library ${library.Id}`;

            const libraryUrl =
                library.RootFolder?.ServerRelativeUrl ||
                "";

            setStatus(
                `Checking library ${libraryNumber} of ${libraries.length}: ${libraryName}`
            );

            // ----------------------------------------------------------------
            // Check the library itself
            // ----------------------------------------------------------------

            try {

                const libraryAssignments = await libraryRoleAssignments(
                    siteSp,
                    library.Id
                );

                const libraryMatches = findRoleAssignmentMatches(
                    libraryAssignments,
                    targetLookup
                );

                for (const match of libraryMatches) {

                    results.push({
                        Type: "Document library",
                        Library: libraryName,
                        Item: libraryName,
                        URL: absoluteSharePointUrl(
                            normalizedSiteUrl,
                            libraryUrl
                        ),
                        Principal: match.Principal,
                        LoginName: match.LoginName,
                        Permission: match.Permission,
                        HasUniquePermissions:
                            library.HasUniqueRoleAssignments
                                ? "Yes"
                                : "No",
                        ParentFolder: "",
                        ContainsMatchingFiles: "",
                        Notes:
                            library.HasUniqueRoleAssignments
                                ? "The library has unique permissions."
                                : "The matching library permission may be inherited from the site."
                    });

                }

            } catch (error) {

                addError({
                    errors,
                    scope: "Document library",
                    location: libraryName,
                    operation: "Read library role assignments",
                    error
                });

            }

            // ----------------------------------------------------------------
            // Get every folder and file item in the library
            // ----------------------------------------------------------------

            let items = [];

            try {

                items = await getAllPaged(
                    siteSp.web.lists
                        .getById(library.Id)
                        .items
                        .select(
                            "Id",
                            "FileRef",
                            "FileLeafRef",
                            "FSObjType",
                            "HasUniqueRoleAssignments"
                        )
                        .top(5000)
                );

            } catch (error) {

                addError({
                    errors,
                    scope: "Document library",
                    location: libraryName,
                    operation: "Enumerate folders and files",
                    error
                });

                continue;
            }

            const uniqueItems = items.filter(
                item => item.HasUniqueRoleAssignments === true
            );

            setStatus(
                `Checking ${uniqueItems.length} uniquely permissioned item${uniqueItems.length === 1 ? "" : "s"} in ${libraryName}...`
            );

            const matchingFileFolders = new Map();

            let itemNumber = 0;

            for (const item of uniqueItems) {

                itemNumber++;

                const isFolder =
                    Number(item.FSObjType) === 1;

                const itemType =
                    isFolder ? "Folder" : "File";

                const itemPath =
                    item.FileRef ||
                    item.FileLeafRef ||
                    `Item ID ${item.Id}`;

                setStatus(
                    `Checking ${libraryName}: item ${itemNumber} of ${uniqueItems.length}`
                );

                try {

                    const assignments = await itemRoleAssignments(
                        siteSp,
                        library.Id,
                        item.Id
                    );

                    const matches = findRoleAssignmentMatches(
                        assignments,
                        targetLookup
                    );

                    for (const match of matches) {

                        const parentFolder = getParentFolder(
                            item.FileRef,
                            libraryUrl
                        );

                        results.push({
                            Type: itemType,
                            Library: libraryName,
                            Item:
                                item.FileLeafRef ||
                                itemPath,
                            URL: absoluteSharePointUrl(
                                normalizedSiteUrl,
                                item.FileRef
                            ),
                            Principal: match.Principal,
                            LoginName: match.LoginName,
                            Permission: match.Permission,
                            HasUniquePermissions: "Yes",
                            ParentFolder: parentFolder,
                            ContainsMatchingFiles: "",
                            Notes: `${itemType} has the target principal assigned through unique permissions.`
                        });

                        if (!isFolder && parentFolder) {

                            const key = [
                                library.Id,
                                parentFolder,
                                normalizePrincipalName(
                                    match.Principal
                                )
                            ].join("|");

                            if (!matchingFileFolders.has(key)) {

                                matchingFileFolders.set(key, {
                                    libraryName,
                                    libraryUrl,
                                    folderPath: parentFolder,
                                    principal: match.Principal,
                                    loginName: match.LoginName,
                                    permission: match.Permission,
                                    filePaths: new Set()
                                });

                            }

                            matchingFileFolders
                                .get(key)
                                .filePaths
                                .add(item.FileRef);

                        }

                    }

                } catch (error) {

                    addError({
                        errors,
                        scope: itemType,
                        location: itemPath,
                        operation: "Read item role assignments",
                        error
                    });

                }

            }

            // ----------------------------------------------------------------
            // Add parent-folder summary rows for folders containing files
            // that matched the selected principals.
            //
            // These rows indicate containment. They do not claim the folder
            // itself has the matching permission unless a separate Folder row
            // was already produced above.
            // ----------------------------------------------------------------

            for (const folderSummary of matchingFileFolders.values()) {

                const fileCount =
                    folderSummary.filePaths.size;

                results.push({
                    Type: "Folder containing matching files",
                    Library: folderSummary.libraryName,
                    Item: getLeafName(
                        folderSummary.folderPath
                    ),
                    URL: absoluteSharePointUrl(
                        normalizedSiteUrl,
                        folderSummary.folderPath
                    ),
                    Principal: folderSummary.principal,
                    LoginName: folderSummary.loginName,
                    Permission: folderSummary.permission,
                    HasUniquePermissions: "Not evaluated by this summary row",
                    ParentFolder: getParentFolder(
                        folderSummary.folderPath,
                        folderSummary.libraryUrl
                    ),
                    ContainsMatchingFiles:
                        String(fileCount),
                    Notes: `This folder contains ${fileCount} uniquely permissioned matching file${fileCount === 1 ? "" : "s"}.`
                });

            }

        }

    }

    // ========================================================================
    // PERMISSION HELPERS
    // ========================================================================

    async function libraryRoleAssignments(
        siteSp,
        libraryId
    ) {

        return siteSp.web.lists
            .getById(libraryId)
            .roleAssignments
            .expand(
                "Member",
                "RoleDefinitionBindings"
            )
            .get();

    }

    async function itemRoleAssignments(
        siteSp,
        libraryId,
        itemId
    ) {

        return siteSp.web.lists
            .getById(libraryId)
            .items
            .getById(itemId)
            .roleAssignments
            .expand(
                "Member",
                "RoleDefinitionBindings"
            )
            .get();

    }

    function findRoleAssignmentMatches(
        assignments,
        targetLookup
    ) {

        const matches = [];

        for (const assignment of assignments || []) {

            const member =
                assignment.Member || {};

            if (!isTargetPrincipal(
                member,
                targetLookup
            )) {
                continue;
            }

            const roles =
                assignment.RoleDefinitionBindings || [];

            if (!roles.length) {

                matches.push({
                    Principal:
                        member.Title || "",
                    LoginName:
                        member.LoginName || "",
                    Permission: ""
                });

                continue;
            }

            for (const role of roles) {

                matches.push({
                    Principal:
                        member.Title || "",
                    LoginName:
                        member.LoginName || "",
                    Permission:
                        role.Name || ""
                });

            }

        }

        return matches;

    }

    function findPrincipalMatches(
        members,
        targetLookup
    ) {

        return (members || []).filter(
            member => isTargetPrincipal(
                member,
                targetLookup
            )
        );

    }

    function isTargetPrincipal(
        principal,
        targetLookup
    ) {

        const title =
            normalizePrincipalName(
                principal?.Title
            );

        const loginName =
            normalizePrincipalName(
                principal?.LoginName
            );

        return (
            targetLookup.has(title) ||
            targetLookup.has(loginName)
        );

    }

    function normalizePrincipalName(value) {

        return String(value || "")
            .trim()
            .toLowerCase();

    }

    // ========================================================================
    // PAGING
    // ========================================================================

    async function getAllPaged(query) {

        const allItems = [];

        let page = await query.getPaged();

        allItems.push(
            ...(page.results || [])
        );

        while (page.hasNext) {

            page = await page.getNext();

            allItems.push(
                ...(page.results || [])
            );

        }

        return allItems;

    }

    // ========================================================================
    // ERROR REPORTING
    // ========================================================================

    function addError({
        errors,
        scope,
        location,
        operation,
        error,
        metadata = {}
    }) {

        const normalizedError =
            normalizeError(error);

        const entry = {
            Scope: scope || "",
            Location: location || "",
            GroupId: metadata.GroupId || "",
            Owner: metadata.Owner || "",
            MembersOnlyVisibility:
                metadata.MembersOnlyVisibility || "",
            Operation: operation || "",
            Status: normalizedError.status,
            Message: normalizedError.message,
            Details: normalizedError.details,
            URL: metadata.URL || ""
        };

        errors.push(entry);

        console.warn(
            `[SharePoint audit] ${operation} failed for ${location}`,
            error
        );

    }

    function normalizeError(error) {

        const status =
            error?.status ||
            error?.statusCode ||
            error?.response?.status ||
            error?.data?.responseBody?.status ||
            "";

        const rawMessage =
            error?.message ||
            error?.data?.responseBody?.error?.message?.value ||
            error?.data?.responseBody?.message ||
            String(error || "Unknown error");

        let message = rawMessage;

        if (
            status === 401 ||
            status === 403 ||
            /access denied|unauthorized|forbidden/i.test(rawMessage)
        ) {

            message =
                "Access denied. The running user may not have permission to view this site, group, library, folder, file, or its permissions.";

        } else if (
            status === 404 ||
            /not found|does not exist/i.test(rawMessage)
        ) {

            message =
                "The requested SharePoint object was not found or is not visible to the running user.";

        } else if (
            status === 429 ||
            /throttl|too many requests/i.test(rawMessage)
        ) {

            message =
                "SharePoint throttled the request because too many requests were made.";

        }

        return {
            status: status || "",
            message,
            details: rawMessage
        };

    }

    // ========================================================================
    // URL / PATH HELPERS
    // ========================================================================

    function absoluteSharePointUrl(
        siteUrl,
        serverRelativeUrl
    ) {

        if (!serverRelativeUrl) {
            return "";
        }

        try {

            return new URL(
                serverRelativeUrl,
                siteUrl
            ).href;

        } catch {

            return serverRelativeUrl;

        }

    }

    function buildGroupMembershipUrl(
        siteUrl,
        groupId
    ) {

        return (
            `${siteUrl}/_layouts/15/people.aspx` +
            `?MembershipGroupId=${encodeURIComponent(groupId)}`
        );

    }

    function getParentFolder(
        itemPath,
        libraryRoot
    ) {

        const path = String(itemPath || "")
            .replace(/\/+$/, "");

        if (!path || path === libraryRoot) {
            return "";
        }

        const slashIndex =
            path.lastIndexOf("/");

        if (slashIndex < 1) {
            return "";
        }

        const parent =
            path.slice(0, slashIndex);

        if (
            libraryRoot &&
            parent.length < libraryRoot.length
        ) {
            return "";
        }

        return parent;

    }

    function getLeafName(path) {

        const cleaned = String(path || "")
            .replace(/\/+$/, "");

        return (
            cleaned.split("/").pop() ||
            cleaned
        );

    }

    // ========================================================================
    // OUTPUT UI
    // ========================================================================

    function createOutputContainer(
        siteUrl,
        mode,
        targets
    ) {

        const existing =
            document.getElementById(
                "sharepoint-access-audit-output"
            );

        if (existing) {
            existing.remove();
        }

        const container =
            document.createElement("div");

        container.id =
            "sharepoint-access-audit-output";

        container.style.cssText = [
            "margin:16px 0",
            "padding:16px",
            "border:1px solid #d2d0ce",
            "border-radius:8px",
            "background:#ffffff",
            "color:#242424",
            "font:14px/1.45 Segoe UI,Arial,sans-serif",
            "box-shadow:0 2px 6px rgba(0,0,0,.08)",
            "overflow:auto"
        ].join(";");

        container.innerHTML = `
            <div style="margin-bottom:16px;">
                <h2 style="margin:0 0 8px;font-size:20px;">
                    SharePoint Access Audit
                </h2>

                <div style="margin-bottom:4px;">
                    <strong>Site:</strong>
                    ${escapeHtml(siteUrl)}
                </div>

                <div style="margin-bottom:4px;">
                    <strong>Mode:</strong>
                    ${escapeHtml(mode)}
                </div>

                <div style="margin-bottom:8px;">
                    <strong>Targets:</strong>
                    ${escapeHtml(targets.join("; "))}
                </div>

                <div
                    data-sp-audit-status
                    style="
                        padding:8px 10px;
                        background:#f3f2f1;
                        border-left:4px solid #605e5c;
                    "
                >
                    Initializing...
                </div>
            </div>

            <div data-sp-audit-results></div>
            <div data-sp-audit-errors></div>
        `;

        document.body.appendChild(container);

        container.scrollIntoView({
            behavior: "smooth",
            block: "start"
        });

        return container;

    }

    function setStatus(
        message,
        state = "working"
    ) {

        if (!statusElement) {
            return;
        }

        const styles = {
            working: {
                background: "#f3f2f1",
                border: "#605e5c"
            },
            success: {
                background: "#dff6dd",
                border: "#107c10"
            },
            warning: {
                background: "#fff4ce",
                border: "#797775"
            },
            error: {
                background: "#fde7e9",
                border: "#a80000"
            }
        };

        const selected =
            styles[state] ||
            styles.working;

        statusElement.textContent = message;

        statusElement.style.background =
            selected.background;

        statusElement.style.borderLeftColor =
            selected.border;

        console.log(
            `[SharePoint audit] ${message}`
        );

    }

    function renderResultsTable(
        container,
        mode,
        rows
    ) {

        const target =
            container.querySelector(
                "[data-sp-audit-results]"
            );

        const columns =
            mode === "groups"
                ? [
                    "Type",
                    "Item",
                    "Principal",
                    "LoginName",
                    "Permission",
                    "URL",
                    "Notes"
                ]
                : [
                    "Type",
                    "Library",
                    "Item",
                    "Principal",
                    "Permission",
                    "HasUniquePermissions",
                    "ParentFolder",
                    "ContainsMatchingFiles",
                    "URL",
                    "Notes"
                ];

        target.innerHTML = `
            <h3 style="margin:20px 0 8px;font-size:17px;">
                Matching Results (${rows.length})
            </h3>
            ${buildHtmlTable(rows, columns)}
        `;

    }

    function renderGroupAuditTables(
        container,
        matchingRows,
        nonMatchingRows,
        errorRows
    ) {

        const resultsTarget =
            container.querySelector(
                "[data-sp-audit-results]"
            );

        const errorsTarget =
            container.querySelector(
                "[data-sp-audit-errors]"
            );

        const matchingColumns = [
            "Group",
            "GroupId",
            "Owner",
            "MembersOnlyVisibility",
            "Principal",
            "LoginName",
            "Permission",
            "MemberCount",
            "URL",
            "Notes"
        ];

        const nonMatchingColumns = [
            "Group",
            "GroupId",
            "Owner",
            "MembersOnlyVisibility",
            "MemberCount",
            "URL",
            "Notes"
        ];

        const errorColumns = [
            "Location",
            "GroupId",
            "Owner",
            "MembersOnlyVisibility",
            "Operation",
            "Status",
            "Message",
            "Details",
            "URL"
        ];

        resultsTarget.innerHTML = `
            ${buildCollapsibleSection(
                `Matching Groups (${matchingRows.length})`,
                buildHtmlTable(
                    matchingRows,
                    matchingColumns
                ),
                true
            )}

            ${buildCollapsibleSection(
                `Non-matching Groups (${nonMatchingRows.length})`,
                buildHtmlTable(
                    nonMatchingRows,
                    nonMatchingColumns
                ),
                false
            )}
        `;

        errorsTarget.innerHTML = `
            ${buildCollapsibleSection(
                `Error Groups (${errorRows.length})`,
                buildHtmlTable(
                    errorRows,
                    errorColumns,
                    {
                        compactErrors: true
                    }
                ),
                false,
                errorRows.length
                    ? "error"
                    : "normal"
            )}
        `;

    }

    function buildCollapsibleSection(
        title,
        content,
        openByDefault = false,
        state = "normal"
    ) {

        const borderColor =
            state === "error"
                ? "#a80000"
                : "#d2d0ce";

        const summaryBackground =
            state === "error"
                ? "#fde7e9"
                : "#f3f2f1";

        return `
            <details
                ${openByDefault ? "open" : ""}
                style="
                    margin:18px 0;
                    border:1px solid ${borderColor};
                    border-radius:6px;
                    overflow:hidden;
                    background:#ffffff;
                "
            >
                <summary
                    style="
                        padding:12px 14px;
                        cursor:pointer;
                        font-size:17px;
                        font-weight:600;
                        background:${summaryBackground};
                        user-select:none;
                    "
                >
                    ${escapeHtml(title)}
                </summary>

                <div style="padding:12px;">
                    ${content}
                </div>
            </details>
        `;

    }

    function renderErrorsTable(
        container,
        rows
    ) {

        const target =
            container.querySelector(
                "[data-sp-audit-errors]"
            );

        const columns = [
            "Scope",
            "Location",
            "Operation",
            "Status",
            "Message",
            "Details"
        ];

        target.innerHTML = `
            <h3 style="margin:20px 0 8px;font-size:17px;">
                Errors (${rows.length})
            </h3>
            ${buildHtmlTable(rows, columns)}
        `;

    }

    function buildHtmlTable(
        rows,
        columns,
        options = {}
    ) {

        if (!rows.length) {

            return `
                <div style="
                    padding:10px;
                    border:1px solid #edebe9;
                    background:#faf9f8;
                ">
                    No entries found.
                </div>
            `;

        }

        const tableId =
            `sp-audit-table-${Math.random()
                .toString(36)
                .slice(2)}`;

        const headerHtml = columns
            .map(column => `
                <th style="
                    padding:8px;
                    text-align:left;
                    vertical-align:top;
                    border:1px solid #d2d0ce;
                    background:#f3f2f1;
                    position:sticky;
                    top:0;
                    z-index:1;
                    white-space:nowrap;
                ">
                    ${escapeHtml(column)}
                </th>
            `)
            .join("");

        const bodyHtml = rows
            .map(row => {

                const cells = columns
                    .map(column => {

                        const value =
                            row[column] ?? "";

                        if (
                            options.compactErrors &&
                            (
                                column === "Message" ||
                                column === "Details"
                            ) &&
                            value
                        ) {

                            return `
                                <td
                                    class="sp-audit-error-cell"
                                    style="
                                        padding:6px 8px;
                                        vertical-align:top;
                                        border:1px solid #edebe9;
                                        width:${column === "Details" ? "340px" : "280px"};
                                        min-width:${column === "Details" ? "240px" : "200px"};
                                        max-width:${column === "Details" ? "420px" : "340px"};
                                    "
                                >
                                    <details class="sp-audit-cell-details">
                                        <summary
                                            title="Select to expand"
                                            style="
                                                cursor:pointer;
                                                list-style:none;
                                            "
                                        >
                                            <span
                                                style="
                                                    display:-webkit-box;
                                                    -webkit-box-orient:vertical;
                                                    -webkit-line-clamp:2;
                                                    overflow:hidden;
                                                    line-height:1.35;
                                                    max-height:2.7em;
                                                    overflow-wrap:anywhere;
                                                "
                                            >
                                                ${escapeHtml(value)}
                                            </span>

                                            <span
                                                aria-hidden="true"
                                                style="
                                                    display:inline-block;
                                                    color:#605e5c;
                                                    font-weight:700;
                                                    letter-spacing:2px;
                                                    margin-top:1px;
                                                "
                                            >
                                                ...
                                            </span>
                                        </summary>

                                        <div
                                            style="
                                                margin-top:6px;
                                                padding-top:6px;
                                                border-top:1px solid #edebe9;
                                                white-space:pre-wrap;
                                                overflow-wrap:anywhere;
                                                max-height:220px;
                                                overflow:auto;
                                            "
                                        >
                                            ${escapeHtml(value)}
                                        </div>
                                    </details>
                                </td>
                            `;

                        }

                        if (
                            column === "URL" &&
                            value
                        ) {

                            return `
                                <td style="
                                    padding:6px 8px;
                                    vertical-align:top;
                                    border:1px solid #edebe9;
                                    max-width:420px;
                                    overflow-wrap:anywhere;
                                ">
                                    <a
                                        href="${escapeAttribute(value)}"
                                        target="_blank"
                                        rel="noopener noreferrer"
                                    >
                                        ${escapeHtml(value)}
                                    </a>
                                </td>
                            `;

                        }

                        return `
                            <td style="
                                padding:8px;
                                vertical-align:top;
                                border:1px solid #edebe9;
                                max-width:420px;
                                overflow-wrap:anywhere;
                            ">
                                ${escapeHtml(value)}
                            </td>
                        `;

                    })
                    .join("");

                return `<tr>${cells}</tr>`;

            })
            .join("");

        return `
            <div style="overflow:auto;max-height:650px;">
                <table
                    id="${tableId}"
                    style="
                        width:100%;
                        border-collapse:collapse;
                        font-size:13px;
                    "
                >
                    <thead>
                        <tr>${headerHtml}</tr>
                    </thead>
                    <tbody>
                        ${bodyHtml}
                    </tbody>
                </table>
            </div>
        `;

    }

    function escapeHtml(value) {

        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");

    }

    function escapeAttribute(value) {

        return escapeHtml(value)
            .replace(/`/g, "&#096;");

    }

}

 auditSharePointAccess(
    "https://bmo.sharepoint.com/teams/A9EE727E/",
    "groups"
);