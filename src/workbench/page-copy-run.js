// Page copy — run executor and journal (design/PAGE-COPY.md §3, §7).
//
// This module owns *when* things happen, never *what* they mean: the plan
// (page-copy.js analyzeCopy) is frozen input, sp-pages.js is the only thing
// that talks to SharePoint, and every decision about carry sets, asset
// identity and content rewriting already happened upstream. What's left is
// sequencing the writes in the order §3 specifies, keeping a journal of what
// is CONFIRMED (a creating request that returned success — never inferred),
// and stopping the instant a response is lost so nothing gets re-created or
// deleted by guesswork.
//
// Failure model (§7): sp-write.js's post() already retries a 403 once (fresh
// digest) and 429/503 once, before any success — that is the only automatic
// retry in the whole path, and it can never double-create anything. A
// `SpFileError` with `code: 'network'` is what a lost response looks like
// (fetch threw, or the read/write never got an HTTP status at all); it is
// the ONE signal this module treats as "we don't know what happened" and
// answers with outcome 'unknown' — stop immediately, no retry, no cleanup
// attempt beyond what the operator asks for via discardCopy(). Every other
// failure in a step that CREATES or MOVES the page (create/name/assets/
// content/move/reset) is 'failed' and also stops the run — those steps only
// make sense in order. A failure in metadata/comments/publish/checkin/verify
// is recorded and the run continues; the worst those can do is leave the
// copy imperfect, never lost.

const STOP_ON_FAILURE = new Set(['create', 'name', 'assets', 'content', 'move', 'reset']);

function originOf(url) {
  try { return new URL(url).origin; } catch { return ''; }
}

function parentOf(path) {
  const s = String(path || '');
  const i = s.lastIndexOf('/');
  return i === -1 ? '' : s.slice(0, i);
}

function basename(path) {
  const s = String(path || '');
  const i = s.lastIndexOf('/');
  return i === -1 ? s : s.slice(i + 1);
}

function isNotFound(err) {
  return err?.status === 404 || err?.code === 'not-found';
}

// Pushes a step record the instant it settles (never before — a step that
// hasn't finished isn't in the journal yet) and hands it to the caller's
// onStep. A UI error in that callback must never take the run down with it.
function pushStep(journal, onStep, name, status, detail = '') {
  const record = { name, status, detail: String(detail || '') };
  journal.steps.push(record);
  if (onStep) {
    try { onStep(record, journal); } catch { /* a reporting error is not a run error */ }
  }
  return record;
}

// Runs one step's work, settles it into the journal, and reports what the
// caller needs to decide what happens next: `err` set means it didn't
// succeed, `network` means the response was lost (the one case every step,
// whatever else it does on failure, must stop for). `work` returns the
// detail string for a successful step.
async function attempt(journal, onStep, name, work) {
  try {
    const detail = await work();
    return { record: pushStep(journal, onStep, name, 'done', detail || ''), err: null, network: false };
  } catch (err) {
    const network = err?.code === 'network';
    const record = pushStep(journal, onStep, name, network ? 'unknown' : 'failed', err?.message || String(err));
    return { record, err, network };
  }
}

function pageUrlOf(journal, plan) {
  if (!journal.currentPath) return '';
  return `${originOf(plan.target.webUrl)}${journal.currentPath}`;
}

export function newJournal() {
  return {
    pageId: null,
    currentPath: '',
    createdBy: 'unknown',
    assets: [],
    steps: [],
    drift: [],
    fieldErrors: {},
    notes: ['A 403 on a write is retried once with a fresh digest before any success; nothing else is retried.'],
  };
}

// frozen = { plan } (page-copy.js analyzeCopy's output, read-only). deps =
// { source: { pages }, target: { pages, write }, rewrite(transferResults) →
// saveFields, verify(dto, savedFields) → drift[] }.
export async function runCopy(frozen, deps, { onStep } = {}) {
  const { plan } = frozen;
  const { source, target } = deps;
  const { pages, write } = target;
  const journal = newJournal(plan);
  const transferResults = new Map();
  let dto = null;
  let saved = null;
  let warnings = false;

  const finish = (outcome) => ({ outcome, journal, pageUrl: pageUrlOf(journal, plan) });

  // ---- create --------------------------------------------------------
  if (plan.engine === 'copyFile') {
    const r = await attempt(journal, onStep, 'create', async () => {
      await pages.copyFileByPath(plan.source.absoluteUrl, plan.target.finalAbsolute, { overwrite: false });
      // The file exists the instant CopyFileByPath returns — confirmed,
      // whatever happens to the id lookup that follows.
      journal.currentPath = plan.target.finalPath;
      journal.createdBy = 'this run';
      journal.pageId = await pages.fileItemId(plan.target.finalPath);
      return plan.target.finalPath;
    });
    if (r.err) return finish(r.network ? 'unknown' : 'failed');
  } else {
    const r = await attempt(journal, onStep, 'create', async () => {
      const created = await pages.createPage({ pageLayoutType: plan.pageLayoutType, promotedState: 0 });
      journal.pageId = created.Id;
      journal.createdBy = 'this run';
      const url = String(created.Url || '').replace(/^\/+/, '');
      journal.currentPath = url ? `${plan.target.webServerRelativeUrl}/${url}` : '';
      dto = created;
      return `page ${created.Id} created`;
    });
    if (r.err) return finish(r.network ? 'unknown' : 'failed');
  }

  // ---- name / assets / content / move (path A only) -------------------
  if (plan.engine === 'copyFile') {
    // §3's "reset" replaces name/assets/content/move for a whole-file copy:
    // the copy already carries the canvas, layout, banner and comments
    // setting, so all that is left is the title (the copy landed with the
    // source's) and, when the source was promoted, clearing that before
    // anyone can find it as news under the wrong title.
    const r = await attempt(journal, onStep, 'reset', async () => {
      const copied = await pages.getPage(journal.pageId);
      // Neutralise first: a checkout or title save that fails must not leave
      // a promoted duplicate behind (§4.2).
      const promoted = Number(copied.PromotedState || 0) > 0;
      if (promoted) {
        await write.validateUpdateListItem(
          { listId: plan.target.libraryId, itemId: journal.pageId },
          [{ FieldName: 'PromotedState', FieldValue: '0' }],
        );
      }
      if (!copied.IsPageCheckedOutToCurrentUser) await pages.checkoutPage(journal.pageId);
      await pages.savePage(journal.pageId, { Title: plan.title });
      // The file copy carried the content, so there is no saved body to
      // compare against — verify checks identity (title, name, layout) only.
      saved = null;
      return promoted ? 'promoted state cleared; title reset' : 'title reset';
    });
    if (r.err) return finish(r.network ? 'unknown' : 'failed');
    pushStep(journal, onStep, 'metadata', 'skipped', 'carried by the file copy');
    pushStep(journal, onStep, 'comments', 'skipped', 'carried by the file copy');
  } else {
    {
      const r = await attempt(journal, onStep, 'name', async () => {
        if (!dto.IsPageCheckedOutToCurrentUser) await pages.checkoutPage(journal.pageId);
        await pages.savePage(journal.pageId, { Title: plan.target.stagingStem });
        dto = await pages.getPage(journal.pageId);
        journal.currentPath = `${plan.target.libraryRoot}/${dto.FileName}`;
        const expected = `${plan.target.stagingStem}.aspx`;
        return String(dto.FileName || '').toLowerCase() === expected.toLowerCase()
          ? dto.FileName
          : `named ${dto.FileName} (expected ${expected})`;
      });
      if (r.err) return finish(r.network ? 'unknown' : 'failed');
    }

    // ---- assets --------------------------------------------------------
    const toTransfer = plan.assets.filter((a) => a.destName);
    if (!toTransfer.length) {
      pushStep(journal, onStep, 'assets', 'skipped', 'nothing to transfer');
    } else {
      const r = await attempt(journal, onStep, 'assets', async () => {
        // Folders/AddUsingPath refuses a folder that exists (spike §11), so
        // each level is probed first; only a folder THIS run created is
        // journaled, and only it can ever be recycled.
        for (const entry of plan.assetFolderChain) {
          if (await pages.folderExists(entry.path)) continue;
          const folder = await write.createFolder(parentOf(entry.path), basename(entry.path));
          journal.assets.push({
            path: folder.serverRelativeUrl || entry.path, kind: 'folder', confirmed: true, recyclable: entry.recyclable,
          });
        }
        for (const asset of toTransfer) {
          // A failed read or upload here fails the whole run (§5.3) — a
          // retained source reference is never silent, so this never
          // downgrades to a warning the way metadata/comments do.
          const bytes = await source.pages.readFileBytes(asset.sourcePath);
          const uploaded = await write.uploadFile(plan.assetFolder, asset.destName, bytes.bytes, { overwrite: false });
          journal.assets.push({ path: uploaded.serverRelativeUrl, kind: 'file', confirmed: true, recyclable: true });
          if (uploaded.checkOutType !== undefined && uploaded.checkOutType !== 2) {
            await write.checkInFile(uploaded.serverRelativeUrl);
          }
          const info = await pages.fileInfo(uploaded.serverRelativeUrl);
          transferResults.set(asset.key, {
            path: uploaded.serverRelativeUrl,
            ids: {
              siteId: info?.SiteId, webId: info?.WebId, listId: info?.ListId, uniqueId: info?.UniqueId,
            },
            url: `${originOf(plan.target.webUrl)}${uploaded.serverRelativeUrl}`,
          });
        }
        return `${toTransfer.length} asset(s) transferred`;
      });
      if (r.err) return finish(r.network ? 'unknown' : 'failed');
    }

    // ---- content ---------------------------------------------------------
    {
      const r = await attempt(journal, onStep, 'content', async () => {
        dto = await pages.getPage(journal.pageId);
        if (!dto.IsPageCheckedOutToCurrentUser) await pages.checkoutPage(journal.pageId);
        saved = await deps.rewrite(transferResults);
        await pages.savePage(journal.pageId, saved);
        return 'content saved';
      });
      if (r.err) return finish(r.network ? 'unknown' : 'failed');
    }

    // ---- move (staging name → final name/folder) ------------------------
    {
      const r = await attempt(journal, onStep, 'move', async () => {
        const srcAbsolute = `${originOf(plan.target.webUrl)}${journal.currentPath}`;
        try {
          await pages.moveFileByPath(srcAbsolute, plan.target.finalAbsolute, { overwrite: false });
        } catch (err) {
          // Naming the staged path is what lets the operator (or
          // discardCopy, which still finds it via journal.currentPath —
          // untouched below because the move never landed) find the orphan.
          const decorated = new Error(`${err?.message || err} (staged at ${plan.target.stagingPath})`);
          decorated.code = err?.code;
          decorated.status = err?.status;
          throw decorated;
        }
        journal.currentPath = plan.target.finalPath;
        try {
          dto = await pages.getPage(journal.pageId);
        } catch (err) {
          if (!isNotFound(err)) throw err;
          journal.pageId = await pages.fileItemId(plan.target.finalPath);
          dto = null;
        }
        return plan.target.finalPath;
      });
      if (r.err) return finish(r.network ? 'unknown' : 'failed');
    }

    // ---- metadata (§5.5 carry set; per-field, never atomic) -------------
    if (!plan.metadata.formValues || !plan.metadata.formValues.length) {
      pushStep(journal, onStep, 'metadata', 'skipped', 'no metadata to carry');
    } else {
      const r = await attempt(journal, onStep, 'metadata', async () => {
        await write.validateUpdateListItem(
          { listId: plan.target.libraryId, itemId: journal.pageId }, plan.metadata.formValues,
        );
        return `${plan.metadata.formValues.length} field(s) written`;
      });
      if (r.err) {
        if (r.network) return finish('unknown');
        if (r.err.fieldErrors) Object.assign(journal.fieldErrors, r.err.fieldErrors);
        warnings = true;
      }
    }

    // ---- comments ---------------------------------------------------------
    if (plan.commentsDisabled === true) {
      const r = await attempt(journal, onStep, 'comments', async () => {
        await pages.setCommentsDisabled(plan.target.libraryId, journal.pageId, true);
        return 'comments disabled';
      });
      if (r.err) {
        if (r.network) return finish('unknown');
        warnings = true;
      }
    } else {
      pushStep(journal, onStep, 'comments', 'skipped', 'source allows comments');
    }
  }

  // ---- publish (both engines) -------------------------------------------
  let published = false;
  if (plan.publish) {
    const r = await attempt(journal, onStep, 'publish', async () => {
      await pages.publishPage(journal.pageId);
      published = true;
      if (plan.promoteAsNews) await pages.promoteToNews(journal.pageId);
      return plan.promoteAsNews ? 'published; promoted to news' : 'published';
    });
    if (r.err) {
      if (r.network) return finish('unknown');
      warnings = true;
      // The publish itself failed, so the page stays a draft — but the
      // operator's promote choice must survive it: mark the draft the same
      // way the not-publishing branch below does, so publishing it later is
      // what makes it news.
      if (!published && plan.promoteAsNews) {
        const p = await attempt(journal, onStep, 'promote', async () => {
          await write.validateUpdateListItem(
            { listId: plan.target.libraryId, itemId: journal.pageId },
            [{ FieldName: 'PromotedState', FieldValue: '1' }],
          );
          return 'publish failed; promoted on publish';
        });
        if (p.err && p.network) return finish('unknown');
      }
    }
  } else if (plan.promoteAsNews) {
    // Not publishing yet, but the operator wants it promoted the moment it
    // is: set PromotedState now so publishing later is what makes it news.
    const r = await attempt(journal, onStep, 'publish', async () => {
      await write.validateUpdateListItem(
        { listId: plan.target.libraryId, itemId: journal.pageId },
        [{ FieldName: 'PromotedState', FieldValue: '1' }],
      );
      return 'promoted on publish';
    });
    if (r.err) {
      if (r.network) return finish('unknown');
      warnings = true;
    }
  } else {
    pushStep(journal, onStep, 'publish', 'skipped', 'not publishing');
  }

  // ---- checkin ------------------------------------------------------------
  // A publish that failed leaves the page checked out; check it in as the
  // draft it still is rather than leave it held.
  if (!published && plan.checkInDraft) {
    const r = await attempt(journal, onStep, 'checkin', async () => {
      await write.checkInFile(journal.currentPath, { comment: 'Copied with DCSPad' });
      return 'checked in';
    });
    if (r.err) {
      if (r.network) return finish('unknown');
      warnings = true;
    }
  } else {
    pushStep(journal, onStep, 'checkin', 'skipped', published ? 'published' : 'left checked out');
  }

  // ---- verify ---------------------------------------------------------------
  {
    const r = await attempt(journal, onStep, 'verify', async () => {
      const finalDto = await pages.getPage(journal.pageId);
      journal.drift = deps.verify(finalDto, saved) || [];
      return journal.drift.length ? `${journal.drift.length} field(s) drifted` : 'matches';
    });
    if (r.err) {
      if (r.network) return finish('unknown');
      warnings = true;
    } else if (journal.drift.length) {
      warnings = true;
    }
  }

  return finish(warnings ? 'done-with-warnings' : 'done');
}

// Cleans up only what THIS run confirmed creating — never a page the
// operator merely opened, never anything inferred from a name. Acts even
// when the run stopped mid-way with steps left 'unknown': whatever made it
// into journal.assets/currentPath as confirmed is fair game; whatever didn't
// is left alone, because it was never confirmed.
export async function discardCopy(journal, { target } = {}) {
  const { pages } = target;
  const recycled = [];
  const leftovers = [];
  if (!journal || journal.createdBy !== 'this run') return { recycled, leftovers };

  // Reading by a confirmed id is not inference — createPage returned this
  // id, so asking SharePoint what it is now is just catching up on a path
  // the run never got to learn, or one that changed under it.
  async function pathFromId() {
    try {
      const [dto, library] = await Promise.all([pages.getPage(journal.pageId), pages.sitePagesLibrary()]);
      const root = String(library?.rootPath || '').replace(/\/+$/, '');
      const name = dto?.FileName || '';
      if (root && name) return `${root}/${name}`;
      leftovers.push({ path: `(page id ${journal.pageId})`, error: 'its current file name could not be read' });
    } catch (err) {
      leftovers.push({ path: `(page id ${journal.pageId})`, error: err?.message || String(err) });
    }
    return '';
  }

  let currentPath = journal.currentPath;
  if (journal.pageId) {
    if (!currentPath) {
      // dto.Url absent at create, and the run stopped before 'name' read it.
      currentPath = await pathFromId();
    } else {
      // The staging save renames the file, and the run learns the new name
      // only from the read that follows it. If that read failed, currentPath
      // still holds the old name, now free for another page to take. A
      // remembered path is recycled only while it still names this run's
      // item; otherwise the page is found again by its id.
      let idAtPath = null;
      try {
        idAtPath = await pages.fileItemId(currentPath);
      } catch (err) {
        leftovers.push({ path: currentPath, error: err?.message || String(err) });
        currentPath = '';
      }
      if (currentPath && Number(idAtPath) !== Number(journal.pageId)) currentPath = await pathFromId();
    }
  }

  if (currentPath) {
    try {
      await pages.recycleFile(currentPath);
      recycled.push(currentPath);
    } catch (err) {
      leftovers.push({ path: currentPath, error: err?.message || String(err) });
    }
  }

  const fileAssets = (journal.assets || []).filter((a) => a.kind === 'file' && a.confirmed);
  for (const asset of fileAssets) {
    try {
      await pages.recycleFile(asset.path);
      recycled.push(asset.path);
    } catch (err) {
      leftovers.push({ path: asset.path, error: err?.message || String(err) });
    }
  }

  // Deepest first: a parent folder recycled before its child would take the
  // child with it and turn the child's own recycle into a spurious failure.
  const folderAssets = (journal.assets || [])
    .filter((a) => a.kind === 'folder' && a.confirmed && a.recyclable)
    .sort((a, b) => b.path.split('/').length - a.path.split('/').length);
  for (const folder of folderAssets) {
    try {
      // Creating a folder doesn't make this run the owner of everything
      // later put in it: another copy of a same-named page shares the asset
      // folder, and an upload whose response was lost is not this run's to
      // remove. Only an empty folder is recycled; anything else is left and
      // reported.
      const remaining = await pages.folderItemCount(folder.path);
      if (remaining === null) continue;   // already gone — nothing to do
      if (remaining > 0) {
        leftovers.push({ path: folder.path, error: `not recycled: it still holds ${remaining} item(s) this run did not confirm creating` });
        continue;
      }
      await pages.recycleFolder(folder.path);
      recycled.push(folder.path);
    } catch (err) {
      leftovers.push({ path: folder.path, error: err?.message || String(err) });
    }
  }

  return { recycled, leftovers };
}
