// SP Workbench — the page-copy dialog (design/PAGE-COPY.md §6).
//
// House style of upload-metadata.js / list-schema-dialog.js: a `<dialog
// class="app-dialog sp-metadata-dialog …">`, a `busy` guard around the run
// phase, Esc/✕ blocked while a write is in flight. The dialog does no
// planning or writing of its own — Check calls the pure `analyzeCopy()`
// (page-copy.js) against read-only target probes, Copy hands the frozen
// plan to `runCopy` (page-copy-run.js), injected by the caller so this
// module never imports it directly.
//
// The target client pair follows list-schema-dialog.js's connect() pattern
// exactly: a fresh client per connect (`createClient()` + `connectWeb()`),
// a `connectSeq` guard so only the latest connect's answer lands, and a
// `writerFor()` that rebuilds the write client from the SAME candidate —
// never the shell's shared `client`, which stays pointed at the inspected
// (source) web for the life of the dialog.

import {
  snapshotFromReads, sourceEligibility, targetEligibility, defaultFileName, fileNameProblem,
  folderProblem, fileStem, analyzeParts, assetRequests, analyzeCopy, rewriteContent,
  compareReadBack, shortRunId, stagingStem, normalizeGuid, dependentConsumers,
} from './page-copy.js';
import { createSpPages } from './sp-pages.js';
import { createSpWriteClient } from './sp-write.js';
import { createFieldEditor } from './field-editor.js';
import { getFavorites, getRecents } from './favorites.js';
import { MODERN_SITE_PAGES_FEATURE_ID } from '../bridge/sp-context.js';

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

// Same select the library-field probes elsewhere in the Workbench use
// (context pack §2 sp-write/pages.js FIELD_SELECT) — source and target read
// through the identical shape so computeCarrySet()'s type comparison means
// something.
const FIELD_SELECT = [
  'Id', 'Title', 'InternalName', 'EntityPropertyName', 'TypeAsString',
  'FieldTypeKind', 'Required', 'Hidden', 'ReadOnlyField', 'Choices', 'FillInChoice',
];

const trimSlash = (v) => String(v || '').replace(/\/+$/, '');
const lower = (v) => String(v || '').toLowerCase();

function fieldsPath(listId) {
  return `web/lists(guid'${listId}')/fields`;
}

// Candidate names in defaultFileName()'s own order — probed live against the
// destination so the pre-filled File name field reflects real collisions,
// not just the ones already known from a prior read.
function nameCandidates(stem) {
  const out = [`${stem}.aspx`, `${stem}-copy.aspx`];
  for (let i = 2; i < 1000; i += 1) out.push(`${stem}-copy-${i}.aspx`);
  return out;
}

function bytesText(n) {
  const num = Number(n) || 0;
  if (num >= 1024 * 1024) return `${(num / (1024 * 1024)).toFixed(1)} MB`;
  if (num >= 1024) return `${Math.round(num / 1024)} KB`;
  return `${num} B`;
}

const VERDICT_LABELS = {
  asset: 'asset', data: 'data', config: 'config', link: 'link',
  dynamic: 'dynamic', unavailable: 'unavailable', unverified: 'unverified',
};

function partVerdictCounts(part) {
  const counts = new Map();
  const bump = (key, n = 1) => counts.set(key, (counts.get(key) || 0) + n);
  for (const ref of part.refs || []) bump(ref.class || 'unverified');
  if (!part.available) bump('unavailable');
  if (part.unverified?.length) bump('unverified', part.unverified.length);
  return counts;
}

// Says which version the copy carries (item.OData__UIVersionString /
// _UIVersionString plus statusOf()'s derivePageStatus() output — best
// effort, so `status` may be null and the sentence falls back to the bare
// version label).
function versionLine(item, status) {
  const ver = String(item?.OData__UIVersionString ?? item?._UIVersionString ?? '').trim();
  const major = Number((ver.split('.')[0] || '').trim());
  const isDraft = Number.isFinite(major) && major === 0;
  if (!ver) return 'Copies the current version of this page.';
  if (isDraft) {
    return status?.published === true
      ? `Copies the current version (${ver} — a draft newer than the last published version).`
      : `Copies the current draft version (${ver}); this page has never been published.`;
  }
  return status?.published === true
    ? `Copies the published version ${ver}.`
    : `Copies the current version (${ver}).`;
}

export function openPageCopyDialog({
  client, createClient, mockWriter, library, pageId, pageName,
  statusOf, analyzers = null, runCopy, discardCopy,
}) {
  return new Promise((resolve) => {
    const dialog = el('dialog', 'app-dialog sp-metadata-dialog wb-page-copy-dialog');
    const panel = el('div', 'app-dialog__panel');
    const head = el('div', 'app-dialog__head');
    head.append(el('h2', '', 'Copy page'));
    const closeIcon = el('button', 'btn btn-ghost btn-xs', '✕');
    closeIcon.type = 'button';
    closeIcon.setAttribute('aria-label', 'Close');
    head.append(closeIcon);
    const context = el('p', 'app-dialog__context', `Copy “${pageName || library?.title || 'this page'}” to a new page.`);
    panel.append(head, context);
    dialog.append(panel);
    document.body.append(dialog);

    let settled = false;
    const finish = (outcome) => {
      if (settled) return;
      settled = true;
      dialog.close();
      dialog.remove();
      resolve(outcome);
    };

    let running = false; // guards the run phase only — reads (Check) never block Esc/✕
    // Once a run has written anything, every exit reports its result — ✕ and
    // Esc included — so the caller never mistakes a copy for a cancel.
    let runResult = null;
    const dismiss = () => {
      if (running) return;
      finish(runResult || 'cancelled');
    };
    closeIcon.addEventListener('click', dismiss);
    dialog.addEventListener('cancel', (e) => { e.preventDefault(); dismiss(); });

    dialog.showModal();

    // ---- phase 0: read the source snapshot ---------------------------------
    const loading = el('p', 'app-dialog__context', 'Reading page…');
    panel.append(loading);

    const sourceWrite = createSpWriteClient({ client, mockWriter });
    const sourcePages = createSpPages({ client, write: sourceWrite });

    (async () => {
      let snapshot;
      try {
        const [page, item, fieldsRes, status, web] = await Promise.all([
          sourcePages.getPage(pageId),
          client.get(`web/lists(guid'${library.listId}')/items(${Number(pageId)})`, {
            select: ['*'], expand: 'FieldValuesAsText',
          }),
          client.getAll(fieldsPath(library.listId), { select: FIELD_SELECT }),
          statusOf ? statusOf().catch(() => null) : Promise.resolve(null),
          sourcePages.webIdentity(),
        ]);
        snapshot = snapshotFromReads({
          page, item, fields: fieldsRes.items, status,
          library: {
            id: library.listId, title: library.title,
            baseTemplate: library.baseTemplate, hidden: Boolean(library.hidden),
            rootPath: library.rootPath,
          },
          web,
        });
      } catch (err) {
        loading.remove();
        renderRefusal(err?.message || String(err));
        return;
      }
      loading.remove();
      const elig = sourceEligibility(snapshot, { sameWeb: true });
      if (!elig.ok) {
        renderRefusal(elig.reason);
        return;
      }
      buildForm(snapshot);
    })();

    function renderRefusal(reason) {
      const error = el('div', 'sp-files-error', reason);
      error.setAttribute('role', 'alert');
      panel.append(error);
      const actions = el('div', 'app-dialog__actions');
      const closeBtn = el('button', 'btn btn-run wb-pc-close', 'Close');
      closeBtn.type = 'button';
      closeBtn.addEventListener('click', dismiss);
      actions.append(closeBtn);
      panel.append(actions);
    }

    // ---- phase 1: form + preflight ------------------------------------------
    function buildForm(snapshot) {
      panel.append(el('p', 'wb-pc-version', versionLine(snapshot.item, snapshot.status)));

      // -- destination ---------------------------------------------------
      const destField = el('div', 'app-dialog__field');
      destField.append(el('label', '', 'Destination'));
      const destRow = el('div', 'wb-pc-dest-row');
      const thisLabel = el('label', 'wb-pc-dest-opt');
      const thisRadio = el('input', 'wb-pc-dest-this');
      thisRadio.type = 'radio'; thisRadio.name = 'wb-pc-dest'; thisRadio.value = 'this'; thisRadio.checked = true;
      thisLabel.append(thisRadio, document.createTextNode('This site'));
      const otherLabel = el('label', 'wb-pc-dest-opt');
      const otherRadio = el('input', 'wb-pc-dest-other');
      otherRadio.type = 'radio'; otherRadio.name = 'wb-pc-dest'; otherRadio.value = 'other';
      otherLabel.append(otherRadio, document.createTextNode('Another site'));
      destRow.append(thisLabel, otherLabel);
      destField.append(destRow);

      const urlRow = el('div', 'wb-pc-url-row');
      const urlInput = el('input', 'wb-pc-url');
      urlInput.type = 'text';
      urlInput.autocomplete = 'off';
      urlInput.disabled = true;
      urlInput.setAttribute('list', 'wb-pc-url-list');
      const datalist = el('datalist');
      datalist.id = 'wb-pc-url-list';
      for (const fav of getFavorites()) {
        const opt = el('option', '', fav.title || fav.url);
        opt.value = fav.url;
        datalist.append(opt);
      }
      for (const rec of getRecents()) {
        const opt = el('option', '', rec.title || rec.url);
        opt.value = rec.url;
        datalist.append(opt);
      }
      const connectBtn = el('button', 'btn btn-xs wb-pc-connect', 'Connect');
      connectBtn.type = 'button';
      urlRow.append(urlInput, datalist, connectBtn);
      destField.append(urlRow);
      const targetStatus = el('div', 'wb-pc-target-status');
      targetStatus.hidden = true;
      destField.append(targetStatus);
      panel.append(destField);

      // -- fields ----------------------------------------------------------
      const fieldsWrap = el('div', 'wb-pc-fields');

      const nameField = el('label', 'app-dialog__field');
      nameField.append(el('span', '', 'File name'));
      const nameInput = el('input', 'wb-pc-filename');
      nameInput.type = 'text'; nameInput.autocomplete = 'off';
      nameField.append(nameInput);
      fieldsWrap.append(nameField);

      const titleField = el('label', 'app-dialog__field');
      titleField.append(el('span', '', 'Title'));
      const titleInput = el('input', 'wb-pc-title');
      titleInput.type = 'text'; titleInput.autocomplete = 'off';
      titleInput.value = snapshot.dto.Title || '';
      titleField.append(titleInput);
      fieldsWrap.append(titleField);

      const folderField = el('label', 'app-dialog__field');
      folderField.append(el('span', '', 'Folder'));
      const folderInput = el('input', 'wb-pc-folder');
      folderInput.type = 'text'; folderInput.autocomplete = 'off';
      folderInput.placeholder = '(library root)';
      folderField.append(folderInput);
      fieldsWrap.append(folderField);

      const checksRow = el('div', 'wb-pc-checks');
      const publishLabel = el('label', '');
      const publishCb = el('input', 'wb-pc-publish'); publishCb.type = 'checkbox';
      publishLabel.append(publishCb, document.createTextNode('Publish'));
      checksRow.append(publishLabel);

      const promoteRow = el('label', '');
      promoteRow.hidden = snapshot.dto.PromotedState <= 0;
      const promoteCb = el('input', 'wb-pc-promote'); promoteCb.type = 'checkbox';
      promoteRow.append(promoteCb, document.createTextNode('Promote as news'));
      checksRow.append(promoteRow);

      const carryLabel = el('label', '');
      const carryCb = el('input', 'wb-pc-carry'); carryCb.type = 'checkbox'; carryCb.checked = true;
      carryLabel.append(carryCb, document.createTextNode('Carry custom metadata'));
      checksRow.append(carryLabel);

      const rewriteLabel = el('label', '');
      rewriteLabel.hidden = true; // shown once we know the destination is cross-web
      const rewriteCb = el('input', 'wb-pc-rewrite'); rewriteCb.type = 'checkbox';
      rewriteLabel.append(rewriteCb, document.createTextNode('Rewrite links into the source site'));
      checksRow.append(rewriteLabel);

      fieldsWrap.append(checksRow);
      panel.append(fieldsWrap);

      // -- actions: Check --------------------------------------------------
      const checkRow = el('div', 'wb-pc-check-row');
      const checkBtn = el('button', 'btn btn-xs wb-pc-check', 'Check');
      checkBtn.type = 'button';
      checkRow.append(checkBtn);
      panel.append(checkRow);

      // A Drop toggle darkens Copy like any other edit, but — unlike a text
      // edit — the report it was toggled inside stays on screen; this note is
      // the only thing that changes next to it.
      const dropNote = el('p', 'wb-pc-recheck-note', 'Check again to apply.');
      dropNote.hidden = true;
      panel.append(dropNote);

      // -- report ------------------------------------------------------------
      const report = el('div', 'wb-pc-report');
      report.hidden = true;
      panel.append(report);

      const errorEl = el('div', 'sp-files-error');
      errorEl.setAttribute('role', 'alert');
      errorEl.hidden = true;
      panel.append(errorEl);

      // -- phase two: step log + outcome -------------------------------------
      const runPanel = el('div', 'wb-pc-run');
      runPanel.hidden = true;
      const stepsList = el('ol', 'wb-pc-steps');
      runPanel.append(stepsList);
      const outcomeEl = el('p', 'wb-pc-outcome');
      outcomeEl.hidden = true;
      runPanel.append(outcomeEl);
      const driftEl = el('div', 'wb-pc-drift');
      driftEl.hidden = true;
      runPanel.append(driftEl);
      panel.append(runPanel);

      // -- actions -------------------------------------------------------
      const actions = el('div', 'app-dialog__actions');
      const cancelBtn = el('button', 'btn btn-ghost wb-pc-cancel', 'Cancel');
      cancelBtn.type = 'button';
      const discardBtn = el('button', 'btn btn-ghost wb-pc-discard', 'Discard copy');
      discardBtn.type = 'button';
      discardBtn.hidden = true;
      const openLink = el('a', 'wb-pc-open', 'Open copy ↗');
      openLink.hidden = true;
      openLink.target = '_blank';
      openLink.rel = 'noopener';
      const closeBtn = el('button', 'btn btn-ghost wb-pc-close', 'Close');
      closeBtn.type = 'button';
      closeBtn.hidden = true;
      const goBtn = el('button', 'btn btn-run wb-pc-go', 'Copy');
      goBtn.type = 'button';
      goBtn.disabled = true;
      actions.append(cancelBtn, discardBtn, openLink, closeBtn, goBtn);
      panel.append(actions);

      // ---- target connect state ------------------------------------------
      let connectSeq = 0;
      let targetClient = null;
      let tp = null;
      let spWrite = null;
      let targetLib = null;
      let targetFields = [];
      let targetWebParts = null; // Set|null — resolved only for a cross-web target
      let targetWebIdentity = null;
      // Readers for runPreflight, which pins its own copies of these (a
      // same-named const there would otherwise sit in its own dead zone).
      const currentTp = () => tp;
      const currentTargetLib = () => targetLib;
      const currentTargetIdentity = () => targetWebIdentity;
      const currentTargetClient = () => targetClient;
      const currentTargetFields = () => targetFields;
      const currentTargetWebParts = () => targetWebParts;
      let connected = false;
      let targetEligible = false;
      let connectedUrl = null;

      // ---- preflight state -------------------------------------------------
      let fileNameTouched = false;
      let checking = false;
      let lastPlan = null;
      let lastKey = null;
      const requiredValues = {};
      const requiredEditors = new Map();
      // Per-part drop (design/PAGE-COPY.md §1 decision 4, §5.2 "dynamic"):
      // instance ids the operator chose to drop, and — for a dropped
      // dynamic-data provider — the consumer instance ids the operator has
      // confirmed will lose their data source. Both are lowercased instance
      // ids, matching page-copy.js's own normalization.
      const dropped = new Set();
      const confirmedConsumers = new Set();

      // Any edit invalidates synchronously (§2.3): Copy goes dark and a Check
      // already in flight discards its answer.
      const edited = () => { invalidatePreflight(); refreshGate(); };
      nameInput.addEventListener('input', () => { fileNameTouched = true; edited(); });
      for (const input of [titleInput, folderInput, publishCb, promoteCb, carryCb, rewriteCb]) {
        input.addEventListener('input', edited);
        input.addEventListener('change', edited);
      }

      function showTargetStatus(text, isError) {
        targetStatus.hidden = !text;
        targetStatus.textContent = text || '';
        targetStatus.classList.toggle('wb-pc-status-error', Boolean(isError));
      }

      function showError(text) {
        errorEl.hidden = !text;
        errorEl.textContent = text || '';
      }

      // Bumped by every invalidation (input edits, connects): a Check that
      // started under an older epoch drops its answer instead of arming Copy.
      let inputEpoch = 0;
      function invalidatePreflight() {
        inputEpoch += 1;
        lastPlan = null;
        lastKey = null;
        confirmedConsumers.clear();
        report.hidden = true;
        report.textContent = '';
        dropNote.hidden = true;
      }

      // A Drop toggle invalidates the preflight exactly like any other edit
      // (Copy goes dark, Check again is required) but the report it was
      // toggled inside is not blanked — only the gate and this note change.
      function invalidatePreflightForDrop() {
        inputEpoch += 1;
        lastPlan = null;
        lastKey = null;
        confirmedConsumers.clear();
        dropNote.hidden = false;
      }

      function computeKey() {
        return JSON.stringify({
          url: connectedUrl,
          fileName: nameInput.value.trim(),
          title: titleInput.value.trim(),
          folder: folderInput.value.trim(),
          publish: publishCb.checked,
          promote: promoteCb.checked,
          carry: carryCb.checked,
          rewrite: rewriteCb.checked,
          dropped: [...dropped].sort(),
          required: Object.keys(requiredValues).sort().map((k) => [k, requiredValues[k]]),
          etag: snapshot.etag,
        });
      }

      // Consumers of any dropped provider on the CURRENT plan, still waiting
      // on their own confirm checkbox (§5.2 "dynamic": every consumer is
      // listed and the drop is confirmed per consumer).
      function pendingConsumerConfirmations() {
        if (!lastPlan) return [];
        return dependentConsumers(lastPlan.analysis, lastPlan.dropped || [])
          .filter((c) => !confirmedConsumers.has(lower(c.instanceId)));
      }

      function refreshGate() {
        checkBtn.disabled = !connected || !targetEligible || checking || running;
        const fresh = Boolean(lastPlan) && !lastPlan.blockers.length && computeKey() === lastKey;
        const pendingConsumers = fresh && pendingConsumerConfirmations().length > 0;
        goBtn.disabled = !fresh || running || pendingConsumers;
      }

      function onDropChanged() {
        invalidatePreflightForDrop();
        refreshGate();
      }

      // ---- connect (list-schema-dialog.js's exact pattern) ------------------
      async function connect(rawUrl) {
        const seq = ++connectSeq;
        const stale = () => seq !== connectSeq;
        connected = false;
        targetEligible = false;
        connectedUrl = null;
        invalidatePreflight();
        showError('');
        showTargetStatus('Connecting…');
        connectBtn.disabled = true;
        refreshGate();

        let candidate;
        try {
          candidate = createClient();
          await candidate.connectWeb(rawUrl);
        } catch (err) {
          if (stale()) return;
          connectBtn.disabled = false;
          showTargetStatus(err?.message || String(err), true);
          refreshGate();
          return;
        }
        if (stale()) return;

        const candidateWrite = createSpWriteClient({ client: candidate, mockWriter });
        const candidateTp = createSpPages({ client: candidate, write: candidateWrite });

        let identity, featureOk, lib, fields;
        try {
          identity = await candidateTp.webIdentity();
          if (stale()) return;
          featureOk = await candidateTp.featureActive(MODERN_SITE_PAGES_FEATURE_ID);
          if (stale()) return;
          lib = await candidateTp.sitePagesLibrary();
          if (stale()) return;
          fields = lib ? (await candidate.getAll(fieldsPath(lib.id), { select: FIELD_SELECT })).items : [];
          if (stale()) return;
        } catch (err) {
          if (stale()) return;
          connectBtn.disabled = false;
          showTargetStatus(err?.message || String(err), true);
          refreshGate();
          return;
        }

        const sameWebNow = Boolean(identity.webId)
          && normalizeGuid(identity.webId) === normalizeGuid(snapshot.web.webId);
        let webParts = null;
        if (!sameWebNow) {
          try {
            const parts = await candidateTp.clientSideWebParts();
            if (stale()) return;
            webParts = new Set(parts.map((p) => normalizeGuid(p.Id)).filter(Boolean));
          } catch (err) {
            if (stale()) return;
            connectBtn.disabled = false;
            showTargetStatus(err?.message || String(err), true);
            refreshGate();
            return;
          }
        }

        const tgtElig = targetEligibility({ featureActive: featureOk, library: lib });
        const srcElig = sourceEligibility(snapshot, { sameWeb: sameWebNow });
        connectBtn.disabled = false;
        connected = true;
        connectedUrl = rawUrl;
        targetClient = candidate;
        tp = candidateTp;
        spWrite = candidateWrite;
        targetLib = lib;
        targetFields = fields;
        targetWebParts = webParts;
        targetWebIdentity = identity;
        rewriteLabel.hidden = sameWebNow;
        if (sameWebNow) rewriteCb.checked = false;

        if (!srcElig.ok) { targetEligible = false; showTargetStatus(srcElig.reason, true); }
        else if (!tgtElig.ok) { targetEligible = false; showTargetStatus(tgtElig.reason, true); }
        else {
          targetEligible = true;
          showTargetStatus('');
          if (!fileNameTouched) {
            nameInput.value = defaultFileName(snapshot.fileName, new Set());
          }
        }
        refreshGate();
      }

      thisRadio.addEventListener('change', () => {
        if (!thisRadio.checked) return;
        urlInput.disabled = true;
        urlInput.value = '';
        connect(client.webUrl());
      });
      // Changing the destination abandons any connect still in flight: its
      // answer would otherwise land after this and install a web the form no
      // longer names as the target. Bumping connectSeq makes that connect
      // stale, and since a stale connect returns without touching the
      // button, Connect is re-enabled here.
      function abandonDestination(status) {
        connectSeq += 1;
        connectBtn.disabled = false;
        connected = false;
        targetEligible = false;
        connectedUrl = null;
        invalidatePreflight();
        showTargetStatus(status);
        refreshGate();
      }
      otherRadio.addEventListener('change', () => {
        if (!otherRadio.checked) return;
        urlInput.disabled = false;
        abandonDestination('Enter a site and Connect.');
      });
      connectBtn.addEventListener('click', () => connect(urlInput.value.trim()));
      urlInput.addEventListener('input', () => abandonDestination('Connect to check this site.'));

      // ---- preflight (Check) -------------------------------------------------
      async function runPreflight() {
        if (checking || !connected || !targetEligible) return;
        // Everything the check reads about the destination is pinned here; a
        // connect or an edit mid-check makes the whole answer stale.
        const epoch = inputEpoch;
        const seq = connectSeq;
        const stale = () => epoch !== inputEpoch || seq !== connectSeq;
        const tp = currentTp();
        const targetLib = currentTargetLib();
        const targetWebIdentity = currentTargetIdentity();
        const targetClient = currentTargetClient();
        const targetFields = currentTargetFields();
        const targetWebParts = currentTargetWebParts();
        checking = true;
        showError('');
        checkBtn.disabled = true;
        goBtn.disabled = true;
        checkBtn.textContent = 'Checking…';
        try {
          const folder = folderInput.value.trim().replace(/^\/+|\/+$/g, '');
          const folderIssue = folderProblem(folder);
          const libraryRoot = trimSlash(targetLib.rootPath);
          const finalDir = folder ? `${libraryRoot}/${folder}` : libraryRoot;

          const extraBlockers = [];
          if (folderIssue) extraBlockers.push(folderIssue);

          let fileName;
          const takenFinal = new Set();
          if (!fileNameTouched) {
            const stem = fileStem(snapshot.fileName) || 'Page';
            let free = null;
            for (const candidate of nameCandidates(stem)) {
              // eslint-disable-next-line no-await-in-loop
              const exists = await tp.exists(`${finalDir}/${candidate}`);
              if (exists) takenFinal.add(lower(candidate));
              else { free = candidate; break; }
            }
            fileName = free || `${stem}-copy-${Date.now()}.aspx`;
            nameInput.value = fileName;
          } else {
            fileName = nameInput.value.trim();
            const nameIssue = fileNameProblem(fileName);
            if (!nameIssue && fileName && !folderIssue) {
              if (await tp.exists(`${finalDir}/${fileName}`)) takenFinal.add(lower(fileName));
            }
          }

          const runId = shortRunId();
          const stem = fileStem(fileName) || 'Page';
          const stageName = `${stagingStem(stem, runId)}.aspx`;
          const takenRoot = new Set();
          if (await tp.exists(`${libraryRoot}/${stageName}`)) takenRoot.add(lower(stageName));

          if (folder && !folderIssue) {
            const exists = await tp.folderExists(finalDir);
            if (!exists) extraBlockers.push(`Folder “${folder}” does not exist.`);
          }

          const webSR = trimSlash(targetWebIdentity.webServerRelativeUrl);
          const siteAssetsRoot = `${webSR}/SiteAssets`;
          const assetFolder = `${siteAssetsRoot}/SitePages/${stem}`;
          const assetFolderExists = await tp.folderExists(assetFolder);

          const sameWebNow = Boolean(targetWebIdentity.webId)
            && normalizeGuid(targetWebIdentity.webId) === normalizeGuid(snapshot.web.webId);
          const assetInfo = new Map();
          if (!sameWebNow) {
            const analysis = analyzeParts(snapshot, { analyzers, targetWebParts });
            for (const req of assetRequests(snapshot, analysis)) {
              // eslint-disable-next-line no-await-in-loop
              let info = req.ids?.uniqueId ? await sourcePages.fileInfoById(req.ids.uniqueId) : null;
              // eslint-disable-next-line no-await-in-loop
              if (!info && req.path) info = await sourcePages.fileInfo(req.path);
              assetInfo.set(req.key, info);
            }
          }

          const target = {
            webUrl: targetClient.webUrl(),
            webServerRelativeUrl: webSR,
            siteId: targetWebIdentity.siteId,
            webId: targetWebIdentity.webId,
            library: { id: targetLib.id, rootPath: targetLib.rootPath },
            siteAssetsRoot,
            fields: targetFields,
            webParts: targetWebParts,
            takenFinal,
            takenRoot,
            assetFolderExists,
          };
          const options = {
            fileName,
            title: titleInput.value.trim(),
            folder,
            publish: publishCb.checked,
            promoteAsNews: promoteCb.checked,
            carryMetadata: carryCb.checked,
            rewriteLinks: rewriteCb.checked,
            requiredValues: { ...requiredValues },
            dropped: [...dropped],
            runId,
          };
          const plan = analyzeCopy({ snapshot, target, options, analyzers, assetInfo });
          plan.blockers = [...extraBlockers, ...plan.blockers];

          if (stale()) return;
          lastPlan = plan;
          lastKey = computeKey();
          dropNote.hidden = true;
          renderReport(plan);
        } catch (err) {
          showError(err?.message || String(err));
          lastPlan = null;
          lastKey = null;
        } finally {
          checking = false;
          checkBtn.textContent = 'Check';
          refreshGate();
        }
      }
      checkBtn.addEventListener('click', runPreflight);

      function renderReport(plan) {
        report.hidden = false;
        report.textContent = '';
        requiredEditors.clear();

        if (plan.metadata.carried.length) {
          report.append(el('h3', '', 'Carried metadata'));
          const list = el('ul', '');
          for (const c of plan.metadata.carried) list.append(el('li', '', `${c.title} = ${c.value}`));
          report.append(list);
        }
        if (plan.metadata.skipped.length) {
          report.append(el('h3', '', 'Not carried'));
          const list = el('ul', '');
          for (const s of plan.metadata.skipped) list.append(el('li', '', `${s.title} — ${s.reason}`));
          report.append(list);
        }
        if (plan.metadata.requiredGaps.length) {
          report.append(el('h3', '', 'Required on the destination'));
          const wrap = el('div', 'wb-pc-gaps');
          for (const gap of plan.metadata.requiredGaps) {
            const row = el('div', 'wb-pc-gap-row');
            row.append(el('span', '', gap.title));
            if (gap.supportable) {
              const field = targetFields.find((f) => f.InternalName === gap.internalName);
              if (field) {
                const editor = createFieldEditor(field, '');
                editor.el.addEventListener('input', () => {
                  requiredValues[gap.internalName] = editor.getValue();
                  refreshGate();
                });
                editor.el.addEventListener('change', () => {
                  requiredValues[gap.internalName] = editor.getValue();
                  refreshGate();
                });
                requiredEditors.set(gap.internalName, editor);
                row.append(editor.el);
              }
            } else {
              row.append(el('span', 'wb-pc-gap-unsupported', 'cannot be filled from here'));
            }
            wrap.append(row);
          }
          report.append(wrap);
        }
        if (plan.analysis?.parts?.length) {
          report.append(el('h3', '', 'Page parts'));
          const table = el('table', 'wb-table wb-pc-parts-table');
          const body = el('tbody');
          const droppedNow = new Set((plan.dropped || []).map(lower));
          for (const part of plan.analysis.parts) {
            const isDropped = Boolean(part.instanceId) && droppedNow.has(lower(part.instanceId));
            const tr = el('tr');
            if (isDropped) tr.classList.add('wb-pc-dropped');
            tr.append(el('td', '', part.label || part.kind));
            const chipCell = el('td');
            const counts = partVerdictCounts(part);
            for (const [cls, n] of counts) {
              const chip = el('span', 'wb-info-chip', n > 1 ? `${VERDICT_LABELS[cls] || cls} (${n})` : (VERDICT_LABELS[cls] || cls));
              chipCell.append(chip);
            }
            tr.append(chipCell);

            // Decision 4 (§1): every web part gets a Drop control — never
            // text, never the title area — off by default, and named
            // differently when the destination cannot render it at all.
            const dropCell = el('td');
            if (part.kind === 'webpart' && part.instanceId) {
              const dropLabel = el('label', 'wb-pc-drop-label');
              const dropCb = el('input', 'wb-pc-drop');
              dropCb.type = 'checkbox';
              dropCb.dataset.instance = part.instanceId;
              dropCb.checked = dropped.has(lower(part.instanceId));
              dropLabel.append(dropCb, document.createTextNode(
                part.available === false ? 'Drop (not available on the destination)' : 'Drop',
              ));
              dropCb.addEventListener('change', () => {
                const id = lower(part.instanceId);
                if (dropCb.checked) dropped.add(id); else dropped.delete(id);
                onDropChanged();
              });
              dropCell.append(dropLabel);
            }
            tr.append(dropCell);
            body.append(tr);

            // A dropped dynamic-data provider (§5.2 "dynamic"): list every
            // consumer under it with its own confirm checkbox. Consumers are
            // read off the CURRENT plan, so this only appears once a Check
            // has actually run with the part dropped.
            if (isDropped) {
              const consumers = dependentConsumers(plan.analysis, [part.instanceId]);
              if (consumers.length) {
                const consumerRow = el('tr', 'wb-pc-consumer-row');
                const consumerCell = el('td');
                consumerCell.colSpan = 3;
                const wrap = el('div', 'wb-pc-consumers');
                wrap.append(el('p', 'wb-pc-consumer-note', `Depends on ${part.label || part.kind}:`));
                for (const consumer of consumers) {
                  const consumerLabel = el('label', 'wb-pc-drop-consumer-label');
                  const confirmCb = el('input', 'wb-pc-drop-consumer');
                  confirmCb.type = 'checkbox';
                  confirmCb.dataset.instance = consumer.instanceId;
                  confirmCb.checked = confirmedConsumers.has(lower(consumer.instanceId));
                  consumerLabel.append(confirmCb, document.createTextNode(
                    `Confirm: ${consumer.label || consumer.kind} will lose its data source`,
                  ));
                  confirmCb.addEventListener('change', () => {
                    const id = lower(consumer.instanceId);
                    if (confirmCb.checked) confirmedConsumers.add(id); else confirmedConsumers.delete(id);
                    refreshGate();
                  });
                  wrap.append(consumerLabel);
                }
                consumerCell.append(wrap);
                consumerRow.append(consumerCell);
                body.append(consumerRow);
              }
            }
          }
          table.append(body);
          report.append(table);
        }
        if (plan.assets.length) {
          report.append(el('h3', '', 'Assets'));
          const list = el('ul', '');
          for (const a of plan.assets) {
            list.append(el('li', '', a.retainReason
              ? `${a.name} — retained (${a.retainReason})`
              : `${a.name} — ${bytesText(a.length)} → ${a.destName}`));
          }
          report.append(list);
        }
        if (plan.warnings.length) {
          report.append(el('h3', '', 'Warnings'));
          const list = el('ul', '');
          for (const w of plan.warnings) list.append(el('li', '', w));
          report.append(list);
        }
        if (plan.analysis?.parts?.some((p) => p.unverified?.length)) {
          report.append(el('h3', '', 'Suspicious references'));
          const list = el('ul', '');
          for (const part of plan.analysis.parts) {
            for (const f of part.unverified || []) {
              list.append(el('li', '', `${part.label}: ${f.value}`));
            }
          }
          report.append(list);
        }
        report.append(el('p', 'wb-schema-note', 'People, lookup and managed-metadata columns are not copied.'));
        showError(plan.blockers.join(' '));
      }

      // ---- Copy (phase two) --------------------------------------------------
      async function runTheCopy() {
        if (running) return;
        const fresh = Boolean(lastPlan) && !lastPlan.blockers.length && computeKey() === lastKey;
        if (!fresh) return;
        running = true;
        goBtn.disabled = true;
        checkBtn.disabled = true;
        cancelBtn.disabled = true;
        showError('');
        report.hidden = true;
        runPanel.hidden = false;
        stepsList.textContent = '';
        outcomeEl.hidden = true;
        const stepRows = new Map();

        const frozenPlan = lastPlan;
        const frozenTp = tp;
        const frozenWrite = spWrite;

        function onStep(step) {
          let row = stepRows.get(step.name);
          if (!row) {
            row = el('li', 'wb-pc-step');
            row.dataset.step = step.name;
            stepsList.append(row);
            stepRows.set(step.name, row);
          }
          row.dataset.status = step.status;
          row.textContent = step.detail ? `${step.name} — ${step.detail}` : step.name;
        }

        let result;
        try {
          result = await runCopy({ plan: frozenPlan }, {
            source: { pages: sourcePages },
            target: { pages: frozenTp, write: frozenWrite },
            rewrite: (transferResults) => rewriteContent(snapshot, frozenPlan, transferResults, analyzers),
            verify: (dto, saved) => compareReadBack(frozenPlan, saved, dto),
          }, { onStep });
        } catch (err) {
          result = {
            outcome: 'failed',
            journal: { pageId: null, currentPath: '', createdBy: 'unknown', assets: [], steps: [], retries: [], drift: [] },
            pageUrl: '',
            error: err?.message || String(err),
          };
        }
        running = false;
        runResult = {
          outcome: result.outcome, journal: result.journal, pageUrl: result.pageUrl, sameWeb: Boolean(frozenPlan.sameWeb),
        };
        renderOutcome(result, frozenTp);
      }
      goBtn.addEventListener('click', runTheCopy);

      function renderOutcome(result, frozenTp) {
        const { outcome, journal, pageUrl } = result;
        cancelBtn.hidden = true;
        goBtn.hidden = true;
        checkBtn.hidden = true;
        closeBtn.hidden = false;
        outcomeEl.hidden = false;
        outcomeEl.dataset.outcome = outcome;
        outcomeEl.textContent = outcomeText(outcome, journal);
        if (result.error) showError(result.error);

        if (pageUrl) {
          openLink.href = pageUrl;
          openLink.hidden = false;
        }
        if ((outcome === 'failed' || outcome === 'unknown') && journal?.createdBy === 'this run') {
          discardBtn.hidden = false;
        }
        if (journal?.drift?.length) {
          driftEl.hidden = false;
          driftEl.textContent = '';
          driftEl.append(el('h3', '', 'Differs from the plan'));
          const list = el('ul', '');
          for (const d of journal.drift) {
            list.append(el('li', '', `${d.field}: expected ${d.expected}, got ${d.actual}`));
          }
          driftEl.append(list);
        }

        discardBtn.addEventListener('click', async () => {
          if (discardBusy) return;
          discardBusy = true;
          discardBtn.disabled = true;
          discardBtn.textContent = 'Discarding…';
          try {
            const outcome2 = await discardCopy(journal, { target: { pages: frozenTp } });
            discardBtn.hidden = true;
            const note = el('p', 'wb-pc-discard-note',
              `Recycled ${outcome2.recycled.length} item${outcome2.recycled.length === 1 ? '' : 's'}.`
              + (outcome2.leftovers.length ? ` ${outcome2.leftovers.length} left behind — see the target folder.` : ''));
            runPanel.append(note);
          } catch (err) {
            showError(err?.message || String(err));
            discardBtn.disabled = false;
            discardBtn.textContent = 'Discard copy';
          } finally {
            discardBusy = false;
          }
        }, { once: true });

        closeBtn.addEventListener('click', () => finish(runResult), { once: true });
      }
      let discardBusy = false;

      function outcomeText(outcome, journal) {
        const doneSteps = (journal?.steps || []).filter((s) => s.status === 'done').map((s) => s.name);
        switch (outcome) {
          case 'done':
            return 'The page was copied.';
          case 'done-with-warnings':
            return 'The page was copied, with warnings — see below.';
          case 'failed':
            return 'The copy failed — see below.';
          case 'unknown': {
            const uncertain = (journal?.steps || []).find((s) => s.status === 'unknown');
            const stage = lastPlan?.target?.finalAbsolute || '';
            const staging = lastPlan?.target?.stagingPath || '';
            return `The result of this copy is unknown — the response to the `
              + `${uncertain?.name || 'last'} step was lost. Confirmed so far: `
              + `${doneSteps.length ? doneSteps.join(', ') : 'nothing'}. Check the target `
              + `folder${stage ? ` (${stage})` : ''} and the staging name${staging ? ` (${staging})` : ''} by hand.`;
          }
          default:
            return '';
        }
      }

      // ---- Cancel / close before any write -----------------------------------
      cancelBtn.addEventListener('click', dismiss);

      // ---- kick off ------------------------------------------------------
      connect(client.webUrl());
      refreshGate();
    }
  });
}
