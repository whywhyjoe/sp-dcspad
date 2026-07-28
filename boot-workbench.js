// SP Workbench web-part bootstrap — trimmed sibling of boot.js.
//
// Permanent "Script URL" of the Modern Script Editor web part on the
// workbench hosting page. Fetches workbench.html from its own folder,
// injects the shell, and imports the versioned workbench bundle. No Monaco,
// no intelligence, no config — the workbench's asset set is three files.
//
// Deliberately a separate file from boot.js: changing boot.js forces a ?v=
// bump in dcspad.webpart.html and risks the live pad; the workbench must be
// able to iterate without touching it. Keep the guards in sync by hand.

(function () {
  'use strict';

  // Never take over the page while it is being edited.
  if (/[?&]mode=edit/i.test(location.search) || /\/_layouts\//i.test(location.pathname)) {
    var anchor = document.querySelector('[data-wb-anchor]');
    if (anchor && !anchor.textContent) {
      anchor.textContent = 'SP Workbench (inactive while the page is in edit mode)';
      anchor.style.cssText = 'font:12px/1.5 Consolas,monospace;color:#888;padding:8px';
    }
    return;
  }

  // Modern pages can re-run web-part scripts on re-render.
  if (window.__DCSPAD_WB_BOOTED__) return;
  window.__DCSPAD_WB_BOOTED__ = true;

  var FALLBACK_BASE =
    'https://nervedotnet.sharepoint.com/sites/NewNerve/SiteAssets/Code/dcspad-live/';
  var self = document.currentScript;
  var base = (self && self.src)
    ? self.src.slice(0, self.src.lastIndexOf('/') + 1)
    : FALLBACK_BASE;

  var mount = document.createElement('div');
  mount.id = 'wb-mount';
  if (self && self.parentNode && document.body.contains(self)) {
    self.parentNode.insertBefore(mount, self.nextSibling);
  } else {
    document.body.appendChild(mount);
  }

  // Hosted marker (shared with the pad's stylesheet conventions) + a small
  // nonce-stamped loading note. The workbench shell is light enough that the
  // pad's full boot-splash curtain isn't warranted.
  document.documentElement.classList.add('dcspad-hosted');
  var hostNonce = (self && self.nonce) ||
    (document.querySelector('script[nonce]') && document.querySelector('script[nonce]').nonce);

  // Pinned, not in flow: app.css (injected below) makes the host page
  // unscrollable, so an in-flow note under SharePoint's chrome would sit
  // below the fold where neither progress nor a boot failure can be seen.
  var loading = document.createElement('div');
  loading.id = 'wb-boot-note';
  loading.textContent = 'Loading SP Workbench…';
  loading.style.cssText =
    'position:fixed;top:60px;left:12px;z-index:1000;' +
    'font:12px/1.5 Consolas,monospace;color:#a2a9b8;background:#14161b;' +
    'padding:14px;border-radius:6px';
  mount.appendChild(loading);

  function fail(stage, err) {
    loading.textContent = 'SP Workbench failed to boot at "' + stage + '": ' +
      (err && err.message ? err.message : String(err));
    loading.style.color = '#ff6b62';
  }

  // Suspend when SharePoint enters page-edit mode via SPA navigation.
  function syncEditSuspension() {
    var editing = /[?&]mode=edit/i.test(location.search);
    document.documentElement.classList.toggle('dcspad-suspended', editing);
  }
  ['pushState', 'replaceState'].forEach(function (fn) {
    var orig = history[fn];
    history[fn] = function () {
      var r = orig.apply(this, arguments);
      syncEditSuspension();
      return r;
    };
  });
  window.addEventListener('popstate', syncEditSuspension);
  syncEditSuspension();

  // Cache discipline: same protocol as boot.js — one versioned URL per
  // deployable file, stamped from Last-Modified (see the gotcha in CLAUDE.md).
  var VERSIONED = [
    'styles/app.css',
    'styles/workbench.css',
    'dcspad.workbench.js',
  ];

  var versions = {};
  var revalidated = Promise.all(VERSIONED.map(function (f) {
    return fetch(base + f, { credentials: 'same-origin', cache: 'no-cache' })
      .then(function (r) {
        var lm = r.headers.get('Last-Modified');
        versions[f] = lm && !isNaN(new Date(lm)) ? String(new Date(lm).getTime()) : String(Date.now());
      })
      .catch(function () { versions[f] = String(Date.now()); });
  }));

  function versioned(f) {
    return base + f + '?v=' + versions[f];
  }

  fetch(base + 'workbench.html', { credentials: 'same-origin', cache: 'no-store' })
    .then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status + ' loading workbench.html');
      return r.text();
    })
    .then(function (html) {
      return revalidated.then(function () { return html; });
    })
    .then(function (html) {
      var doc = new DOMParser().parseFromString(html, 'text/html');

      ['styles/app.css', 'styles/workbench.css'].forEach(function (href, i) {
        var id = 'wb-style-' + i;
        if (document.getElementById(id)) return;
        var link = document.createElement('link');
        link.id = id;
        link.rel = 'stylesheet';
        link.href = versioned(href);
        if (hostNonce) link.setAttribute('nonce', hostNonce);
        document.head.appendChild(link);
      });

      var frag = document.createDocumentFragment();
      Array.prototype.slice.call(doc.body.children).forEach(function (el) {
        if (el.tagName !== 'SCRIPT') frag.appendChild(el);
      });
      mount.appendChild(frag);
      loading.remove();

      return import(versioned('dcspad.workbench.js'));
    })
    .catch(function (err) { fail('load', err); });
})();
