// DCSPad web-part bootstrap.
//
// This file is the permanent "Script URL" of the Modern Script Editor web
// part on the hosting page. It is a classic script (the web part injects it
// as <script src>), and it exists so the web part configuration never has to
// change again: it fetches index.html from its own folder, injects the app
// shell into the page, and loads src/main.js as a module from an absolute
// URL — after which every relative import below main.js resolves on its own
// (spike tests 2 + 3, deploy/webpart-spike.html).
//
// index.html stays the single source of truth for the app shell, and keeps
// working standalone — the test suites depend on that.

(function () {
  'use strict';

  // Never take over the page while it is being edited — the pad pinning
  // itself over the canvas would make the web part impossible to configure.
  // A small inert note keeps the web part findable on the canvas.
  if (/[?&]mode=edit/i.test(location.search) || /\/_layouts\//i.test(location.pathname)) {
    var anchor = document.querySelector('[data-dcspad-anchor]');
    if (anchor && !anchor.textContent) {
      anchor.textContent = 'DCSPad (inactive while the page is in edit mode)';
      anchor.style.cssText = 'font:12px/1.5 Consolas,monospace;color:#888;padding:8px';
    }
    return;
  }

  // Modern pages can re-run web-part scripts on re-render; booting twice
  // would duplicate the DOM shell and double-init every module.
  if (window.__DCSPAD_BOOTED__) return;
  window.__DCSPAD_BOOTED__ = true;

  // Resolve the library folder from this script's own URL so a folder move
  // only requires re-pointing the web part, never editing this file.
  var FALLBACK_BASE =
    'https://nervedotnet.sharepoint.com/sites/NewNerve/SiteAssets/Code/dcspad-live/';
  var self = document.currentScript;
  var base = (self && self.src)
    ? self.src.slice(0, self.src.lastIndexOf('/') + 1)
    : FALLBACK_BASE;

  // Mount where the web part put this script if that spot is in the visible
  // canvas; otherwise fall back to the end of <body>.
  var mount = document.createElement('div');
  mount.id = 'dcspad-mount';
  if (self && self.parentNode && document.body.contains(self)) {
    self.parentNode.insertBefore(mount, self.nextSibling);
  } else {
    document.body.appendChild(mount);
  }

  function fail(stage, err) {
    mount.innerHTML = '';
    var box = document.createElement('div');
    box.style.cssText =
      'font:13px/1.5 Consolas,monospace;padding:12px;border:1px solid #b00;color:#b00';
    box.textContent = 'DCSPad failed to boot at "' + stage + '": ' +
      (err && err.message ? err.message : String(err));
    mount.appendChild(box);
  }

  // Hosted-mode flag: app.css pins .app over the viewport when this class
  // is present (see "Web-part hosting" section there). Standalone
  // index.html never sets it.
  document.documentElement.classList.add('dcspad-hosted');

  // SharePoint enters page-edit mode via SPA navigation (no reload), so the
  // boot-time guard above never sees it. Watch the URL and suspend the pad
  // (hide + undo host-page overrides, see .dcspad-suspended in app.css)
  // whenever Mode=Edit appears; restore when it goes away.
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

  fetch(base + 'index.html', { credentials: 'same-origin' })
    .then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status + ' loading index.html');
      return r.text();
    })
    .then(function (html) {
      var doc = new DOMParser().parseFromString(html, 'text/html');

      // Stylesheet: index.html's relative <link> rebased onto the library.
      if (!document.getElementById('dcspad-style')) {
        var link = document.createElement('link');
        link.id = 'dcspad-style';
        link.rel = 'stylesheet';
        link.href = base + 'styles/app.css';
        document.head.appendChild(link);
      }

      // App shell: everything in <body> except scripts (main.js is loaded
      // below as a module; anything else script-shaped doesn't belong here).
      var frag = document.createDocumentFragment();
      Array.prototype.slice.call(doc.body.children).forEach(function (el) {
        if (el.tagName !== 'SCRIPT') frag.appendChild(el);
      });
      mount.appendChild(frag);

      // Entry module from an absolute URL; its own imports resolve against
      // module URLs, not the page, so nothing below it needs rewriting.
      return import(base + 'src/main.js');
    })
    .catch(function (err) { fail('load', err); });
})();
