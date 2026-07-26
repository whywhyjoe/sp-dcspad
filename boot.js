// DCSPad web-part bootstrap.
//
// This file is the permanent "Script URL" of the Modern Script Editor web
// part on the hosting page. It is a classic script (the web part injects it
// as <script src>), and it exists so the web part configuration never has to
// change again: it fetches index.html from its own folder, injects the app
// shell into the page, and loads the versioned hosted app bundle from an
// absolute URL (spike tests 2 + 3, deploy/webpart-spike.html).
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

  // Mark hosted mode before the first curtain frame so the critical dark
  // underlay below can cover SharePoint's white canvas immediately.
  document.documentElement.classList.add('dcspad-hosted');

  // Paint a dependency-free curtain immediately. This is deliberately
  // before index.html, cache probes, app.css, the app bundle, and Monaco:
  // on a cold SharePoint load it explains every otherwise-blank wait.
  var BOOT_LOGO = [
    ' ██████╗  ██████╗ ███████╗ ██████╗  █████╗  ██████╗',
    ' ██╔══██╗██╔════╝ ██╔════╝ ██╔══██╗██╔══██╗ ██╔══██╗',
    ' ██║  ██║██║      ███████╗ ██████╔╝███████║ ██║  ██║',
    ' ██║  ██║██║      ╚════██║ ██╔═══╝ ██╔══██║ ██║  ██║',
    ' ██████╔╝╚██████╗ ███████║ ██║     ██║  ██║ ██████╔╝',
    ' ╚═════╝  ╚═════╝ ╚══════╝ ╚═╝     ╚═╝  ╚═╝ ╚═════╝',
  ].join('\n');
  var bootStyle = document.createElement('style');
  bootStyle.id = 'dcspad-boot-style';
  var hostNonce = (self && self.nonce) ||
    (document.querySelector('script[nonce]') && document.querySelector('script[nonce]').nonce);
  if (hostNonce) bootStyle.setAttribute('nonce', hostNonce);
  bootStyle.textContent =
    'html.dcspad-hosted,html.dcspad-hosted body{background:#101216}' +
    'html.dcspad-hosted #dcspad-mount:before{content:"";position:fixed;inset:0;' +
    'z-index:998;background:#101216;pointer-events:none}' +
    '#splash.dcspad-boot-splash{position:fixed;inset:0;z-index:1000;background:#16181d;' +
    'display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;' +
    'opacity:0;transition:opacity .75s ease;color:#4ec9b0;cursor:pointer}' +
    '#splash.dcspad-boot-splash.visible{opacity:1}' +
    '#splash.dcspad-boot-splash.fading{opacity:0;pointer-events:none}' +
    '#splash.dcspad-boot-splash pre{font:clamp(8px,1.6vw,14px)/1.15 Consolas,monospace;' +
    'white-space:pre;margin:0;text-shadow:0 0 14px #4ec9b044}' +
    '#splash.dcspad-boot-splash .splash-status{min-height:1.4em;color:#b8bdc9;' +
    'font:12px/1.4 Consolas,monospace;letter-spacing:.04em}' +
    '#splash.dcspad-boot-splash.failed .splash-status{color:#f47067}' +
    '@media(prefers-reduced-motion:reduce){#splash.dcspad-boot-splash{transition:none}}';
  document.head.appendChild(bootStyle);

  var bootSplashEl = document.createElement('div');
  bootSplashEl.id = 'splash';
  bootSplashEl.className = 'splash dcspad-boot-splash';
  var bootLogo = document.createElement('pre');
  bootLogo.id = 'splash-logo';
  bootLogo.className = 'splash-logo';
  bootLogo.textContent = BOOT_LOGO;
  var bootStatus = document.createElement('div');
  bootStatus.id = 'splash-status';
  bootStatus.className = 'splash-status';
  bootStatus.setAttribute('role', 'status');
  bootStatus.setAttribute('aria-live', 'polite');
  bootStatus.textContent = 'Loading workbench shell…';
  bootSplashEl.appendChild(bootLogo);
  bootSplashEl.appendChild(bootStatus);
  mount.appendChild(bootSplashEl);
  // Commit the opacity:0 state before adding .visible. A single rAF can be
  // batched into SharePoint's current render and skip the transition.
  bootSplashEl.getBoundingClientRect();
  requestAnimationFrame(function () {
    requestAnimationFrame(function () {
      bootSplashEl.classList.add('visible');
    });
  });

  var splashStartedAt = Date.now();
  var splashMinimumMs = 700;
  var splashSettled = false;
  var splashRemoved = false;
  function removeBootSplash() {
    if (splashRemoved) return;
    splashRemoved = true;
    // Monaco readiness can resolve in the same frame as its final paint.
    // Give the completed app two frames underneath the opaque curtain, then
    // fade only the curtain. The app and its dark surround stay fully opaque,
    // so SharePoint's white wrappers can never show through the midpoint.
    requestAnimationFrame(function () {
      requestAnimationFrame(function () {
        document.documentElement.classList.add('dcspad-crossfading');
        bootSplashEl.classList.add('fading');
        setTimeout(function () {
          if (bootSplashEl.parentNode) bootSplashEl.parentNode.removeChild(bootSplashEl);
          if (bootStyle.parentNode) bootStyle.parentNode.removeChild(bootStyle);
          document.documentElement.classList.remove('dcspad-crossfading');
        }, 800);
      });
    });
  }
  var bootSplash = {
    minimum: function (ms) { splashMinimumMs = Math.max(0, Number(ms) || 0); },
    status: function (message) {
      if (!splashSettled) bootStatus.textContent = message;
    },
    finish: function () {
      if (splashSettled) return;
      splashSettled = true;
      bootStatus.textContent = 'Editor ready';
      setTimeout(removeBootSplash, Math.max(0, splashMinimumMs - (Date.now() - splashStartedAt)));
    },
    fail: function (message) {
      if (splashRemoved) return false;
      splashSettled = true;
      bootSplashEl.classList.remove('fading');
      bootSplashEl.classList.add('failed', 'visible');
      bootStatus.textContent = message;
      return true;
    },
    skip: function () {
      if (splashSettled) return;
      splashSettled = true;
      removeBootSplash();
    },
  };
  bootSplashEl.addEventListener('click', bootSplash.skip);
  window.__DCSPAD_BOOT_SPLASH__ = bootSplash;

  function fail(stage, err) {
    var message = 'DCSPad failed to boot at "' + stage + '": ' +
      (err && err.message ? err.message : String(err));
    if (bootSplash.fail(message)) return;
    var box = document.createElement('div');
    box.style.cssText =
      'font:13px/1.5 Consolas,monospace;padding:12px;border:1px solid #b00;color:#b00';
    box.textContent = message;
    mount.appendChild(box);
  }

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

  // ---- cache discipline -------------------------------------------------
  // SharePoint serves library files with `cache-control: public,
  // max-age=86400`, Chrome caches module-script requests separately from
  // fetch(), and this host page freezes import-map registration — all
  // proven empirically on this tenant. So hosted mode loads the app as ONE
  // bundled ESM file (dcspad.app.js, built by tools/build-app.mjs) behind a
  // versioned URL: a conditional GET reads its Last-Modified, that stamps
  // the import URL, and a deploy busts exactly one entry. Same for the
  // stylesheet. Monaco is a separately generated vendor set whose manifest
  // versions its runtime, CSS/font, workers and PnPjs declarations together.
  // index.html is always fetched no-store. The harness is fetched as text
  // with no-cache by the runner, so it stays fresh on its own.
  var VERSIONED = ['styles/app.css', 'dcspad.app.js', 'vendor/monaco/version.json'];

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

  // The runner fetches src/bridge/harness.js relative to this base (the
  // bundle's import.meta.url would point at the bundle itself).
  window.__DCSPAD_SRC_BASE__ = base + 'src/';
  // Monaco is one separately-versioned ESM bundle plus same-origin worker,
  // CSS, font, and type assets. Every file in that vendor set shares the
  // version stamp below, so workers never resolve relative to the SP page.
  window.__DCSPAD_ASSET_BASE__ = base;

  fetch(base + 'index.html', { credentials: 'same-origin', cache: 'no-store' })
    .then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status + ' loading index.html');
      return r.text();
    })
    .then(function (html) {
      return revalidated.then(function () {
        window.__DCSPAD_MONACO_VERSION__ = versions['vendor/monaco/version.json'];
        bootSplash.status('Starting application…');
        return html;
      });
    })
    .then(function (html) {
      var doc = new DOMParser().parseFromString(html, 'text/html');

      // Stylesheet: index.html's relative <link> rebased onto the library.
      if (!document.getElementById('dcspad-style')) {
        var link = document.createElement('link');
        link.id = 'dcspad-style';
        link.rel = 'stylesheet';
        link.href = versioned('styles/app.css');
        document.head.appendChild(link);
      }

      // App shell: everything in <body> except scripts (the app itself is
      // imported below; anything else script-shaped doesn't belong here).
      var frag = document.createDocumentFragment();
      Array.prototype.slice.call(doc.body.children).forEach(function (el) {
        if (el.tagName !== 'SCRIPT' && el.id !== 'splash') frag.appendChild(el);
      });
      mount.appendChild(frag);

      bootSplash.status('Starting Monaco editor…');
      return import(versioned('dcspad.app.js'));
    })
    .catch(function (err) { fail('load', err); });
})();
