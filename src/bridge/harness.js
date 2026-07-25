// DCSPad iframe harness — injected inline as the FIRST script of every
// assembled preview document, before any library or user code, so it can
// capture everything (including library load failures).
//
// Plain classic script, no imports. Talks to the parent app exclusively
// via postMessage; the __DCSPAD_TOKEN__ placeholder is replaced per run
// so messages from stale iframes are ignored by the parent.
(function () {
  'use strict';
  var TOKEN = '__DCSPAD_TOKEN__';
  var t0 = performance.now();
  var native = {};
  ['log', 'info', 'warn', 'error', 'debug', 'table', 'group', 'groupCollapsed', 'groupEnd', 'clear', 'dir'].forEach(function (m) {
    native[m] = console[m] ? console[m].bind(console) : function () {};
  });

  function post(msg) {
    msg.dcspad = TOKEN;
    try { parent.postMessage(msg, '*'); } catch (e) { /* unserializable payloads never reach here: we pre-serialize */ }
  }

  // ---------------------------------------------------------------
  // Safe structural serializer -> type-tagged JSON tree
  // ---------------------------------------------------------------
  var MAX_DEPTH = 4, MAX_KEYS = 100, MAX_ITEMS = 100, MAX_STR = 5000;

  function preview(value) {
    // One-line compact preview for depth-capped values.
    try {
      if (value === null) return 'null';
      var t = typeof value;
      if (t === 'string') return JSON.stringify(value.length > 40 ? value.slice(0, 40) + '…' : value);
      if (t !== 'object' && t !== 'function') return String(value);
      if (Array.isArray(value)) return 'Array(' + value.length + ')';
      var name = value.constructor ? value.constructor.name : 'Object';
      return name === 'Object' ? '{…}' : name + ' {…}';
    } catch (e) { return '?'; }
  }

  function serialize(value, depth, path) {
    depth = depth || 0;
    path = path || [];
    try {
      if (value === null) return { t: 'null' };
      var t = typeof value;
      if (t === 'undefined') return { t: 'undef' };
      if (t === 'boolean') return { t: 'bool', v: value };
      if (t === 'number') return { t: 'num', v: Number.isFinite(value) ? value : String(value) };
      if (t === 'bigint') return { t: 'num', v: String(value) + 'n' };
      if (t === 'string') {
        return value.length > MAX_STR
          ? { t: 'str', v: value.slice(0, MAX_STR), trunc: value.length }
          : { t: 'str', v: value };
      }
      if (t === 'symbol') return { t: 'sym', v: String(value) };
      if (t === 'function') return { t: 'fn', v: (value.name || 'anonymous') };

      // objects ----
      if (path.indexOf(value) !== -1) return { t: 'circ' };
      if (value instanceof Error || (value && typeof value.message === 'string' && typeof value.stack === 'string' && value.name)) {
        var errNode = { t: 'err', name: value.name || 'Error', msg: value.message, stack: value.stack };
        // PnPjs HttpRequestError carries status/statusText/isHttpRequestError.
        if (value.status !== undefined) errNode.status = value.status;
        if (value.statusText !== undefined) errNode.statusText = value.statusText;
        return errNode;
      }
      if (typeof Node !== 'undefined' && value instanceof Node) {
        var desc = value.nodeName.toLowerCase();
        if (value.id) desc += '#' + value.id;
        if (value.classList && value.classList.length) desc += '.' + Array.prototype.join.call(value.classList, '.');
        return { t: 'node', v: '<' + desc + '>' };
      }
      if (value instanceof Date) return { t: 'date', v: isNaN(value) ? 'Invalid Date' : value.toISOString() };
      if (value instanceof RegExp) return { t: 'regex', v: String(value) };

      if (depth >= MAX_DEPTH) return { t: 'maxdepth', v: preview(value) };

      var nextPath = path.concat([value]);
      if (Array.isArray(value)) {
        var items = [];
        var n = Math.min(value.length, MAX_ITEMS);
        for (var i = 0; i < n; i++) items.push(serialize(value[i], depth + 1, nextPath));
        return { t: 'arr', n: value.length, items: items, trunc: value.length > MAX_ITEMS };
      }
      if (value instanceof Map || value instanceof Set) {
        var entries = [];
        var isMap = value instanceof Map;
        var count = 0;
        value.forEach(function (v, k) {
          if (count++ >= MAX_ITEMS) return;
          entries.push(isMap
            ? [preview(k), serialize(v, depth + 1, nextPath)]
            : [String(count - 1), serialize(v, depth + 1, nextPath)]);
        });
        return { t: 'obj', cls: (isMap ? 'Map' : 'Set') + '(' + value.size + ')', keys: entries, trunc: value.size > MAX_ITEMS };
      }

      var keys = [];
      var names = Object.keys(value);
      var kn = Math.min(names.length, MAX_KEYS);
      for (var j = 0; j < kn; j++) {
        var k = names[j];
        var child;
        try { child = serialize(value[k], depth + 1, nextPath); }
        catch (e) { child = { t: 'str', v: '[getter threw: ' + e.message + ']' }; }
        keys.push([k, child]);
      }
      var cls = 'Object';
      try { cls = (value.constructor && value.constructor.name) || 'Object'; } catch (e) {}
      return { t: 'obj', cls: cls, keys: keys, trunc: names.length > MAX_KEYS };
    } catch (e) {
      return { t: 'str', v: '[unserializable: ' + e.message + ']' };
    }
  }

  function serializeArgs(args) {
    return Array.prototype.map.call(args, function (a) { return serialize(a); });
  }

  // ---------------------------------------------------------------
  // Console capture
  // ---------------------------------------------------------------
  ['log', 'info', 'warn', 'error', 'debug'].forEach(function (level) {
    console[level] = function () {
      native[level].apply(null, arguments);
      post({ kind: 'console', level: level, args: serializeArgs(arguments) });
    };
  });
  console.dir = function (obj) {
    native.dir.apply(null, arguments);
    post({ kind: 'console', level: 'log', args: [serialize(obj)] });
  };
  console.table = function (data, columns) {
    native.table.apply(null, arguments);
    post({ kind: 'table', data: serialize(data), columns: Array.isArray(columns) ? columns : null });
  };
  console.group = function (label) {
    native.group.apply(null, arguments);
    post({ kind: 'group', label: label !== undefined ? String(label) : 'group', collapsed: false });
  };
  console.groupCollapsed = function (label) {
    native.groupCollapsed.apply(null, arguments);
    post({ kind: 'group', label: label !== undefined ? String(label) : 'group', collapsed: true });
  };
  console.groupEnd = function () {
    native.groupEnd();
    post({ kind: 'groupEnd' });
  };
  console.clear = function () {
    native.clear();
    post({ kind: 'clear' });
  };

  // ---------------------------------------------------------------
  // Error capture
  // ---------------------------------------------------------------
  window.addEventListener('error', function (e) {
    // Resource load failures (script/link/img) surface here in capture
    // phase with a target but no error object — vital for library 404s.
    if (e.target && e.target !== window && (e.target.src || e.target.href)) {
      post({
        kind: 'error',
        message: 'Failed to load resource: ' + (e.target.src || e.target.href),
        resource: true,
      });
      return;
    }
    post({
      kind: 'error',
      message: e.message,
      source: e.filename,
      line: e.lineno,
      col: e.colno,
      stack: e.error && e.error.stack ? e.error.stack : null,
    });
  }, true);

  window.addEventListener('unhandledrejection', function (e) {
    var r = e.reason;
    post({
      kind: 'error',
      rejection: true,
      message: (r && r.message) ? r.message : String(r),
      stack: (r && r.stack) ? r.stack : null,
      reason: serialize(r),
    });
  });

  // ---------------------------------------------------------------
  // Fragment links
  // ---------------------------------------------------------------
  // The <base href> that makes _api paths resolve against the SP web
  // also makes "#foo" resolve against that URL rather than this
  // document — so a plain in-page link would navigate the preview away
  // and destroy the run. Re-create the same-document navigation the
  // user's code would get on a real page, hashchange included.
  //
  // Bubble phase, and defaultPrevented is honoured: user code that
  // calls preventDefault() on an <a href="#"> button (the classic
  // no-op-link pattern) must still win, exactly as on a real page.
  document.addEventListener('click', function (e) {
    if (e.defaultPrevented || e.button !== 0) return;
    if (e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
    var a = e.target && e.target.closest ? e.target.closest('a[href]') : null;
    if (!a) return;
    var target = a.getAttribute('target');
    if (target && target !== '_self') return;
    var href = a.getAttribute('href');
    if (!href || href.charAt(0) !== '#') return;

    e.preventDefault();
    var id = href.slice(1);
    if (!id) { window.scrollTo(0, 0); return; }          // bare "#" scrolls to top
    // Assign first so hashchange fires, then scroll explicitly: the
    // assignment is a no-op when the hash already matches, and we do
    // not rely on about:srcdoc honouring fragment navigation at all.
    if (location.hash.slice(1) !== id) {
      try { location.hash = id; } catch (err) { /* fragment nav unsupported — scroll still works */ }
    }
    var el = document.getElementById(id) ||
      (document.getElementsByName(id) || [])[0];
    if (el && el.scrollIntoView) el.scrollIntoView();
  });

  // ---------------------------------------------------------------
  // Network capture: fetch + XMLHttpRequest
  // ---------------------------------------------------------------
  // Ids are namespaced by TOKEN: the parent panel keeps rows across runs,
  // and every fresh iframe restarts netId at 0, so bare f1/x1 would
  // collide with (and silently overwrite) a previous run's entries.
  var netId = 0;

  var nativeFetch = window.fetch ? window.fetch.bind(window) : null;
  if (nativeFetch) {
    window.fetch = function (input, init) {
      var id = TOKEN + ':f' + (++netId);
      var method = (init && init.method) || (input && input.method) || 'GET';
      var url = '';
      try { url = typeof input === 'string' ? input : (input && input.url) || String(input); } catch (e) {}
      var start = performance.now();
      post({ kind: 'net-start', id: id, method: method.toUpperCase(), url: url, api: 'fetch' });
      return nativeFetch(input, init).then(function (res) {
        var ms = Math.round(performance.now() - start);
        var clone;
        try { clone = res.clone(); } catch (e) { clone = null; }
        var finish = function (size, previewText) {
          post({
            kind: 'net-end', id: id, status: res.status, statusText: res.statusText,
            ok: res.ok, ms: ms, size: size,
            contentType: res.headers.get('content-type') || '',
            preview: previewText,
          });
        };
        if (clone) {
          clone.text().then(function (text) {
            finish(text.length, text.slice(0, 20000));
          }, function () { finish(null, null); });
        } else finish(null, null);
        return res;
      }, function (err) {
        post({
          kind: 'net-end', id: id, status: 0, statusText: String(err && err.message || err),
          ok: false, ms: Math.round(performance.now() - start), size: null, preview: null, failed: true,
        });
        throw err;
      });
    };
  }

  var XHR = window.XMLHttpRequest;
  if (XHR) {
    var origOpen = XHR.prototype.open;
    var origSend = XHR.prototype.send;
    XHR.prototype.open = function (method, url) {
      this.__dcspad = { method: String(method).toUpperCase(), url: String(url) };
      return origOpen.apply(this, arguments);
    };
    XHR.prototype.send = function () {
      var xhr = this;
      var meta = xhr.__dcspad || { method: 'GET', url: '' };
      var id = TOKEN + ':x' + (++netId);
      var start = performance.now();
      post({ kind: 'net-start', id: id, method: meta.method, url: meta.url, api: 'xhr' });
      xhr.addEventListener('loadend', function () {
        var previewText = null, size = null;
        try {
          if (xhr.responseType === '' || xhr.responseType === 'text') {
            size = xhr.responseText.length;
            previewText = xhr.responseText.slice(0, 20000);
          }
        } catch (e) {}
        post({
          kind: 'net-end', id: id, status: xhr.status, statusText: xhr.statusText,
          ok: xhr.status >= 200 && xhr.status < 400,
          ms: Math.round(performance.now() - start), size: size,
          contentType: (xhr.getResponseHeader && xhr.getResponseHeader('content-type')) || '',
          preview: previewText, failed: xhr.status === 0,
        });
      });
      return origSend.apply(this, arguments);
    };
  }

  // ---------------------------------------------------------------
  // REPL: evaluate expressions from the parent in this run's context
  // ---------------------------------------------------------------
  window.addEventListener('message', function (e) {
    var d = e.data;
    if (!d || d.dcspad !== TOKEN || d.kind !== 'eval') return;
    var result;
    try {
      result = (0, eval)(d.code);
    } catch (err) {
      post({ kind: 'eval-result', id: d.id, ok: false, value: serialize(err) });
      return;
    }
    // Settle thenables before reporting so `sp.web.get()` is pleasant to poke at.
    if (result && typeof result.then === 'function') {
      result.then(
        function (v) { post({ kind: 'eval-result', id: d.id, ok: true, awaited: true, value: serialize(v) }); },
        function (err) { post({ kind: 'eval-result', id: d.id, ok: false, awaited: true, value: serialize(err) }); }
      );
    } else {
      post({ kind: 'eval-result', id: d.id, ok: true, value: serialize(result) });
    }
  });

  window.addEventListener('load', function () {
    post({ kind: 'loaded', ms: Math.round(performance.now() - t0) });
  });
})();
