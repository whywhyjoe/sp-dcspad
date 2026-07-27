// Boot splash lifecycle. Hosted mode starts the curtain in boot.js before
// any app asset is fetched; standalone mode starts it here before Monaco.
// A timer enforces only a minimum brand moment — editor readiness, not time,
// decides when the curtain may leave.

import { getState, updateNested } from './state.js';

const LOGO = String.raw`
 ██████╗  ██████╗ ███████╗ ██████╗  █████╗  ██████╗
 ██╔══██╗██╔════╝ ██╔════╝ ██╔══██╗██╔══██╗ ██╔══██╗
 ██║  ██║██║      ███████╗ ██████╔╝███████║ ██║  ██║
 ██║  ██║██║      ╚════██║ ██╔═══╝ ██╔══██║ ██║  ██║
 ██████╔╝╚██████╗ ███████║ ██║     ██║  ██║ ██████╔╝
 ╚═════╝  ╚═════╝ ╚══════╝ ╚═╝     ╚═╝  ╚═╝ ╚═════╝`.slice(1);

function standaloneController(splash) {
  const statusEl = document.getElementById('splash-status');
  const reduced = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const seen = getState().settings.seenSplash;
  const minimumMs = reduced ? 0 : (seen ? 150 : 700);
  const startedAt = performance.now();
  let settled = false;
  let finishTimer = null;
  let removing = false;

  const remove = () => {
    if (removing) return;
    removing = true;
    // Ensure the fully initialized workbench has painted beneath the opaque
    // curtain before revealing it. This prevents an intermediate body frame.
    requestAnimationFrame(() => requestAnimationFrame(() => {
      document.documentElement.classList.add('dcspad-crossfading');
      splash.classList.add('fading');
      setTimeout(() => {
        splash.remove();
        document.documentElement.classList.remove('dcspad-crossfading');
      }, reduced ? 0 : 800);
    }));
    if (!seen) updateNested('settings', { seenSplash: true });
  };

  const controller = {
    status(message) {
      if (!settled && statusEl) statusEl.textContent = message;
    },
    finish() {
      if (settled) return;
      settled = true;
      if (statusEl) statusEl.textContent = 'Editor ready';
      const remaining = Math.max(0, minimumMs - (performance.now() - startedAt));
      finishTimer = setTimeout(remove, remaining);
    },
    fail(message) {
      if (settled) return;
      settled = true;
      clearTimeout(finishTimer);
      splash.classList.add('failed');
      if (statusEl) statusEl.textContent = message;
    },
    skip() {
      if (settled) return;
      settled = true;
      clearTimeout(finishTimer);
      remove();
    },
  };

  splash.addEventListener('click', () => controller.skip());
  return controller;
}

export function showSplash() {
  if (window.__DCSPAD_BOOT_SPLASH__) {
    const hosted = window.__DCSPAD_BOOT_SPLASH__;
    const reduced = matchMedia('(prefers-reduced-motion: reduce)').matches;
    const seen = getState().settings.seenSplash;
    hosted.minimum?.(reduced ? 0 : (seen ? 150 : 700));
    hosted.status('Restoring workspace…');
    if (!seen) updateNested('settings', { seenSplash: true });
    return hosted;
  }

  const splash = document.getElementById('splash');
  const logoEl = document.getElementById('splash-logo');
  if (!splash) {
    return { status() {}, finish() {}, fail() {}, skip() {} };
  }

  // Two flat layers: the block mass stays accent, the box-drawing outline
  // drops to .dim (accent-line). LOGO contains no HTML-significant
  // characters, so the wrap is injection-safe. U+2550–U+256C is the
  // double-line box-drawing range; █ (U+2588) sits outside it.
  logoEl.innerHTML = LOGO.replace(/([═-╬]+)/g, '<span class="dim">$1</span>');
  if (!splash.querySelector('.splash-version')) {
    const ver = document.createElement('div');
    ver.className = 'splash-version';
    ver.innerHTML = 'developer workbench · <b>sharepoint</b>';
    logoEl.after(ver);
  }
  splash.hidden = false;
  splash.getBoundingClientRect();
  if (matchMedia('(prefers-reduced-motion: reduce)').matches) {
    splash.classList.add('visible');
  } else {
    requestAnimationFrame(() => requestAnimationFrame(() => {
      splash.classList.add('visible');
    }));
  }
  return standaloneController(splash);
}
