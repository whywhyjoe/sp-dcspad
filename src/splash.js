// Boot splash: ASCII DCSPAD wordmark with a typewriter reveal and an
// honest little boot readout. Click-to-skip; short shimmer after the
// first visit; instant under prefers-reduced-motion. Purely cosmetic —
// the app is fully interactive the moment the overlay fades.

import { getState, updateNested } from './state.js';

const LOGO = String.raw`
 ██████╗  ██████╗ ███████╗ ██████╗  █████╗  ██████╗
 ██╔══██╗██╔════╝ ██╔════╝ ██╔══██╗██╔══██╗ ██╔══██╗
 ██║  ██║██║      ███████╗ ██████╔╝███████║ ██║  ██║
 ██║  ██║██║      ╚════██║ ██╔═══╝ ██╔══██║ ██║  ██║
 ██████╔╝╚██████╗ ███████║ ██║     ██║  ██║ ██████╔╝
 ╚═════╝  ╚═════╝ ╚══════╝ ╚═╝     ╚═╝  ╚═╝ ╚═════╝`.slice(1);

export function showSplash({ spContext }) {
  const splash = document.getElementById('splash');
  const logoEl = document.getElementById('splash-logo');
  const bootEl = document.getElementById('splash-boot');
  if (!splash) return;

  const reduced = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const seen = getState().settings.seenSplash;
  splash.hidden = false;

  const bootLines = [
    ['▸ editors ............... ', 'ready', 'ok'],
    ['▸ harness ............... ', 'armed', 'ok'],
    ['▸ inspector ............. ', 'SP-aware', 'ok'],
    ['▸ context ............... ', spContext.live ? `LIVE · ${spContext.label}` : 'MOCK (not in SharePoint)', spContext.live ? 'ok' : 'warn'],
  ];

  let done = false;
  const timers = [];
  const finish = () => {
    if (done) return;
    done = true;
    timers.forEach(clearTimeout);
    splash.classList.add('fading');
    setTimeout(() => splash.remove(), reduced ? 0 : 400);
    if (!seen) updateNested('settings', { seenSplash: true });
  };
  splash.addEventListener('click', finish);

  const logoLines = LOGO.split('\n');

  if (reduced || seen) {
    // Quick shimmer: everything at once, brief hold.
    logoEl.textContent = LOGO;
    for (const [label, value, cls] of bootLines) {
      bootEl.append(label);
      const span = document.createElement('span');
      span.className = cls;
      span.textContent = value;
      bootEl.append(span, '\n');
    }
    timers.push(setTimeout(finish, reduced ? 400 : 700));
    return;
  }

  // First visit: line-by-line logo reveal, then typed boot readout.
  let t = 0;
  logoLines.forEach((line, i) => {
    timers.push(setTimeout(() => { logoEl.textContent += line + '\n'; }, t));
    t += 70;
  });
  t += 200;
  bootLines.forEach(([label, value, cls]) => {
    timers.push(setTimeout(() => {
      bootEl.append(label);
      const span = document.createElement('span');
      span.className = cls;
      span.textContent = value;
      bootEl.append(span, '\n');
    }, t));
    t += 160;
  });
  timers.push(setTimeout(finish, t + 700));
}
