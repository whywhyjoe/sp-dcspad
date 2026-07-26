// Boot splash: the DCSPAD wordmark on a dark curtain — a brief brand
// moment, then a crossfade straight into the editor. Click to skip;
// near-instant under prefers-reduced-motion. Purely cosmetic — the app is
// fully interactive the moment the overlay fades.

import { getState, updateNested } from './state.js';

const LOGO = String.raw`
 ██████╗  ██████╗ ███████╗ ██████╗  █████╗  ██████╗
 ██╔══██╗██╔════╝ ██╔════╝ ██╔══██╗██╔══██╗ ██╔══██╗
 ██║  ██║██║      ███████╗ ██████╔╝███████║ ██║  ██║
 ██║  ██║██║      ╚════██║ ██╔═══╝ ██╔══██║ ██║  ██║
 ██████╔╝╚██████╗ ███████║ ██║     ██║  ██║ ██████╔╝
 ╚═════╝  ╚═════╝ ╚══════╝ ╚═╝     ╚═╝  ╚═╝ ╚═════╝`.slice(1);

export function showSplash() {
  const splash = document.getElementById('splash');
  const logoEl = document.getElementById('splash-logo');
  if (!splash) return;

  const reduced = matchMedia('(prefers-reduced-motion: reduce)').matches;
  logoEl.textContent = LOGO;
  splash.hidden = false;

  let done = false;
  const finish = () => {
    if (done) return;
    done = true;
    splash.classList.add('fading');
    setTimeout(() => splash.remove(), reduced ? 0 : 400);
    if (!getState().settings.seenSplash) updateNested('settings', { seenSplash: true });
  };
  splash.addEventListener('click', finish);
  setTimeout(finish, reduced ? 150 : 700);
}
