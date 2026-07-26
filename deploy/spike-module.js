// Loaded by webpart-spike.html as type="module" from an absolute URL.
// Its own import is RELATIVE — the point of the test. If this resolves,
// every relative module specifier resolves too, because module
// specifiers resolve against the importing module's URL, not the page's.
import { SPIKE_IMPORT_OK } from './spike-import.js';

window.__spikeModuleRan = true;
window.__spikeReport('2. module script executes', true, 'type="module" ran inside the web part');
window.__spikeReport('3. relative import from an absolute module URL', SPIKE_IMPORT_OK === true,
  SPIKE_IMPORT_OK === true
    ? 'src/ import tree will resolve — only entry URLs need rewriting'
    : 'unexpected: module ran but its import did not resolve');
