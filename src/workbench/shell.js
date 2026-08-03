// Workbench shell: nav rail + view host + internal routing.
//
// Routing is deliberately NOT bound to location.hash — the host is a modern
// SharePoint page whose SPA router owns the URL, and touching it risks
// fighting that router. Routes are plain objects { view, ...params } kept in
// memory and remembered per-tab in sessionStorage (never localStorage:
// state.js in the pad owns that seam and the workbench stays out of it).
//
// Views are factories so they can later mount inside the DCSPad sidebar:
//   create({ client, navigate }) => { el, load(route), destroy?() }
// Views never reach for shell element IDs; everything arrives via deps.

const ROUTE_KEY = 'dcspad.workbench.route';

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

export function createShell({ mount, deps, views }) {
  const instances = new Map();   // id -> { el, load, destroy? }
  let currentRoute = null;

  const rail = el('nav', 'wb-rail');
  rail.setAttribute('aria-label', 'Workbench sections');
  const host = el('main', 'wb-host');

  const buttons = new Map();
  let lastGroup = null;
  for (const view of views) {
    if (view.group !== undefined && view.group !== lastGroup) {
      if (lastGroup !== null) rail.append(el('div', 'wb-rail-sep'));
      rail.append(el('div', 'wb-rail-group', view.group));
    }
    lastGroup = view.group ?? lastGroup;
    const btn = el('button', 'wb-rail-btn');
    btn.type = 'button';
    btn.dataset.view = view.id;
    if (view.glyph) {
      const glyph = el('span', 'wb-rail-glyph');
      glyph.innerHTML = view.glyph;   // static markup from the view registry, not data
      btn.append(glyph);
    }
    btn.append(el('span', 'wb-rail-label', view.label));
    btn.addEventListener('click', () => navigate({ view: view.id }));
    buttons.set(view.id, btn);
    rail.append(btn);
  }
  mount.append(rail, host);

  function instance(id) {
    if (instances.has(id)) return instances.get(id);
    const def = views.find((v) => v.id === id);
    if (!def) return null;
    const inst = def.create({ ...deps, navigate });
    instances.set(id, inst);
    return inst;
  }

  function navigate(route) {
    const def = views.find((v) => v.id === route?.view) ? route : { view: views[0].id };
    currentRoute = def;
    try { sessionStorage.setItem(ROUTE_KEY, JSON.stringify(def)); } catch { /* private mode */ }

    for (const [id, btn] of buttons) {
      btn.classList.toggle('active', id === def.view);
      btn.setAttribute('aria-current', id === def.view ? 'page' : 'false');
    }

    const inst = instance(def.view);
    host.textContent = '';
    host.append(inst.el);
    inst.load?.(def);
  }

  function restore() {
    let saved = null;
    try { saved = JSON.parse(sessionStorage.getItem(ROUTE_KEY) || 'null'); } catch { /* ignore */ }
    navigate(saved || { view: views[0].id });
  }

  // Drop every cached view instance and reland on the current section with
  // no route params. Used when the inspected web changes: everything a view
  // cached belongs to the old web.
  function reset() {
    for (const inst of instances.values()) inst.destroy?.();
    instances.clear();
    navigate({ view: currentRoute?.view || views[0].id });
  }

  return { navigate, restore, reset, getRoute: () => currentRoute };
}
