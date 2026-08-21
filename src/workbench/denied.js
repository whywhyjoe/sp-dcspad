// A denied read is a FACT about the site, not a failure of the workbench.
//
// Plenty of what the inspector asks for is legitimately off limits: classic
// webs routinely refuse web/webs to anyone without rights on the child webs,
// a list with broken inheritance refuses its items, a library refuses its
// role assignments. Painting SharePoint's "access denied" in the error
// register told the operator something had gone wrong and invited them to
// retry it — when the honest reading is "this is not yours to see", and the
// only fix is a permission change nobody makes from here.
//
// So a denial renders in the neutral register with a plain sentence, and
// SharePoint's own words stay one hover away. Everything else — a rejected
// query, a broken endpoint, a network failure — is still loud, because those
// are the ones worth acting on.

// requireOk() normalizes 401/403 to code 'permission'; the raw statuses are
// checked too, so errors that never went through it (mock resolvers, hand
// built rejections) classify the same way.
export function isDeniedRead(err) {
  return err?.code === 'permission' || err?.status === 401 || err?.status === 403;
}

// `subject` names what could not be read, in the plural of the thing the
// operator asked for ('subwebs', 'this page'). Sentence case, no jargon: the
// server's own wording is the tooltip's job.
export function deniedNote(subject = '') {
  return subject
    ? `Your account doesn’t have permission to see ${subject} here.`
    : 'Your account doesn’t have permission to see this.';
}

// Paint a status node for a failed read. One helper so a card, a grid and a
// drilldown never disagree about which register a denial belongs in: the
// classes are swapped rather than assigned, so a node keeps whatever layout
// class it was born with (.wb-grid-status, .wb-qb-loading …).
export function showFailure(node, err, subject = '') {
  const denied = isDeniedRead(err);
  node.textContent = denied ? deniedNote(subject) : (err?.message || String(err));
  node.classList.remove('wb-error', 'wb-denied');
  node.classList.add(denied ? 'wb-denied' : 'wb-error');
  // SharePoint's own sentence is still the ground truth — it just isn't the
  // headline. (Only for a denial: a loud error already shows it verbatim.)
  if (denied && err?.message) node.title = err.message;
  else node.removeAttribute('title');
  node.hidden = false;
  return node;
}
