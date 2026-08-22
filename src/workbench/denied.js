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

// 401 is NOT one of them. "Access denied" and "who are you?" are different
// facts with different fixes: a 403 is a standing property of the site that
// only a permission change alters, while a 401 means the sign-in this tab is
// holding has expired and a reload repairs it. Reading them out in the same
// neutral sentence sent the operator to audit their group memberships for a
// problem a refresh solves — so 401 keeps the loud register and gets the one
// instruction that actually works.
//
// requireOk() normalizes 403 to code 'permission' and 401 to 'auth'; the raw
// statuses are checked too, so errors that never went through it (mock
// resolvers, hand built rejections) classify the same way.
export function isDeniedRead(err) {
  return err?.code === 'permission' || err?.status === 403;
}

export function isExpiredSession(err) {
  return err?.code === 'auth' || err?.status === 401;
}

// The whole point of separating 401: name the fix, not the symptom.
export const EXPIRED_SESSION_NOTE =
  'Your SharePoint sign-in has expired — reload the page to sign in again.';

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
  // Loud, like every other failure — but with the fix as the headline. SPO
  // answers an expired session with its own "Access denied" wording, which
  // read verbatim is the one sentence guaranteed to send the operator off to
  // audit permissions that were never the problem.
  const expired = !denied && isExpiredSession(err);
  node.textContent = denied ? deniedNote(subject)
    : expired ? EXPIRED_SESSION_NOTE
      : (err?.message || String(err));
  node.classList.remove('wb-error', 'wb-denied');
  node.classList.add(denied ? 'wb-denied' : 'wb-error');
  // SharePoint's own sentence is still the ground truth — it just isn't the
  // headline. (Only where we replaced it: a loud error already shows it
  // verbatim.)
  if ((denied || expired) && err?.message) node.title = err.message;
  else node.removeAttribute('title');
  node.hidden = false;
  return node;
}
