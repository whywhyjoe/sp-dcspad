// Hosted-only SharePoint toolbar visibility. The suite bar remains owned by
// SharePoint; DCSPad only translates it out of view and reclaims the top inset.

export function initSpChromeToggle(initialContext) {
  const root = document.documentElement;
  const chip = document.getElementById('sp-chip');
  const suiteNav = document.getElementById('SuiteNavWrapper');
  const available = root.classList.contains('dcspad-hosted') && !!suiteNav;
  let context = initialContext;

  function update() {
    const hidden = root.classList.contains('dcspad-chrome-hidden');
    chip.disabled = !available;
    chip.setAttribute('aria-expanded', String(available && !hidden));
    chip.title = available
      ? `${hidden ? 'Show' : 'Hide'} SharePoint toolbar`
      : 'SharePoint toolbar unavailable outside hosted mode';

    const contextLabel = context?.live
      ? `SP: Live; connected to ${context.label}${context.user ? ` as ${context.user}` : ''}`
      : 'SP: Mock; not connected to a SharePoint web';
    chip.setAttribute('aria-label', available ? `${contextLabel}; ${chip.title}` : contextLabel);
  }

  chip.addEventListener('click', () => {
    if (!available) return;
    root.classList.toggle('dcspad-chrome-hidden');
    update();
  });
  update();

  return {
    setContext(nextContext) {
      context = nextContext;
      update();
    },
  };
}
