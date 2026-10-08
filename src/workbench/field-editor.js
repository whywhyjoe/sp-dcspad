// Per-field-type metadata editors, shared by the Pages metadata tab and the
// file browser. Two layers:
//
//   1. Pure conversions (unit-tested): UI value <-> the FieldValue *string*
//      ValidateUpdateListItem expects. Conventions, per TypeAsString:
//        Text/Note      plain string ('' clears)
//        Choice         the choice string verbatim (fill-ins pass through)
//        MultiChoice    ';#A;#B;#' — ;#-delimited with leading AND trailing ;#
//        Boolean        '1' / '0'
//        Number/Currency  invariant numeric string, '.' decimal separator
//        DateTime       the WEB's locale format in the WEB's time zone, e.g.
//                       '10/6/2026 9:00 AM' (LocaleId 1033). NEVER ISO 8601:
//                       SharePoint Online refuses every ISO form ('…T16:00:00Z',
//                       with milliseconds, without the Z) with "You must
//                       specify a valid date within the range of 1/1/1900 and
//                       12/31/8900" (verified live, 2026-10-06). toFormValue
//                       therefore needs `webDate` (the web's RegionalSettings
//                       format plus its UTC offset at that instant, from
//                       web-dates.js) and throws without it. A date-only field
//                       (DisplayFormat 0) writes the picked calendar date, with
//                       no zone shift.
//        URL            'https://…, description' (comma-space separator)
//      Not editable in v1 (display-only via FieldValuesAsText), formats
//      documented for a later tier:
//        User/UserMulti   '[{"Key":"i:0#.f|membership|user@x"}]'
//        Lookup(Multi)    '1' / '1;#2;#'
//        TaxonomyFieldType 'Label|guid;'
//
//   2. DOM editors: createFieldEditor per field, createFieldEditorForm for a
//      whole item. Plain DOM, .wb-scoped classes, no storage.

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

export const EDITABLE_TYPES = new Set([
  'Text', 'Note', 'Choice', 'MultiChoice', 'Boolean',
  'Number', 'Currency', 'DateTime', 'URL',
]);

// Fields whose values are item *content*, not metadata — corrupting a modern
// page body from a metadata form is the one unrecoverable mistake here.
// WikiField and PublishingPageContent are the classic page bodies — the same
// hazard as CanvasContent1 once the Metadata tab lists Wiki Content.
export const NO_EDIT_INTERNAL = new Set([
  'CanvasContent1', 'LayoutWebpartsContent', 'ContentType', 'Attachments',
  'WikiField', 'PublishingPageContent',
]);

export function isEditable(field) {
  return !field.ReadOnlyField
    && !field.Hidden
    && EDITABLE_TYPES.has(String(field.TypeAsString || ''))
    && !NO_EDIT_INTERNAL.has(String(field.InternalName || ''));
}

const choicesOf = (field) => {
  const v = field?.Choices;
  const arr = Array.isArray(v) ? v : v?.results;
  return Array.isArray(arr) ? arr : [];
};

// ---- pure conversions ------------------------------------------------------

const pad2 = (n) => String(n).padStart(2, '0');

export const isDateOnly = (field) =>
  field?.DisplayFormat === 0 || field?.DisplayFormat === 'DateOnly';

// Wall-clock parts ({ y, m, d, h, min }, m 1-based) → the web's own date
// string, shaped like SharePoint's own display: '10/6/2026 9:00 AM',
// '06.10.2026 16:00' and so on. `webDate`: { order: 'mdy'|'dmy'|'ymd', sep,
// timeSep, time24, am, pm } (web-dates.js webDateFormatOf).
export function formatWebDate(parts, webDate, { dateOnly = false } = {}) {
  const f = webDate || {};
  const byKey = { y: String(parts.y), m: String(parts.m), d: String(parts.d) };
  const date = String(f.order || 'mdy').split('').map((k) => byKey[k]).join(f.sep ?? '/');
  if (dateOnly) return date;
  const timeSep = f.timeSep ?? ':';
  if (f.time24) return `${date} ${pad2(parts.h)}${timeSep}${pad2(parts.min)}`;
  const h12 = parts.h % 12 || 12;
  return `${date} ${h12}${timeSep}${pad2(parts.min)} ${parts.h < 12 ? (f.am || 'AM') : (f.pm || 'PM')}`;
}

// The DateTime editor's value ('YYYY-MM-DDTHH:mm', browser-local) as an
// instant, or null when it is blank or not a date at all.
export function dateTimeInstantOf(uiValue) {
  const s = String(uiValue ?? '').trim();
  if (!s) return null;
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}

// UI value -> the FieldValue string for ValidateUpdateListItem.
// opts.webDate (DateTime only): { order, sep, timeSep, time24, am, pm,
// offsetMinutes }, where offsetMinutes is the web's UTC offset at THIS value's
// instant (createWebDateResolver in web-dates.js builds one per instant).
export function toFormValue(field, uiValue, { webDate = null } = {}) {
  switch (String(field?.TypeAsString || '')) {
    case 'MultiChoice': {
      const arr = Array.isArray(uiValue) ? uiValue.filter(Boolean) : [];
      return arr.length ? `;#${arr.join(';#')};#` : '';
    }
    case 'Boolean':
      return uiValue ? '1' : '0';
    case 'Number':
    case 'Currency': {
      const s = String(uiValue ?? '').trim();
      return s === '' ? '' : String(Number(s.replace(',', '.')));
    }
    case 'DateTime': {
      const s = String(uiValue ?? '').trim();
      if (!s) return '';
      const d = dateTimeInstantOf(s);
      // Not a date: send it verbatim and let SharePoint's own error surface.
      if (!d) return s;
      if (!webDate) {
        throw new Error('A date can only be written in the site’s own regional format, '
          + 'and that format was not available. SharePoint Online refuses ISO 8601 dates.');
      }
      if (isDateOnly(field)) {
        // The calendar date as picked: no zone shift for a date-only field.
        const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
        const parts = m
          ? { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) }
          : { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
        return formatWebDate(parts, webDate, { dateOnly: true });
      }
      // The instant, moved onto the web's wall clock.
      const wall = new Date(d.getTime() + (Number(webDate.offsetMinutes) || 0) * 60000);
      return formatWebDate({
        y: wall.getUTCFullYear(), m: wall.getUTCMonth() + 1, d: wall.getUTCDate(),
        h: wall.getUTCHours(), min: wall.getUTCMinutes(),
      }, webDate);
    }
    case 'URL': {
      const url = String(uiValue?.url ?? '').trim();
      const description = String(uiValue?.description ?? '').trim();
      if (!url) return '';
      return description ? `${url}, ${description}` : url;
    }
    default:
      return String(uiValue ?? '');
  }
}

// REST item value -> the UI value the matching editor consumes.
export function fromItemValue(field, itemValue) {
  switch (String(field?.TypeAsString || '')) {
    case 'MultiChoice': {
      if (Array.isArray(itemValue)) return itemValue;
      if (Array.isArray(itemValue?.results)) return itemValue.results;
      return String(itemValue ?? '').split(';#').filter(Boolean);
    }
    case 'Boolean':
      return itemValue === true || itemValue === 1
        || /^(1|true|yes)$/i.test(String(itemValue ?? ''));
    case 'Number':
    case 'Currency':
      return itemValue === null || itemValue === undefined ? '' : String(itemValue);
    case 'DateTime': {
      const s = String(itemValue ?? '').trim();
      if (!s) return '';
      const d = new Date(s);
      if (Number.isNaN(d.getTime())) return s;
      // datetime-local wants local 'YYYY-MM-DDTHH:mm'.
      return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`
        + `T${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
    }
    case 'URL':
      return {
        url: String(itemValue?.Url ?? itemValue?.url ?? '').trim(),
        description: String(itemValue?.Description ?? itemValue?.description ?? '').trim(),
      };
    default:
      return itemValue === null || itemValue === undefined ? '' : String(itemValue);
  }
}

// ---- DOM editors -----------------------------------------------------------

// One editor row for a field. Returns { el, getValue, isDirty, setError, field }.
export function createFieldEditor(field, initialValue) {
  const type = String(field.TypeAsString || '');
  const initial = fromItemValue(field, initialValue);
  const row = el('div', 'wb-editor-row');
  row.dataset.internal = field.InternalName || '';
  const label = el('label', 'wb-editor-label', field.Title || field.InternalName);
  const typeBadge = el('span', 'wb-editor-type', type);
  label.append(typeBadge);
  const control = el('div', 'wb-editor-control');
  const error = el('div', 'wb-editor-error');
  error.hidden = true;
  row.append(label, control, error);

  let getValue = () => '';

  const textInput = (tag, value) => {
    const input = el(tag === 'textarea' ? 'textarea' : 'input');
    if (tag !== 'textarea') input.type = tag;
    input.value = value ?? '';
    control.append(input);
    return input;
  };

  switch (type) {
    case 'Note': {
      const input = textInput('textarea', initial);
      getValue = () => input.value;
      break;
    }
    case 'Choice': {
      const select = el('select');
      const options = choicesOf(field);
      const blank = el('option', '', '—');
      blank.value = '';
      select.append(blank);
      for (const choice of options) {
        const opt = el('option', '', choice);
        opt.value = choice;
        select.append(opt);
      }
      // Preserve a value that isn't in Choices (fill-in or removed choice).
      if (initial && !options.includes(initial)) {
        const opt = el('option', '', `${initial} (current)`);
        opt.value = initial;
        select.append(opt);
      }
      select.value = initial ?? '';
      control.append(select);
      if (field.FillInChoice) {
        const fillIn = textInput('text', '');
        fillIn.placeholder = 'Fill-in value…';
        getValue = () => fillIn.value.trim() || select.value;
      } else {
        getValue = () => select.value;
      }
      break;
    }
    case 'MultiChoice': {
      const listBox = el('div', 'wb-editor-choices');
      const initialSet = new Set(Array.isArray(initial) ? initial : []);
      const boxes = [];
      for (const choice of choicesOf(field)) {
        const lab = el('label', 'wb-editor-choice');
        const box = el('input');
        box.type = 'checkbox';
        box.value = choice;
        box.checked = initialSet.has(choice);
        lab.append(box, document.createTextNode(choice));
        listBox.append(lab);
        boxes.push(box);
      }
      control.append(listBox);
      getValue = () => boxes.filter((b) => b.checked).map((b) => b.value);
      break;
    }
    case 'Boolean': {
      const box = el('input');
      box.type = 'checkbox';
      box.checked = Boolean(initial);
      control.append(box);
      getValue = () => box.checked;
      break;
    }
    case 'Number':
    case 'Currency': {
      const input = textInput('number', initial);
      input.step = 'any';
      getValue = () => input.value;
      break;
    }
    case 'DateTime': {
      const input = textInput('datetime-local', initial);
      getValue = () => input.value;
      break;
    }
    case 'URL': {
      const url = textInput('text', initial?.url);
      url.placeholder = 'https://…';
      const description = textInput('text', initial?.description);
      description.placeholder = 'Description';
      getValue = () => ({ url: url.value, description: description.value });
      break;
    }
    default: {   // Text and anything else that slipped through as editable
      const input = textInput('text', initial);
      getValue = () => input.value;
    }
  }

  // Dirty tracking compares the editor's own value for a DateTime: its wire
  // form needs the web's regional settings, which only the save resolves.
  const keyOf = (v) => (type === 'DateTime' ? String(v ?? '').trim() : toFormValue(field, v));
  let baseline = keyOf(getValue());
  return {
    el: row,
    field,
    getValue,
    isDirty: () => keyOf(getValue()) !== baseline,
    markClean() { baseline = keyOf(getValue()); },
    setError(message) {
      error.textContent = message || '';
      error.hidden = !message;
      row.classList.toggle('wb-editor-invalid', Boolean(message));
    },
  };
}

// Read-only display row for hidden-from-editing fields. `hint` overrides the
// value's tooltip (a synthesized row is neither a field nor read-only).
function readOnlyRow(field, displayText, hint = '') {
  const row = el('div', 'wb-editor-row wb-editor-readonly');
  row.dataset.internal = field.InternalName || '';
  const label = el('label', 'wb-editor-label', field.Title || field.InternalName);
  label.append(el('span', 'wb-editor-type', String(field.TypeAsString || '')));
  const value = el('div', 'wb-editor-static', displayText || '');
  value.title = hint || (field.ReadOnlyField ? 'Read-only field' : 'Not editable in the workbench');
  row.append(label, value);
  return row;
}

// Whole-item form. fields = the list's field entities; item = raw REST item;
// itemAsText = FieldValuesAsText (display strings for complex types).
// onSave(formValues) must return a promise; a thrown err.fieldErrors map
// ({ InternalName: message }) is routed back onto the matching editors.
//
// layout (optional) fixes which rows appear and in what order, instead of
// every non-hidden field in schema order. Entries:
//   { field, label?, display? }  a list field; label renames the row, display
//                                (value, text) => string replaces the shown
//                                text (read-only rows only)
//   { internal, label, text }    a synthesized read-only row
// A layout field is still edited only when isEditable() allows it.
//
// webDateFor (async (instant: Date) => webDate) supplies the web's regional
// date format and UTC offset for a DateTime value. Pass
// createWebDateResolver(client) from web-dates.js. Without it, saving a
// changed date fails on that field instead of sending ISO, which SPO refuses.
export function createFieldEditorForm({
  fields, item = {}, itemAsText = {}, onSave, layout = null, webDateFor = null,
}) {
  const root = el('div', 'wb-editor-form');
  const rows = el('div', 'wb-editor-rows');
  const editors = [];

  const entries = Array.isArray(layout)
    ? layout
    : (fields || []).filter((f) => !f.Hidden).map((field) => ({ field }));
  for (const entry of entries) {
    if (!entry.field) {
      rows.append(readOnlyRow(
        { InternalName: entry.internal || '', Title: entry.label || '' },
        String(entry.text ?? ''),
        'Page status, read from SharePoint — not an editable field',
      ));
      continue;
    }
    const field = entry.label ? { ...entry.field, Title: entry.label } : entry.field;
    const internal = field.InternalName;
    if (isEditable(field)) {
      const editor = createFieldEditor(field, item[internal]);
      editors.push(editor);
      rows.append(editor.el);
    } else {
      const text = itemAsText?.[internal];
      const raw = item[internal];
      let display = text
        ?? (raw === null || raw === undefined || typeof raw === 'object' ? '' : String(raw));
      if (typeof entry.display === 'function') display = entry.display(raw, text);
      rows.append(readOnlyRow(field, String(display ?? '')));
    }
  }

  const bar = el('div', 'wb-editor-bar');
  const save = el('button', 'btn btn-xs', 'Save metadata');
  save.type = 'button';
  const status = el('span', 'wb-editor-status');
  bar.append(save, status);
  root.append(rows, bar);

  // Async because a DateTime value's wire form needs the web's offset at
  // that instant. A value that cannot be converted carries fieldErrors, so
  // the message lands on its own row.
  async function dirtyFormValues() {
    const formValues = [];
    for (const e of editors.filter((x) => x.isDirty())) {
      const ui = e.getValue();
      try {
        const instant = e.field.TypeAsString === 'DateTime' ? dateTimeInstantOf(ui) : null;
        const webDate = instant && webDateFor ? await webDateFor(instant) : null;
        formValues.push({ FieldName: e.field.InternalName, FieldValue: toFormValue(e.field, ui, { webDate }) });
      } catch (err) {
        if (err?.code === 'auth') throw err;
        throw Object.assign(new Error(err?.message || String(err)), {
          fieldErrors: { [e.field.InternalName]: err?.message || String(err) },
        });
      }
    }
    return formValues;
  }

  save.addEventListener('click', async () => {
    for (const editor of editors) editor.setError('');
    if (!editors.some((e) => e.isDirty())) {
      status.textContent = 'No changes to save.';
      status.className = 'wb-editor-status';
      return;
    }
    save.disabled = true;
    status.textContent = 'Saving…';
    status.className = 'wb-editor-status';
    let formValues = [];
    try {
      formValues = await dirtyFormValues();
      await onSave(formValues);
      for (const editor of editors) editor.markClean();
      status.textContent = `Saved ${formValues.length} field${formValues.length === 1 ? '' : 's'}.`;
      status.className = 'wb-editor-status wb-editor-saved';
    } catch (err) {
      const fieldErrors = err?.fieldErrors || {};
      let mapped = false;
      for (const editor of editors) {
        const message = fieldErrors[editor.field.InternalName];
        if (message) { editor.setError(message); mapped = true; }
      }
      status.textContent = mapped
        ? 'Some fields were rejected — see the messages above.'
        : (err?.message || String(err));
      status.className = 'wb-editor-status wb-editor-failed';
    } finally {
      save.disabled = false;
    }
  });

  return { el: root, getDirtyFormValues: dirtyFormValues, editors };
}
