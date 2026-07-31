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
//        DateTime       ISO 8601 (site-locale strings also accepted by SPO;
//                       server errors surface verbatim so users can hand-fix)
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
export const NO_EDIT_INTERNAL = new Set([
  'CanvasContent1', 'LayoutWebpartsContent', 'ContentType', 'Attachments',
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

// UI value -> the FieldValue string for ValidateUpdateListItem.
export function toFormValue(field, uiValue) {
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
      const d = new Date(s);
      return Number.isNaN(d.getTime()) ? s : d.toISOString();
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
      const pad = (n) => String(n).padStart(2, '0');
      return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
        + `T${pad(d.getHours())}:${pad(d.getMinutes())}`;
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

  let baselineForm = toFormValue(field, getValue());
  return {
    el: row,
    field,
    getValue,
    isDirty: () => toFormValue(field, getValue()) !== baselineForm,
    markClean() { baselineForm = toFormValue(field, getValue()); },
    setError(message) {
      error.textContent = message || '';
      error.hidden = !message;
      row.classList.toggle('wb-editor-invalid', Boolean(message));
    },
  };
}

// Read-only display row for hidden-from-editing fields.
function readOnlyRow(field, displayText) {
  const row = el('div', 'wb-editor-row wb-editor-readonly');
  row.dataset.internal = field.InternalName || '';
  const label = el('label', 'wb-editor-label', field.Title || field.InternalName);
  label.append(el('span', 'wb-editor-type', String(field.TypeAsString || '')));
  const value = el('div', 'wb-editor-static', displayText || '');
  value.title = field.ReadOnlyField ? 'Read-only field' : 'Not editable in the workbench';
  row.append(label, value);
  return row;
}

// Whole-item form. fields = the list's field entities; item = raw REST item;
// itemAsText = FieldValuesAsText (display strings for complex types).
// onSave(formValues) must return a promise; a thrown err.fieldErrors map
// ({ InternalName: message }) is routed back onto the matching editors.
export function createFieldEditorForm({ fields, item = {}, itemAsText = {}, onSave }) {
  const root = el('div', 'wb-editor-form');
  const rows = el('div', 'wb-editor-rows');
  const editors = [];

  const shown = (fields || []).filter((f) => !f.Hidden);
  for (const field of shown) {
    const internal = field.InternalName;
    if (isEditable(field)) {
      const editor = createFieldEditor(field, item[internal]);
      editors.push(editor);
      rows.append(editor.el);
    } else {
      const display = itemAsText?.[internal]
        ?? (item[internal] === null || item[internal] === undefined
          || typeof item[internal] === 'object' ? '' : String(item[internal]));
      rows.append(readOnlyRow(field, String(display ?? '')));
    }
  }

  const bar = el('div', 'wb-editor-bar');
  const save = el('button', 'btn btn-xs', 'Save metadata');
  save.type = 'button';
  const status = el('span', 'wb-editor-status');
  bar.append(save, status);
  root.append(rows, bar);

  function dirtyFormValues() {
    return editors
      .filter((e) => e.isDirty())
      .map((e) => ({
        FieldName: e.field.InternalName,
        FieldValue: toFormValue(e.field, e.getValue()),
      }));
  }

  save.addEventListener('click', async () => {
    for (const editor of editors) editor.setError('');
    const formValues = dirtyFormValues();
    if (!formValues.length) {
      status.textContent = 'No changes to save.';
      status.className = 'wb-editor-status';
      return;
    }
    save.disabled = true;
    status.textContent = 'Saving…';
    status.className = 'wb-editor-status';
    try {
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
