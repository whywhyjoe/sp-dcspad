// Export + clipboard helpers for workbench grids.
//
// Everything operates on (rows, columns) exactly as the grid renders them —
// what you see (after filter/sort) is what you export. File downloads go
// through the pad's io.js so byte-moving stays in one place.

import { downloadText } from '../io.js?v=2';

export function copyText(text, flashEl) {
  const done = () => {
    if (!flashEl) return;
    flashEl.classList.add('copied');
    setTimeout(() => flashEl.classList.remove('copied'), 900);
  };
  if (navigator.clipboard?.writeText) {
    navigator.clipboard.writeText(text).then(done, done);
  } else {
    const ta = document.createElement('textarea');
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); } catch { /* best effort */ }
    ta.remove();
    done();
  }
}

const cellValue = (row, col) =>
  typeof col.value === 'function' ? col.value(row) : row[col.key];

export function cellText(row, col) {
  const v = cellValue(row, col);
  if (typeof col.format === 'function') return String(col.format(v, row) ?? '');
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'Yes' : '';
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
}

// Columns marked `action` exist only to host per-row controls (export
// buttons, and the like). They carry no data, so every export drops them \u2014
// otherwise they contribute an empty CSV column and a JSON key whose value
// is whatever the render happened to key off.
const dataColumns = (columns) => (columns || []).filter((c) => !c.action);

// RFC 4180: quote fields containing commas, quotes, or line breaks; double
// embedded quotes. BOM prefix so Excel opens UTF-8 correctly.
export function toCsv(rows, columns) {
  const cols = dataColumns(columns);
  const quote = (s) => (/[",\r\n]/.test(s) ? `"${s.replaceAll('"', '""')}"` : s);
  const lines = [cols.map((c) => quote(String(c.label ?? c.key))).join(',')];
  for (const row of rows) {
    lines.push(cols.map((c) => quote(cellText(row, c))).join(','));
  }
  return `\uFEFF${lines.join('\r\n')}`;
}

// JSON keeps raw row values (not display formatting) under column keys.
export function toJson(rows, columns) {
  const cols = dataColumns(columns);
  const out = rows.map((row) => {
    const record = {};
    for (const c of cols) {
      const v = cellValue(row, c);
      record[c.key] = v === undefined ? null : v;
    }
    return record;
  });
  return JSON.stringify(out, null, 2);
}

export function toMarkdown(rows, columns) {
  const cols = dataColumns(columns);
  const esc = (s) => s.replaceAll('|', '\\|').replaceAll('\r', '').replaceAll('\n', ' ');
  const lines = [
    `| ${cols.map((c) => esc(String(c.label ?? c.key))).join(' | ')} |`,
    `| ${cols.map(() => '---').join(' | ')} |`,
  ];
  for (const row of rows) {
    lines.push(`| ${cols.map((c) => esc(cellText(row, c))).join(' | ')} |`);
  }
  return lines.join('\n');
}

export function downloadCsv(name, rows, columns) {
  downloadText(`${name}.csv`, toCsv(rows, columns), 'text/csv;charset=utf-8');
}

export function downloadJson(name, rows, columns) {
  downloadText(`${name}.json`, toJson(rows, columns), 'application/json');
}

// Prebuilt markdown documents (e.g. the list-content export) — the text is
// already assembled, this just names the bytes.
export function downloadMarkdown(name, text) {
  downloadText(`${name}.md`, text, 'text/markdown;charset=utf-8');
}
