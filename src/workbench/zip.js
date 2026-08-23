// A minimal ZIP writer — pure, no dependency, no DOM.
//
// Store-only (method 0, no deflate) on purpose. What this packs is markdown:
// small, and the archive exists to group files rather than to shrink them.
// Deflate would mean `CompressionStream`, which is async and would turn every
// caller into a promise chain for a few hundred KB. The format below is the
// 1989 baseline every unzipper reads — no Zip64, no data descriptors, no
// encryption — so sizes and CRCs are all known before a byte is written.
//
// It fails closed. Anything classic ZIP cannot encode exactly — too many
// entries, a name too long or empty or carrying a control character, a
// payload or archive over 4 GB — throws rather than being narrowed into a
// field too small to hold it. None of it is reachable from the Pages grid
// (bundleEntryName slugs every name, and the caller caps the run), but a
// writer that silently emits an archive contradicting its own headers is not
// worth having.

const LOCAL_SIG = 0x04034b50;
const CENTRAL_SIG = 0x02014b50;
const EOCD_SIG = 0x06054b50;
const VERSION = 20;          // 2.0 — the floor for a stored entry
const FLAG_UTF8 = 0x0800;    // bit 11: names are UTF-8, not CP437
const METHOD_STORE = 0;

// Classic ZIP holds these in fixed 16- and 32-bit fields, and this writer
// deliberately implements no Zip64. A value that does not fit is refused
// rather than narrowed: `setUint16(65536)` writes 0, and an archive whose
// headers disagree with its own bytes is worse than no archive at all.
const MAX_ENTRIES = 0xffff;
const MAX_NAME_BYTES = 0xffff;
const MAX_UINT32 = 0xffffffff;

// A name an extractor would read differently from how it was written. NUL is
// the dangerous one — Info-ZIP and Python both truncate the name there, so
// 'report.md\0a' and 'report.md\0b' are two entries to this writer and one
// file to whoever unpacks it, the second silently overwriting the first.
const CONTROL_CHARS = /[\u0000-\u001f\u007f]/;

let crcTable = null;
function table() {
  if (crcTable) return crcTable;
  crcTable = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[n] = c >>> 0;
  }
  return crcTable;
}

export function crc32(bytes) {
  const t = table();
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) c = t[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// MS-DOS packed time/date. Pre-1980 dates have no representation at all, so
// they clamp rather than wrap into a nonsense year.
function dosStamp(date) {
  const year = Math.max(1980, date.getFullYear());
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1),
    date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  };
}

// A zip entry name is always '/'-separated and always relative: an absolute
// path or a '..' segment makes an archive that writes outside the folder the
// user extracted it into.
export function safeEntryName(name) {
  return String(name || '')
    .replace(/\\/g, '/')
    .split('/')
    .filter((seg) => seg && seg !== '.' && seg !== '..')
    .join('/');
}

// entries: [{ name, text }] — text is UTF-8 encoded here, so callers hand over
// strings and never think about bytes. Returns the whole archive.
export function buildZip(entries, { date = new Date() } = {}) {
  const enc = new TextEncoder();
  const stamp = dosStamp(date);
  const list = entries || [];
  if (list.length > MAX_ENTRIES) {
    throw new Error(`A zip cannot hold more than ${MAX_ENTRIES} entries (got ${list.length}).`);
  }
  const files = list.map((entry) => {
    const safe = safeEntryName(entry.name);
    if (!safe) {
      throw new Error(`Zip entry name ${JSON.stringify(String(entry.name ?? ''))} is empty once `
        + 'it is made relative — an entry with no name cannot be extracted.');
    }
    if (CONTROL_CHARS.test(safe)) {
      throw new Error(`Zip entry name ${JSON.stringify(safe)} carries a control character — `
        + 'extractors truncate the name there, so it would unpack under a different name '
        + 'than it was written under.');
    }
    const name = enc.encode(safe);
    if (name.length > MAX_NAME_BYTES) {
      throw new Error(`Zip entry name is ${name.length} bytes; the limit is ${MAX_NAME_BYTES}.`);
    }
    const body = enc.encode(String(entry.text ?? ''));
    if (body.length > MAX_UINT32) {
      throw new Error(`Zip entry ${JSON.stringify(safe)} is ${body.length} bytes; `
        + 'anything over 4 GB needs Zip64, which this writer does not implement.');
    }
    return { name, body, crc: crc32(body) };
  });

  const localSize = files.reduce((n, f) => n + 30 + f.name.length + f.body.length, 0);
  const centralSize = files.reduce((n, f) => n + 46 + f.name.length, 0);
  // The central directory's offset and size are u32 fields, and every local
  // header offset is measured from the start — so the whole archive has to
  // fit, not just each entry.
  if (localSize + centralSize > MAX_UINT32) {
    throw new Error(`This archive would be ${localSize + centralSize} bytes; anything over `
      + '4 GB needs Zip64, which this writer does not implement.');
  }
  const out = new Uint8Array(localSize + centralSize + 22);
  const view = new DataView(out.buffer);
  let at = 0;
  const u16 = (v) => { view.setUint16(at, v, true); at += 2; };
  const u32 = (v) => { view.setUint32(at, v >>> 0, true); at += 4; };
  const raw = (bytes) => { out.set(bytes, at); at += bytes.length; };

  for (const file of files) {
    file.offset = at;
    u32(LOCAL_SIG);
    u16(VERSION);
    u16(FLAG_UTF8);
    u16(METHOD_STORE);
    u16(stamp.time);
    u16(stamp.date);
    u32(file.crc);
    u32(file.body.length);   // stored: compressed size is the real size
    u32(file.body.length);
    u16(file.name.length);
    u16(0);                  // no extra field
    raw(file.name);
    raw(file.body);
  }

  const centralAt = at;
  for (const file of files) {
    u32(CENTRAL_SIG);
    u16(VERSION);            // version made by
    u16(VERSION);            // version needed
    u16(FLAG_UTF8);
    u16(METHOD_STORE);
    u16(stamp.time);
    u16(stamp.date);
    u32(file.crc);
    u32(file.body.length);
    u32(file.body.length);
    u16(file.name.length);
    u16(0);                  // extra
    u16(0);                  // comment
    u16(0);                  // disk number start
    u16(0);                  // internal attributes
    u32(0);                  // external attributes
    u32(file.offset);
    raw(file.name);
  }

  // Snapshotted before the EOCD is written — `at` is about to move past it.
  const centralEnd = at;
  u32(EOCD_SIG);
  u16(0);                    // this disk
  u16(0);                    // disk with the central directory
  u16(files.length);
  u16(files.length);
  u32(centralEnd - centralAt);
  u32(centralAt);
  u16(0);                    // archive comment
  return out;
}
