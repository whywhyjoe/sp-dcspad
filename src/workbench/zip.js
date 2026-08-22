// A minimal ZIP writer — pure, no dependency, no DOM.
//
// Store-only (method 0, no deflate) on purpose. What this packs is markdown:
// small, and the archive exists to group files rather than to shrink them.
// Deflate would mean `CompressionStream`, which is async and would turn every
// caller into a promise chain for a few hundred KB. The format below is the
// 1989 baseline every unzipper reads — no Zip64, no data descriptors, no
// encryption — so sizes and CRCs are all known before a byte is written.

const LOCAL_SIG = 0x04034b50;
const CENTRAL_SIG = 0x02014b50;
const EOCD_SIG = 0x06054b50;
const VERSION = 20;          // 2.0 — the floor for a stored entry
const FLAG_UTF8 = 0x0800;    // bit 11: names are UTF-8, not CP437
const METHOD_STORE = 0;

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
  const files = (entries || []).map((entry) => {
    const name = enc.encode(safeEntryName(entry.name));
    const body = enc.encode(String(entry.text ?? ''));
    return { name, body, crc: crc32(body) };
  });

  const localSize = files.reduce((n, f) => n + 30 + f.name.length + f.body.length, 0);
  const centralSize = files.reduce((n, f) => n + 46 + f.name.length, 0);
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
