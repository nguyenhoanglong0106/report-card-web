// Minimal browser-side ZIP reader/writer used only to merge several .docx
// files (themselves ZIP containers) into one, without pulling in a full zip
// library. Reading supports STORE and DEFLATE entries (DEFLATE decoded via the
// browser's native DecompressionStream); writing only produces STORE entries,
// matching what the Worker itself generates, so merging never needs a JS
// deflate implementation.

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

async function inflateRaw(bytes) {
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

// Returns Map<filename, Uint8Array> of every entry in the zip, decompressed.
async function parseZip(arrayBuffer) {
  const view = new DataView(arrayBuffer);
  const bytes = new Uint8Array(arrayBuffer);

  let eocd = -1;
  for (let i = bytes.length - 22; i >= 0 && i >= bytes.length - 22 - 65536; i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd === -1) throw new Error('File .docx không hợp lệ (thiếu EOCD).');

  const entryCount = view.getUint16(eocd + 10, true);
  let cdOffset = view.getUint32(eocd + 16, true);

  const entries = new Map();
  for (let i = 0; i < entryCount; i++) {
    if (view.getUint32(cdOffset, true) !== 0x02014b50) throw new Error('File .docx không hợp lệ (central directory).');
    const method = view.getUint16(cdOffset + 10, true);
    const compressedSize = view.getUint32(cdOffset + 20, true);
    const nameLen = view.getUint16(cdOffset + 28, true);
    const extraLen = view.getUint16(cdOffset + 30, true);
    const commentLen = view.getUint16(cdOffset + 32, true);
    const localHeaderOffset = view.getUint32(cdOffset + 42, true);
    const name = new TextDecoder().decode(bytes.subarray(cdOffset + 46, cdOffset + 46 + nameLen));

    const lhNameLen = view.getUint16(localHeaderOffset + 26, true);
    const lhExtraLen = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + lhNameLen + lhExtraLen;
    const raw = bytes.subarray(dataStart, dataStart + compressedSize);

    entries.set(name, method === 0 ? raw.slice() : await inflateRaw(raw));
    cdOffset += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

// Builds a STORE-only zip from Map<filename, Uint8Array>.
function buildZip(entries) {
  const encoder = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const [name, data] of entries) {
    const nameBytes = encoder.encode(name);
    const crc = crc32(data);
    const local = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true); // version needed
    lv.setUint16(6, 0, true); // flags
    lv.setUint16(8, 0, true); // method: STORE
    lv.setUint16(10, 0, true); // mod time
    lv.setUint16(12, 0, true); // mod date
    lv.setUint32(14, crc, true);
    lv.setUint32(18, data.length, true); // compressed size
    lv.setUint32(22, data.length, true); // uncompressed size
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true); // extra length
    local.set(nameBytes, 30);

    localParts.push(local, data);

    const central = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0, true);
    cv.setUint16(10, 0, true);
    cv.setUint16(12, 0, true);
    cv.setUint16(14, 0, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);
    cv.setUint16(32, 0, true);
    cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true);
    cv.setUint32(38, 0, true);
    cv.setUint32(42, offset, true);
    central.set(nameBytes, 46);
    centralParts.push(central);

    offset += local.length + data.length;
  }

  const centralStart = offset;
  const centralSize = centralParts.reduce((n, p) => n + p.length, 0);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, entries.size, true);
  ev.setUint16(10, entries.size, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, centralStart, true);

  return new Blob([...localParts, ...centralParts, eocd]);
}

const decoder = new TextDecoder();

// Merges several single-student .docx ArrayBuffers (all rendered from the same
// template) into one .docx with a page break between each student. Mirrors the
// server's renderCombinedDocx, run in the browser instead so a whole class
// never has to be rendered in a single Worker request.
export async function mergeDocx(arrayBuffers) {
  if (!arrayBuffers.length) throw new Error('Không có phiếu để ghép.');
  const zips = [];
  for (const buf of arrayBuffers) zips.push(await parseZip(buf));

  const docXmls = zips.map((entries) => decoder.decode(entries.get('word/document.xml')));
  const bodies = docXmls.map((xml) => xml.match(/<w:body>([\s\S]*)<\/w:body>/)[1]);
  const sectPrMatch = bodies[0].match(/(<w:sectPr[\s\S]*<\/w:sectPr>)/);
  const sectPr = sectPrMatch ? sectPrMatch[1] : '';
  const bodiesNoSect = bodies.map((b) => b.replace(/<w:sectPr[\s\S]*<\/w:sectPr>/, ''));
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const mergedBody = bodiesNoSect.join(pageBreak) + sectPr;
  const mergedXml = docXmls[0].replace(/<w:body>[\s\S]*<\/w:body>/, `<w:body>${mergedBody}</w:body>`);

  const baseEntries = zips[0];
  baseEntries.set('word/document.xml', new TextEncoder().encode(mergedXml));
  return buildZip(baseEntries);
}
