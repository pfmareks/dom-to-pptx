// src/deferred-media.js
//
// Keeps raster image bytes out of PptxGenJS's base64 path. PptxGenJS decodes every image data URL
// with JSZip's JavaScript base64 decoder while writing its intermediate zip, which costs seconds on
// image-heavy decks. Instead it receives a small unique placeholder per image, and the real bytes,
// decoded natively, are patched into its zip before the export loads it back.
import JSZip from 'jszip';

const BASE64_RASTER = /^data:image\/(png|jpe?g|gif);base64,/;
const MARKER = 'dom-to-pptx-deferred-media:';
const MARKER_BYTES = new TextEncoder().encode(MARKER);

function decodeBase64(base64) {
  if (typeof Uint8Array.fromBase64 === 'function') return Uint8Array.fromBase64(base64);
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/**
 * Returns the image options to hand PptxGenJS: unchanged unless `options.data` is a base64 raster
 * data URL, in which case the data is parked in `store` and replaced by a same-type placeholder.
 *
 * @param {Object} options - PptxGenJS addImage options.
 * @param {string[] | null} store - Per-export list of parked data URLs; null disables deferral.
 * @returns {Object}
 */
export function deferImageData(options, store) {
  const data = options.data;
  if (!store || typeof data !== 'string') return options;
  const prefix = data.match(BASE64_RASTER);
  if (!prefix) return options;
  const placeholder = prefix[0] + btoa(MARKER + store.length);
  store.push(data);
  return { ...options, data: placeholder };
}

let crcTables = null;

// CRC-32 (zip polynomial), eight bytes per step.
function crc32(bytes) {
  if (!crcTables) {
    crcTables = new Int32Array(256 * 8);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTables[n] = c;
    }
    for (let n = 0; n < 256; n++) {
      let c = crcTables[n];
      for (let k = 1; k < 8; k++) {
        c = crcTables[c & 0xff] ^ (c >>> 8);
        crcTables[k * 256 + n] = c;
      }
    }
  }
  const t = crcTables;
  let crc = -1;
  let i = 0;
  const end = bytes.length - (bytes.length % 8);
  for (; i < end; i += 8) {
    const a = (bytes[i] | (bytes[i + 1] << 8) | (bytes[i + 2] << 16) | (bytes[i + 3] << 24)) ^ crc;
    crc =
      t[1792 + (a & 0xff)] ^
      t[1536 + ((a >>> 8) & 0xff)] ^
      t[1280 + ((a >>> 16) & 0xff)] ^
      t[1024 + (a >>> 24)] ^
      t[768 + bytes[i + 4]] ^
      t[512 + bytes[i + 5]] ^
      t[256 + bytes[i + 6]] ^
      t[bytes[i + 7]];
  }
  for (; i < bytes.length; i++) crc = t[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ -1) >>> 0;
}

function placeholderIndex(data) {
  if (data.length <= MARKER_BYTES.length || data.length > MARKER_BYTES.length + 16) return -1;
  for (let i = 0; i < MARKER_BYTES.length; i++) if (data[i] !== MARKER_BYTES[i]) return -1;
  return Number(new TextDecoder().decode(data.subarray(MARKER_BYTES.length)));
}

// Rewrites PptxGenJS's zip with each placeholder entry's data replaced by the parked image bytes.
// Every other header field is copied as-is, so the loaded entries look exactly as PptxGenJS wrote
// them, and as stored entries with a known CRC they pass through the final re-zip untouched.
// Returns null (or throws on a truncated zip) for any layout other than the plain, contiguous zip
// JSZip writes.
function patchPlaceholders(buf, store) {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const eocd = buf.length - 22;
  if (eocd < 0 || view.getUint32(eocd, true) !== 0x06054b50 || view.getUint16(eocd + 20, true) !== 0) return null;
  const count = view.getUint16(eocd + 10, true);
  const cdSize = view.getUint32(eocd + 12, true);
  const cdOffset = view.getUint32(eocd + 16, true);
  if (cdOffset + cdSize !== eocd) return null;

  const localParts = [];
  const centralParts = [];
  const restored = [];
  let expectedLocal = 0;
  let newOffset = 0;
  let cd = cdOffset;
  for (let n = 0; n < count; n++) {
    if (view.getUint32(cd, true) !== 0x02014b50) return null;
    const cdLength = 46 + view.getUint16(cd + 28, true) + view.getUint16(cd + 30, true) + view.getUint16(cd + 32, true);
    const flags = view.getUint16(cd + 8, true);
    const method = view.getUint16(cd + 10, true);
    const compressedSize = view.getUint32(cd + 20, true);
    const local = view.getUint32(cd + 42, true);
    if (flags & 0x08 || local !== expectedLocal || view.getUint32(local, true) !== 0x04034b50) return null;
    const headerLength = 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    const dataStart = local + headerLength;
    let data = buf.subarray(dataStart, dataStart + compressedSize);
    const localHeader = buf.slice(local, dataStart);
    const centralHeader = buf.slice(cd, cd + cdLength);
    const index = method === 0 ? placeholderIndex(data) : -1;
    if (index >= 0 && store[index]) {
      const dataUrl = store[index];
      restored.push(index);
      data = decodeBase64(dataUrl.slice(dataUrl.indexOf(',') + 1));
      const crc = crc32(data);
      for (const [header, at] of [
        [localHeader, 14],
        [centralHeader, 16],
      ]) {
        const h = new DataView(header.buffer);
        h.setUint32(at, crc, true);
        h.setUint32(at + 4, data.length, true);
        h.setUint32(at + 8, data.length, true);
      }
    }
    new DataView(centralHeader.buffer).setUint32(42, newOffset, true);
    localParts.push(localHeader, data);
    centralParts.push(centralHeader);
    newOffset += headerLength + data.length;
    expectedLocal = dataStart + compressedSize;
    cd += cdLength;
  }
  if (expectedLocal !== cdOffset || cd !== eocd) return null;
  // The decoded bytes replace the data URLs; drop them so both are not held through the re-zip.
  for (const index of restored) store[index] = null;
  const end = buf.slice(eocd);
  new DataView(end.buffer).setUint32(16, newOffset, true);
  return new Blob([...localParts, ...centralParts, end]);
}

/**
 * Loads PptxGenJS's output with the parked image bytes in place of their placeholders, keeping
 * each entry's name, position and date.
 *
 * @param {Blob} blob - The zip PptxGenJS wrote.
 * @param {string[] | null} store - The list `deferImageData` filled during this export.
 * @returns {Promise<import('jszip')>}
 */
export async function loadWithDeferredMedia(blob, store) {
  if (!store || store.length === 0) return JSZip.loadAsync(blob);
  let patched = null;
  try {
    patched = patchPlaceholders(new Uint8Array(await blob.arrayBuffer()), store);
  } catch {
    // Truncated or unexpected zip: leave it to JSZip below.
  }
  if (patched) return JSZip.loadAsync(patched);
  const zip = await JSZip.loadAsync(blob);
  await restoreDeferredMedia(zip, store);
  return zip;
}

// Fallback for a zip layout patchPlaceholders does not handle: swaps the entries through JSZip.
async function restoreDeferredMedia(zip, store) {
  if (!store || store.length === 0) return;
  const entries = [];
  zip.forEach((relativePath, file) => {
    if (!file.dir && relativePath.startsWith('ppt/media/') && !relativePath.endsWith('.svg')) {
      entries.push([relativePath, file]);
    }
  });
  for (const [relativePath, file] of entries) {
    const content = await file.async('string');
    if (!content.startsWith(MARKER)) continue;
    const index = Number(content.slice(MARKER.length));
    const data = store[index];
    zip.file(relativePath, decodeBase64(data.slice(data.indexOf(',') + 1)), { binary: true, date: file.date });
  }
}
