// src/deferred-media.js
//
// Keeps raster image bytes out of PptxGenJS's base64 path. PptxGenJS decodes every image data URL
// with JSZip's JavaScript base64 decoder while writing its intermediate zip, which costs seconds on
// image-heavy decks. Instead it receives a small unique placeholder per image, and once that zip is
// loaded back the placeholder entries are swapped for the real bytes, decoded natively.

const BASE64_RASTER = /^data:image\/(png|jpe?g|gif);base64,/;
const MARKER = 'dom-to-pptx-deferred-media:';

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

/**
 * Replaces the placeholder media entries PptxGenJS wrote with the parked image bytes, keeping each
 * entry's name, position and date. Call it on the zip loaded from PptxGenJS's output.
 *
 * @param {import('jszip')} zip - JSZip instance with the loaded PPTX package.
 * @param {string[] | null} store - The list `deferImageData` filled during this export.
 */
export async function restoreDeferredMedia(zip, store) {
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
