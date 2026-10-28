// src/__tests__/deferred-media.test.js
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { deferImageData, loadWithDeferredMedia } from '../deferred-media.js';

const PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

// A PptxGenJS-like stored zip with one deferred image and one XML part.
async function pptxWithPlaceholder(store, generateOptions = {}) {
  const { data } = deferImageData({ data: `data:image/png;base64,${PNG_BASE64}` }, store);
  const pptx = new JSZip();
  pptx.file('ppt/slides/slide1.xml', '<p:sld/>');
  pptx.file('ppt/media/image-1-1.png', data.slice(data.indexOf(',') + 1), { base64: true });
  return new Uint8Array(await pptx.generateAsync({ type: 'uint8array', ...generateOptions }));
}

describe('loadWithDeferredMedia', () => {
  // The fast path patches the zip in place and only handles the plain layout JSZip writes by
  // default; anything else (here: data descriptors from streamFiles) must still get the real bytes.
  it('restores image bytes in a zip layout it does not patch', async () => {
    const store = [];
    const bytes = await pptxWithPlaceholder(store, { streamFiles: true });

    const zip = await loadWithDeferredMedia(new Blob([bytes]), store);

    expect(await zip.file('ppt/media/image-1-1.png').async('base64')).toBe(PNG_BASE64);
    expect(await zip.file('ppt/slides/slide1.xml').async('string')).toBe('<p:sld/>');
  });

  it('restores image bytes when the layout check fails after the entries were read', async () => {
    const store = [];
    const bytes = await pptxWithPlaceholder(store);
    // Four stray bytes between the last entry and the central directory, with the end record's
    // central directory offset moved to match.
    const eocd = bytes.length - 22;
    const cdOffset = new DataView(bytes.buffer).getUint32(eocd + 16, true);
    const padded = new Uint8Array(bytes.length + 4);
    padded.set(bytes.subarray(0, cdOffset));
    padded.set(bytes.subarray(cdOffset), cdOffset + 4);
    new DataView(padded.buffer).setUint32(eocd + 4 + 16, cdOffset + 4, true);

    const zip = await loadWithDeferredMedia(new Blob([padded]), store);

    expect(await zip.file('ppt/media/image-1-1.png').async('base64')).toBe(PNG_BASE64);
  });
});
