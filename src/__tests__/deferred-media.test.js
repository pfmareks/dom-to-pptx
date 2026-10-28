// src/__tests__/deferred-media.test.js
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { deferImageData, generateWithDeferredMedia } from '../deferred-media.js';

const PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

// A PptxGenJS-like zip with one XML part and a media entry per item of `media`: a number is that
// parked image's placeholder, a string is written as-is. Parks `images` images in `store`.
async function pptxWithMedia(store, media, images = 1) {
  const placeholders = Array.from({ length: images }, () => {
    const { data } = deferImageData({ data: `data:image/png;base64,${PNG_BASE64}` }, store);
    return data.slice(data.indexOf(',') + 1);
  });
  const pptx = new JSZip();
  pptx.file('ppt/slides/slide1.xml', '<p:sld/>');
  media.forEach((item, i) => {
    const path = `ppt/media/image-1-${i + 1}.png`;
    if (typeof item === 'number') pptx.file(path, placeholders[item], { base64: true });
    else pptx.file(path, item);
  });
  return JSZip.loadAsync(await pptx.generateAsync({ type: 'uint8array' }));
}

describe('generateWithDeferredMedia', () => {
  // The fast path patches the generated zip and only handles stored placeholders in the plain layout
  // JSZip writes by default; anything else must still end up with the real bytes.
  it.each([
    ['a layout it does not patch (data descriptors)', { streamFiles: true }],
    ['placeholders that were deflated rather than stored', { compression: 'DEFLATE' }],
  ])('writes the image bytes for %s', async (_, generateOptions) => {
    const store = [];
    const zip = await pptxWithMedia(store, [0]);

    const blob = await generateWithDeferredMedia(zip, store, () =>
      zip.generateAsync({ type: 'blob', ...generateOptions })
    );

    const out = await JSZip.loadAsync(blob, { checkCRC32: true });
    expect(await out.file('ppt/media/image-1-1.png').async('base64')).toBe(PNG_BASE64);
    expect(await out.file('ppt/slides/slide1.xml').async('string')).toBe('<p:sld/>');
  });

  // A repeated or lost placeholder must fail the export, not give one image another's bytes or
  // leave placeholder text in the deck; both the patch and the JSZip fallback see it.
  it.each([
    ['repeated', [0, 0], /repeated/],
    ['missing', [0, 'unexpected replacement'], /Deferred image 1 has no placeholder/],
  ])('rejects when a placeholder is %s', async (_, media, error) => {
    const store = [];
    const zip = await pptxWithMedia(store, media, 2);

    await expect(generateWithDeferredMedia(zip, store, () => zip.generateAsync({ type: 'blob' }))).rejects.toThrow(
      error
    );
  });
});
