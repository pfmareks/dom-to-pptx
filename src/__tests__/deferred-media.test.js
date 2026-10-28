// src/__tests__/deferred-media.test.js
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { deferImageData, generateWithDeferredMedia } from '../deferred-media.js';

const PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

// A PptxGenJS-like zip holding one deferred image's placeholder and one XML part.
async function pptxWithPlaceholder(store) {
  const { data } = deferImageData({ data: `data:image/png;base64,${PNG_BASE64}` }, store);
  const pptx = new JSZip();
  pptx.file('ppt/slides/slide1.xml', '<p:sld/>');
  pptx.file('ppt/media/image-1-1.png', data.slice(data.indexOf(',') + 1), { base64: true });
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
    const zip = await pptxWithPlaceholder(store);

    const blob = await generateWithDeferredMedia(zip, store, () =>
      zip.generateAsync({ type: 'blob', ...generateOptions })
    );

    const out = await JSZip.loadAsync(blob, { checkCRC32: true });
    expect(await out.file('ppt/media/image-1-1.png').async('base64')).toBe(PNG_BASE64);
    expect(await out.file('ppt/slides/slide1.xml').async('string')).toBe('<p:sld/>');
  });
});
