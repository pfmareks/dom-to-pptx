// src/__tests__/output-zip.test.js
import { describe, it, expect, beforeAll } from 'vitest';
import JSZip from 'jszip';
import { exportToPptx } from '../index.js';

// 1x1 red PNG.
const PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

// Compression method per entry name, read from the local file headers (0 = STORE, 8 = DEFLATE).
function entryMethods(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const methods = {};
  let offset = 0;
  while (offset + 30 <= bytes.length && view.getUint32(offset, true) === 0x04034b50) {
    const method = view.getUint16(offset + 8, true);
    const compressedSize = view.getUint32(offset + 18, true);
    const nameLength = view.getUint16(offset + 26, true);
    const extraLength = view.getUint16(offset + 28, true);
    const name = new TextDecoder().decode(bytes.subarray(offset + 30, offset + 30 + nameLength));
    methods[name] = method;
    offset += 30 + nameLength + extraLength + compressedSize;
  }
  return methods;
}

describe('exported package compression', () => {
  beforeAll(() => {
    // JSDOM has no canvas backend: colour parsing needs a 2D context, the <canvas> export path toDataURL.
    HTMLCanvasElement.prototype.getContext = () => ({
      fillStyle: '',
      clearRect: () => {},
      fillRect: () => {},
      getImageData: () => ({ data: [0, 0, 0, 0] }),
    });
    HTMLCanvasElement.prototype.toDataURL = () => `data:image/png;base64,${PNG_BASE64}`;
  });

  it('stores PNG media as-is and deflates the XML parts', async () => {
    const slide = document.createElement('div');
    const canvas = document.createElement('canvas');
    slide.appendChild(canvas);
    document.body.appendChild(slide);
    // JSDOM has no layout engine.
    slide.getBoundingClientRect = () => ({
      width: 960,
      height: 540,
      left: 0,
      top: 0,
      right: 960,
      bottom: 540,
      x: 0,
      y: 0,
    });
    canvas.getBoundingClientRect = () => ({
      width: 100,
      height: 100,
      left: 10,
      top: 10,
      right: 110,
      bottom: 110,
      x: 10,
      y: 10,
    });

    const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
    const bytes = new Uint8Array(await blob.arrayBuffer());
    document.body.removeChild(slide);

    const methods = entryMethods(bytes);
    const media = Object.keys(methods).filter((name) => name.startsWith('ppt/media/') && !name.endsWith('/'));
    expect(media).toHaveLength(1);
    expect(methods[media[0]]).toBe(0);
    expect(methods['ppt/slides/slide1.xml']).toBe(8);
    expect(methods['[Content_Types].xml']).toBe(8);

    const zip = await JSZip.loadAsync(bytes);
    expect(await zip.file(media[0]).async('base64')).toBe(PNG_BASE64);
  });
});
