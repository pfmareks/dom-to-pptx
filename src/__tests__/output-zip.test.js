// src/__tests__/output-zip.test.js
import { describe, it, expect, beforeAll } from 'vitest';
import JSZip from 'jszip';
import { exportToPptx } from '../index.js';

// Two distinct 1x1 PNGs.
const RED_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';
const BLUE_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGOQkZIGAACnAFJfxMSBAAAAAElFTkSuQmCC';

// JSDOM has no layout engine.
function stubRect(node, left, top, width, height) {
  node.getBoundingClientRect = () => ({
    width,
    height,
    left,
    top,
    right: left + width,
    bottom: top + height,
    x: left,
    y: top,
  });
}

// Exports one slide holding a <canvas> per PNG, left to right, and returns the .pptx bytes.
async function exportCanvases(pngs, options = {}) {
  const slide = document.createElement('div');
  document.body.appendChild(slide);
  stubRect(slide, 0, 0, 960, 540);
  pngs.forEach((png, i) => {
    const canvas = document.createElement('canvas');
    canvas.toDataURL = () => `data:image/png;base64,${png}`;
    slide.appendChild(canvas);
    stubRect(canvas, 10 + i * 200, 10, 100, 100);
  });
  try {
    const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false, ...options });
    return new Uint8Array(await blob.arrayBuffer());
  } finally {
    document.body.removeChild(slide);
  }
}

// Base64 content of every PNG media entry, in entry-name order; fails on any CRC mismatch.
async function mediaPngs(bytes) {
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const names = Object.keys(zip.files)
    .filter((name) => /^ppt\/media\/.+\.png$/.test(name))
    .sort();
  return Promise.all(names.map((name) => zip.file(name).async('base64')));
}

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

describe('exported package media', () => {
  beforeAll(() => {
    // Colour parsing needs a 2D context, which JSDOM does not provide.
    HTMLCanvasElement.prototype.getContext = () => ({
      fillStyle: '',
      clearRect: () => {},
      fillRect: () => {},
      getImageData: () => ({ data: [0, 0, 0, 0] }),
    });
  });

  it("writes each image's own bytes, stored as-is, and deflates the XML parts", async () => {
    const bytes = await exportCanvases([RED_PNG, BLUE_PNG]);

    expect(await mediaPngs(bytes)).toEqual([RED_PNG, BLUE_PNG]);
    const methods = entryMethods(bytes);
    for (const name of Object.keys(methods).filter((n) => /^ppt\/media\/.+\.png$/.test(n))) {
      expect(methods[name]).toBe(0);
    }
    expect(methods['ppt/slides/slide1.xml']).toBe(8);
    expect(methods['[Content_Types].xml']).toBe(8);
  });

  it('stores an image placed twice once, with both pictures pointing at it', async () => {
    const bytes = await exportCanvases([RED_PNG, BLUE_PNG, RED_PNG]);

    expect(await mediaPngs(bytes)).toEqual([RED_PNG, BLUE_PNG]);
    const zip = await JSZip.loadAsync(bytes);
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('string');
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
    const targets = Object.fromEntries(
      Array.from(relsXml.matchAll(/Id="(rId\d+)"[^>]*Target="\.\.\/media\/([^"]+)"/g), (m) => [m[1], m[2]])
    );
    const pictures = await Promise.all(
      Array.from(slideXml.matchAll(/r:embed="(rId\d+)"/g), (m) =>
        zip.file(`ppt/media/${targets[m[1]]}`).async('base64')
      )
    );
    expect(pictures).toEqual([RED_PNG, BLUE_PNG, RED_PNG]);
  });

  it('keeps the image bytes when skipNormalize returns the PptxGenJS output as-is', async () => {
    const bytes = await exportCanvases([RED_PNG, BLUE_PNG], { skipNormalize: true });

    expect(await mediaPngs(bytes)).toEqual([RED_PNG, BLUE_PNG]);
  });

  it('writes the image bytes in browsers without Uint8Array.fromBase64', async () => {
    const fromBase64 = Uint8Array.fromBase64;
    delete Uint8Array.fromBase64;
    try {
      const bytes = await exportCanvases([RED_PNG, BLUE_PNG]);

      expect(await mediaPngs(bytes)).toEqual([RED_PNG, BLUE_PNG]);
    } finally {
      if (fromBase64) Uint8Array.fromBase64 = fromBase64;
    }
  });
});
