// src/__tests__/border-svg.test.js
import { describe, it, expect } from 'vitest';
import { generateCompositeBorderSVG } from '../utils.js';

const sides = {
  top: { width: 2, color: 'E0DDDA', opacity: 1 },
  right: { width: 0, color: null },
  bottom: { width: 1, color: '1C1A1B', opacity: 0.5 },
  left: { width: 0, color: null },
};

describe('generateCompositeBorderSVG', () => {
  // PptxGenJS rasterizes a PNG preview per SVG image; the browser only reuses a decoded image when
  // the data URL is identical, so identical borders must produce identical data.
  it('returns identical data for identical borders', () => {
    expect(generateCompositeBorderSVG(369, 67, 8, sides)).toBe(generateCompositeBorderSVG(369, 67, 8, sides));
  });

  it('clips the border rects with the clip path it defines', () => {
    const data = generateCompositeBorderSVG(369, 67, 8, sides);
    const svg = atob(data.slice('data:image/svg+xml;base64,'.length));
    const id = svg.match(/<clipPath id="([^"]+)"/)[1];
    expect(svg).toContain(`<g clip-path="url(#${id})">`);
  });
});
