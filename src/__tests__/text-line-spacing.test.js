import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import { getTextStyle } from '../utils.js';
import { exportToPptx } from '../index.js';

const DRAWINGML_PERCENT_SCALE = 100000;
const PPTX_SINGLE_SPACING_BASIS = 1.2;
const SCALE = 0.5;

beforeAll(() => {
  let fillStyle = '';
  HTMLCanvasElement.prototype.getContext = () => ({
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(value) {
      fillStyle = value;
    },
    clearRect: () => {},
    fillRect: () => {},
    getImageData: () => ({ data: [0, 0, 0, 0] }),
  });
});

// The minimal computed-style shape getTextStyle reads.
function computedStyle({ fontSizePx = 16, lineHeight = '25px', whiteSpace = 'normal' } = {}) {
  return {
    color: '#000000',
    opacity: '1',
    webkitBackgroundClip: 'border-box',
    backgroundClip: 'border-box',
    backgroundImage: 'none',
    fontSize: `${fontSizePx}px`,
    lineHeight,
    whiteSpace,
    marginTop: '0px',
    marginBottom: '0px',
    fontFamily: 'Arial, sans-serif',
    fontWeight: '400',
    fontStyle: 'normal',
    textDecoration: 'none',
    backgroundColor: 'transparent',
    letterSpacing: 'normal',
    getPropertyValue: () => '',
  };
}

async function serialize(style) {
  const options = getTextStyle(style, SCALE);
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  slide.addText([{ text: 'A paragraph of text.', options: { ...options } }], {
    x: 1,
    y: 1,
    w: 2,
    h: 1,
    margin: 0,
  });

  const zip = await JSZip.loadAsync(await pptx.write({ outputType: 'nodebuffer' }));
  return { options, xml: await zip.file('ppt/slides/slide1.xml').async('string') };
}

describe('line spacing', () => {
  it('serializes wrapping text as relative spacing, floored to DrawingML precision', async () => {
    const fontSizePx = 14;
    const lineHeightPx = 25;
    const { options, xml } = await serialize(computedStyle({ fontSizePx, lineHeight: `${lineHeightPx}px` }));

    // 25 / 14 = 1.7857em of line box, over the 1.2em "single spacing" basis.
    const expectedSpcPct = Math.floor((lineHeightPx / fontSizePx / PPTX_SINGLE_SPACING_BASIS) * DRAWINGML_PERCENT_SCALE);

    expect(options.lineSpacing).toBeUndefined();
    expect(options.lineSpacingMultiple).toBe(expectedSpcPct / DRAWINGML_PERCENT_SCALE);
    expect(xml).toContain(`<a:lnSpc><a:spcPct val="${expectedSpcPct}"/></a:lnSpc>`);
    expect(xml).not.toContain('<a:lnSpc><a:spcPts');
  });

  it('reproduces the browser-measured line height within serialization rounding', async () => {
    const lineHeightPx = 25;
    const { options } = await serialize(computedStyle({ fontSizePx: 16, lineHeight: `${lineHeightPx}px` }));

    const browserPointHeight = lineHeightPx * 0.75 * SCALE;
    const relativePointHeight = options.fontSize * PPTX_SINGLE_SPACING_BASIS * options.lineSpacingMultiple;

    expect(relativePointHeight).toBeCloseTo(browserPointHeight, 4);
  });

  it('keeps exact point spacing for no-wrap text, which cannot re-wrap', async () => {
    const lineHeightPx = 25;
    const { options, xml } = await serialize(
      computedStyle({ lineHeight: `${lineHeightPx}px`, whiteSpace: 'nowrap' })
    );

    expect(options.lineSpacing).toBe(lineHeightPx * 0.75 * SCALE);
    expect(options.lineSpacingMultiple).toBeUndefined();
    expect(xml).toContain('<a:lnSpc><a:spcPts');
    expect(xml).not.toContain('<a:lnSpc><a:spcPct');
  });

  it('emits no line spacing at all for line-height: normal', async () => {
    const { options, xml } = await serialize(computedStyle({ lineHeight: 'normal' }));

    expect(options.lineSpacing).toBeUndefined();
    expect(options.lineSpacingMultiple).toBeUndefined();
    expect(xml).not.toContain('<a:lnSpc>');
  });

  it('resolves a unitless line-height against the font size', async () => {
    const { options } = await serialize(computedStyle({ fontSizePx: 20, lineHeight: '1.5' }));

    // 1.5em over the 1.2em basis
    expect(options.lineSpacingMultiple).toBe(
      Math.floor((1.5 / PPTX_SINGLE_SPACING_BASIS) * DRAWINGML_PERCENT_SCALE) / DRAWINGML_PERCENT_SCALE
    );
  });

  it('drops line spacing entirely for vertical writing modes, which PPTX lays out differently', async () => {
    const boxPx = { left: 200, top: 200, width: 60, height: 400 };
    const measured = (box) => () => ({
      x: box.left,
      y: box.top,
      left: box.left,
      top: box.top,
      width: box.width,
      height: box.height,
      right: box.left + box.width,
      bottom: box.top + box.height,
      toJSON() {
        return this;
      },
    });

    const slide = document.createElement('div');
    slide.setAttribute('style', 'position:relative;width:1920px;height:1080px;background:#fff');
    slide.getBoundingClientRect = measured({ left: 0, top: 0, width: 1920, height: 1080 });

    const vertical = document.createElement('div');
    vertical.setAttribute(
      'style',
      `position:absolute;left:${boxPx.left}px;top:${boxPx.top}px;width:${boxPx.width}px;height:${boxPx.height}px;` +
        'color:#111;font-size:16px;line-height:25px;white-space:normal;writing-mode:vertical-rl'
    );
    vertical.textContent = 'Text laid out top to bottom.';
    vertical.getBoundingClientRect = measured(boxPx);

    slide.appendChild(vertical);
    document.body.appendChild(slide);

    try {
      const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
      const zip = await JSZip.loadAsync(blob);
      const xml = await zip.file('ppt/slides/slide1.xml').async('string');

      expect(xml).toContain('Text laid out top to bottom.');
      expect(xml).not.toContain('<a:lnSpc>');
    } finally {
      slide.remove();
    }
  });
});
