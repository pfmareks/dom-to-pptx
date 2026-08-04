import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { exportToPptx } from '../index.js';

// A 1920px-wide root is exported as a 10in slide, so 1 CSS px is 914400 / 96 / 2 EMU.
const EMU_PER_PX = 4762.5;
const SLIDE_WIDTH_PX = 1920;
const SLIDE_HEIGHT_PX = 1080;

function rect({ left, top, width, height }) {
  return {
    x: left,
    y: top,
    left,
    top,
    width,
    height,
    right: left + width,
    bottom: top + height,
    toJSON() {
      return this;
    },
  };
}

// Builds a slide root, stubs every element's measured rectangle from its authored
// px geometry, and returns the serialized slide XML.
async function exportSlide(elements) {
  const slide = document.createElement('div');
  slide.setAttribute('style', `position:relative;width:${SLIDE_WIDTH_PX}px;height:${SLIDE_HEIGHT_PX}px;background:#fff`);
  slide.getBoundingClientRect = () => rect({ left: 0, top: 0, width: SLIDE_WIDTH_PX, height: SLIDE_HEIGHT_PX });

  for (const { node, box } of elements) {
    node.getBoundingClientRect = () => rect(box);
    slide.appendChild(node);
  }
  document.body.appendChild(slide);

  try {
    const blob = await exportToPptx(slide, { skipDownload: true, autoEmbedFonts: false });
    const zip = await JSZip.loadAsync(blob);
    return await zip.file('ppt/slides/slide1.xml').async('string');
  } finally {
    slide.remove();
  }
}

function div(style, text) {
  const node = document.createElement('div');
  node.setAttribute('style', `position:absolute;${style}`);
  node.textContent = text;
  return node;
}

function shapes(xml) {
  return Array.from(new DOMParser().parseFromString(xml, 'text/xml').getElementsByTagName('p:sp'));
}

function shapeWithText(xml, text) {
  const shape = shapes(xml).find((candidate) =>
    Array.from(candidate.getElementsByTagName('a:t')).some((run) => run.textContent.includes(text))
  );
  expect(shape, `no shape carries the text "${text}"`).toBeDefined();
  return shape;
}

function geometry(shape) {
  const offset = shape.getElementsByTagName('a:off')[0];
  const extent = shape.getElementsByTagName('a:ext')[0];
  return {
    x: Number(offset.getAttribute('x')),
    y: Number(offset.getAttribute('y')),
    width: Number(extent.getAttribute('cx')),
    height: Number(extent.getAttribute('cy')),
  };
}

function autofit(shape) {
  const bodyPr = shape.getElementsByTagName('a:bodyPr')[0];
  if (bodyPr.getElementsByTagName('a:normAutofit').length) return 'normAutofit';
  if (bodyPr.getElementsByTagName('a:spAutoFit').length) return 'spAutoFit';
  return 'none';
}

function wrapMode(shape) {
  return shape.getElementsByTagName('a:bodyPr')[0].getAttribute('wrap');
}

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
    getImageData: () => ({ data: [0, 0, 0, 255] }),
  });
});

describe('text box autofit', () => {
  it('shrinks wrapping text to fit and leaves no-wrap text unfitted', async () => {
    const paragraph = div(
      'left:200px;top:200px;width:600px;height:200px;color:#111;font-size:16px;line-height:24px;white-space:normal',
      'A wrapping paragraph long enough to occupy several lines inside its authored box.'
    );
    const label = div(
      'left:200px;top:600px;width:200px;height:20px;color:#111;font-size:12px;line-height:20px;white-space:nowrap',
      'A single measured line'
    );

    const xml = await exportSlide([
      { node: paragraph, box: { left: 200, top: 200, width: 600, height: 200 } },
      { node: label, box: { left: 200, top: 600, width: 200, height: 20 } },
    ]);

    // Wrapping text: shrink-to-fit, so a re-wrap in another renderer stays inside the box.
    expect(autofit(shapeWithText(xml, 'A wrapping paragraph'))).toBe('normAutofit');
    // No-wrap text: one measured line, nothing to fit.
    expect(autofit(shapeWithText(xml, 'A single measured line'))).toBe('none');
  });
});
