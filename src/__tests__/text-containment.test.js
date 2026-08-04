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

// Shape-level fill, i.e. what is painted behind the text - not the run colors,
// which are solidFill elements inside <a:rPr>.
function hasShapeFill(shape) {
  const shapeProperties = shape.getElementsByTagName('p:spPr')[0];
  return Array.from(shapeProperties.childNodes).some((child) => child.nodeName === 'a:solidFill');
}

function preset(shape) {
  return shape.getElementsByTagName('a:prstGeom')[0].getAttribute('prst');
}

// Serialization rounds inches to whole EMU, so geometry that should coincide can
// differ by a unit.
function expectSameEmu(actual, expected) {
  expect(Math.abs(actual - expected)).toBeLessThanOrEqual(1);
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

describe('no-wrap width slack', () => {
  const NO_WRAP_SLACK_RATIO = 0.06;

  it('widens a no-wrap label and keeps a wrapping paragraph at its measured width', async () => {
    const labelPx = { left: 200, top: 600, width: 200, height: 20 };
    const paragraphPx = { left: 200, top: 200, width: 600, height: 200 };

    const label = div(
      `left:${labelPx.left}px;top:${labelPx.top}px;width:${labelPx.width}px;height:${labelPx.height}px;` +
        'color:#111;font-size:12px;line-height:20px;white-space:nowrap;text-align:left',
      'A single measured line'
    );
    const paragraph = div(
      `left:${paragraphPx.left}px;top:${paragraphPx.top}px;width:${paragraphPx.width}px;height:${paragraphPx.height}px;` +
        'color:#111;font-size:16px;line-height:24px;white-space:normal;overflow-wrap:anywhere',
      'A wrapping paragraph long enough to occupy several lines inside its authored box.'
    );

    const xml = await exportSlide([
      { node: label, box: labelPx },
      { node: paragraph, box: paragraphPx },
    ]);

    const labelShape = shapeWithText(xml, 'A single measured line');
    expect(wrapMode(labelShape)).toBe('none');
    // Left-anchored, so the box grows to the right and x is unmoved.
    expect(geometry(labelShape).x).toBe(Math.round(labelPx.left * EMU_PER_PX));
    expect(geometry(labelShape).width).toBe(Math.round(labelPx.width * (1 + NO_WRAP_SLACK_RATIO) * EMU_PER_PX));

    const paragraphShape = shapeWithText(xml, 'A wrapping paragraph');
    expect(wrapMode(paragraphShape)).toBe('square');
    // Wrapping text keeps its exact measured rectangle, so it cannot cross the
    // right edge the browser laid it out against.
    expect(geometry(paragraphShape).width).toBe(Math.round(paragraphPx.width * EMU_PER_PX));
    expect(geometry(paragraphShape).x + geometry(paragraphShape).width).toBeLessThanOrEqual(
      Math.round((paragraphPx.left + paragraphPx.width) * EMU_PER_PX)
    );
  });

  it('keeps a centered no-wrap label centered and a right-aligned one right-anchored', async () => {
    const centeredPx = { left: 400, top: 300, width: 300, height: 20 };
    const rightPx = { left: 1200, top: 300, width: 300, height: 20 };

    const centered = div(
      `left:${centeredPx.left}px;top:${centeredPx.top}px;width:${centeredPx.width}px;height:${centeredPx.height}px;` +
        'color:#111;font-size:12px;line-height:20px;white-space:nowrap;text-align:center',
      'Centered label'
    );
    const rightAligned = div(
      `left:${rightPx.left}px;top:${rightPx.top}px;width:${rightPx.width}px;height:${rightPx.height}px;` +
        'color:#111;font-size:12px;line-height:20px;white-space:nowrap;text-align:right',
      'Right aligned label'
    );

    const xml = await exportSlide([
      { node: centered, box: centeredPx },
      { node: rightAligned, box: rightPx },
    ]);

    const centeredShape = geometry(shapeWithText(xml, 'Centered label'));
    const authoredCenter = (centeredPx.left + centeredPx.width / 2) * EMU_PER_PX;
    expectSameEmu(centeredShape.x + centeredShape.width / 2, authoredCenter);

    const rightShape = geometry(shapeWithText(xml, 'Right aligned label'));
    const authoredRightEdge = (rightPx.left + rightPx.width) * EMU_PER_PX;
    expectSameEmu(rightShape.x + rightShape.width, authoredRightEdge);

    // Both were actually widened - otherwise the anchor assertions above hold trivially.
    for (const shape of [centeredShape, rightShape]) {
      expect(shape.width).toBeGreaterThan(Math.round(centeredPx.width * EMU_PER_PX));
    }
  });

  it('splits no-wrap text out of a filled shape, leaving the shape at its measured size', async () => {
    const badgePx = { left: 100, top: 100, width: 300, height: 40 };

    const badge = div(
      `left:${badgePx.left}px;top:${badgePx.top}px;width:${badgePx.width}px;height:${badgePx.height}px;` +
        'color:#fff;background-color:#123456;border-radius:16px;font-size:12px;line-height:16px;' +
        'white-space:nowrap;text-align:center;padding:8px 18px',
      'BADGE'
    );

    const xml = await exportSlide([{ node: badge, box: badgePx }]);

    const textShape = shapeWithText(xml, 'BADGE');
    const textGeometry = geometry(textShape);

    // The text moved into its own box: a plain rectangle that paints nothing of its own,
    // and is the only shape allowed to take the slack.
    expect(hasShapeFill(textShape)).toBe(false);
    expect(preset(textShape)).toBe('rect');
    expect(textGeometry.width).toBe(Math.round(badgePx.width * (1 + NO_WRAP_SLACK_RATIO) * EMU_PER_PX));

    // The visible badge is a separate, textless round rect at exactly the measured geometry.
    const visible = shapes(xml).find(
      (shape) =>
        !shape.getElementsByTagName('a:t').length &&
        preset(shape) === 'roundRect' &&
        geometry(shape).x === Math.round(badgePx.left * EMU_PER_PX)
    );
    expect(visible, 'no textless round-rect companion at the authored x').toBeDefined();
    expect(geometry(visible)).toEqual({
      x: Math.round(badgePx.left * EMU_PER_PX),
      y: Math.round(badgePx.top * EMU_PER_PX),
      width: Math.round(badgePx.width * EMU_PER_PX),
      height: Math.round(badgePx.height * EMU_PER_PX),
    });
    expect(hasShapeFill(visible)).toBe(true);

    // The two stay glued: same vertical band, same center, so the text reads as the badge's own.
    expect(textGeometry.y).toBe(geometry(visible).y);
    expect(textGeometry.height).toBe(geometry(visible).height);
    expectSameEmu(textGeometry.x + textGeometry.width / 2, geometry(visible).x + geometry(visible).width / 2);

    // The authored padding survives as insets - the slack did not carve into it.
    const bodyPr = textShape.getElementsByTagName('a:bodyPr')[0];
    const insetEmu = (px) => Math.round(px * 0.75 * 0.5 * 12700);
    expect(Number(bodyPr.getAttribute('lIns'))).toBe(insetEmu(18));
    expect(Number(bodyPr.getAttribute('rIns'))).toBe(insetEmu(18));
    expect(Number(bodyPr.getAttribute('tIns'))).toBe(insetEmu(8));
    expect(Number(bodyPr.getAttribute('bIns'))).toBe(insetEmu(8));
  });
});
