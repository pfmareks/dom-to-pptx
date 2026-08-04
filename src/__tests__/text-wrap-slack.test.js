import { describe, expect, it } from 'vitest';
import { textWrapOptions, textWraps, withNoWrapWidthSlack } from '../utils.js';

describe('textWraps', () => {
  it('is false for the white-space values that keep text on one measured line', () => {
    expect(textWraps({ whiteSpace: 'nowrap' })).toBe(false);
    expect(textWraps({ whiteSpace: 'pre' })).toBe(false);
  });

  it('is true for the white-space values that soft-wrap', () => {
    expect(textWraps({ whiteSpace: 'normal' })).toBe(true);
    expect(textWraps({ whiteSpace: 'pre-wrap' })).toBe(true);
    expect(textWraps({ whiteSpace: 'pre-line' })).toBe(true);
  });
});

describe('textWrapOptions', () => {
  it('gives wrapping text shrink-to-fit autofit', () => {
    expect(textWrapOptions({ whiteSpace: 'normal' })).toEqual({ wrap: true, fit: 'shrink' });
  });

  it('gives no-wrap text no autofit at all', () => {
    expect(textWrapOptions({ whiteSpace: 'nowrap' })).toEqual({ wrap: false, fit: 'none' });
  });
});

describe('withNoWrapWidthSlack', () => {
  // A 1in wide no-wrap box, so the 6% ratio governs and the 0.02in floor does not.
  const RATIO = 0.06;
  const FLOOR_IN = 0.02;
  const box = (overrides = {}) => ({ x: 1, y: 1, w: 1, h: 0.2, wrap: false, ...overrides });

  it('leaves wrapping text alone: it must stay inside its measured rectangle', () => {
    const options = box({ wrap: true });
    expect(withNoWrapWidthSlack(options)).toBe(options);
  });

  it('leaves rotated, vertical, and zero-width boxes alone', () => {
    for (const options of [box({ rotate: 90 }), box({ vert: 'eaVert' }), box({ w: 0 })]) {
      expect(withNoWrapWidthSlack(options)).toBe(options);
    }
  });

  it('widens a left-anchored box to the right, leaving x where it was', () => {
    const out = withNoWrapWidthSlack(box());
    expect(out.w).toBeCloseTo(1 + RATIO);
    expect(out.x).toBe(1);
  });

  it('grows a centered box evenly around its center', () => {
    const out = withNoWrapWidthSlack(box({ align: 'center' }));
    expect(out.w).toBeCloseTo(1 + RATIO);
    expect(out.x).toBeCloseTo(1 - RATIO / 2);
    // the center of the box is unmoved
    expect(out.x + out.w / 2).toBeCloseTo(1 + 0.5);
  });

  it('grows a right-anchored box to the left, leaving its right edge in place', () => {
    const out = withNoWrapWidthSlack(box({ align: 'right' }));
    expect(out.x).toBeCloseTo(1 - RATIO);
    expect(out.x + out.w).toBeCloseTo(1 + 1);
  });

  it('applies an absolute floor so tiny boxes get usable slack', () => {
    // 6% of 0.1in is 0.006in, well under the floor
    const out = withNoWrapWidthSlack(box({ w: 0.1 }));
    expect(out.w).toBeCloseTo(0.1 + FLOOR_IN);
  });

  it('never touches the text insets', () => {
    const margin = [6.75, 6.75, 3, 3];
    expect(withNoWrapWidthSlack(box({ margin })).margin).toEqual(margin);
  });
});
