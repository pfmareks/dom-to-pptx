import { describe, expect, it } from 'vitest';
import { textWrapOptions, textWraps } from '../utils.js';

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
