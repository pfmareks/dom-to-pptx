import { describe, expect, it } from 'vitest';
import { textWraps } from '../utils.js';

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
