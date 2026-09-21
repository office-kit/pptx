import { describe, expect, it } from 'vitest';
import {
  textFormatActive,
  toggleTextFormat,
} from '../site/src/lib/editor/core/text-format-toggle.ts';

describe('text format toggles', () => {
  it('recognizes imported strike styles and turns mixed selections on', () => {
    expect(toggleTextFormat([{ strike: 'dblStrike' }], 'strike')).toEqual({ strike: false });
    expect(textFormatActive([{ strike: 'noStrike' }], 'strike')).toBe(false);
    expect(toggleTextFormat([{ strike: true }, {}], 'strike')).toEqual({ strike: true });
    expect(toggleTextFormat([{ underline: 'wavy' }], 'underline')).toEqual({ underline: false });
  });
  it('toggles arbitrary imported baseline offsets and switches between scripts', () => {
    expect(toggleTextFormat([{ baseline: 0.42 }], 'superscript')).toEqual({ baseline: 0 });
    expect(toggleTextFormat([{ baseline: -0.1 }], 'subscript')).toEqual({ baseline: 0 });
    expect(toggleTextFormat([{ baseline: 0.3 }], 'subscript')).toEqual({ baseline: -0.25 });
    expect(toggleTextFormat([{ baseline: -0.25 }], 'superscript')).toEqual({ baseline: 0.3 });
    expect(toggleTextFormat([{ baseline: 0.3 }, {}], 'superscript')).toEqual({ baseline: 0.3 });
    expect(textFormatActive([], 'subscript')).toBe(false);
  });
});
