// The CSS the renderer's HTML text path and the editor's inline editing share
// for `<a:bodyPr vert=… numCol=… spcCol=…>`.

import { describe, expect, it } from 'vitest';
import { textColumnsStyle, verticalTextStyle } from '../packages/preview/src/text-body-style.ts';

describe('vertical writing as CSS', () => {
  it('turns each authored direction into a writing mode', () => {
    expect(verticalTextStyle('vert')).toEqual({
      declarations: 'writing-mode:vertical-rl',
      transform: '',
    });
    // eaVert is the same rotated reading direction, named for East-Asian text.
    expect(verticalTextStyle('eaVert')).toEqual(verticalTextStyle('vert'));
    expect(verticalTextStyle('mongolianVert')).toEqual({
      declarations: 'writing-mode:vertical-lr',
      transform: '',
    });
    expect(verticalTextStyle('wordArtVert')).toEqual({
      declarations: 'writing-mode:vertical-rl;text-orientation:upright',
      transform: '',
    });
    expect(verticalTextStyle('wordArtVertRtl')).toEqual({
      declarations: 'writing-mode:vertical-rl;text-orientation:upright;direction:rtl',
      transform: '',
    });
  });

  it('reports vert270’s half turn separately from its writing mode', () => {
    // No writing-mode reads bottom-to-top, so the turn is a transform — and it
    // is reported apart so a caller that already rotates the box can compose
    // the two instead of dropping one.
    expect(verticalTextStyle('vert270')).toEqual({
      declarations: 'writing-mode:vertical-lr',
      transform: 'rotate(180deg)',
    });
  });

  it('says nothing for horizontal text', () => {
    expect(verticalTextStyle(null)).toEqual({ declarations: '', transform: '' });
  });
});

describe('text columns as CSS', () => {
  it('splits the body and spaces the columns', () => {
    expect(textColumnsStyle({ count: 3, gapEmu: 190500 })).toBe(
      'column-count:3;column-gap:20.00px',
    );
  });

  it('falls back to PowerPoint’s gap when spcCol is absent', () => {
    expect(textColumnsStyle({ count: 2 })).toBe('column-count:2;column-gap:12px');
  });

  it('says nothing below two columns, which is not a split at all', () => {
    expect(textColumnsStyle(null)).toBe('');
    expect(textColumnsStyle({ count: 1, gapEmu: 190500 })).toBe('');
  });
});
