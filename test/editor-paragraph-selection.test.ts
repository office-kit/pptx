import { describe, expect, it } from 'vitest';
import { paragraphsInTextRange } from '../site/src/lib/editor/core/paragraph-selection.ts';

describe('paragraph selection', () => {
  it('targets the caret paragraph including empty and final paragraphs', () => {
    for (const [offset, expected] of [
      [0, 0],
      [3, 0],
      [4, 1],
      [5, 2],
      [7, 2],
      [8, 3],
    ]) {
      expect(paragraphsInTextRange([3, 0, 2, 0], { start: offset!, end: offset! })).toEqual([
        expected,
      ]);
    }
  });
  it('excludes a paragraph when selection ends at its start', () => {
    expect(paragraphsInTextRange([3, 4, 5], { start: 1, end: 4 })).toEqual([0]);
    expect(paragraphsInTextRange([3, 4, 5], { start: 1, end: 5 })).toEqual([0, 1]);
  });
  it('includes empty paragraphs spanned by a selection', () => {
    expect(paragraphsInTextRange([3, 0, 2], { start: 2, end: 6 })).toEqual([0, 1, 2]);
  });
});
