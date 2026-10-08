// The Bullets ▸ and Numbering ▸ galleries of the reference desktop app's (Mac) text menus, in
// their native order (captured 2026-10-07): None, then the built-in styles,
// five tiles per row. The reference desktop app draws the bullet glyphs from Wingdings; the
// editor writes the Unicode characters they show.
import type { BulletStyle } from '@office-kit/pptx';

export interface ListGalleryEntry {
  readonly label: string;
  readonly style: BulletStyle;
  /** The three markers the tile previews; empty for None. */
  readonly markers: readonly string[];
}

const bullet = (label: string, char: string): ListGalleryEntry => ({
  label,
  style: { char },
  markers: [char, char, char],
});

const numbering = (autoNum: string, markers: readonly string[]): ListGalleryEntry => ({
  label: markers.join(' '),
  style: { autoNum },
  markers,
});

export const BULLET_GALLERY: readonly ListGalleryEntry[] = [
  { label: 'None', style: 'none', markers: [] },
  bullet('Filled Round Bullets', '•'),
  bullet('Hollow Round Bullets', 'o'),
  bullet('Filled Square Bullets', '▪'),
  bullet('Hollow Square Bullets', '❑'),
  bullet('Star Bullets', '❖'),
  bullet('Arrow Bullets', '➢'),
  bullet('Checkmark Bullets', '✓'),
];

export const NUMBERING_GALLERY: readonly ListGalleryEntry[] = [
  { label: 'None', style: 'none', markers: [] },
  numbering('arabicPeriod', ['1.', '2.', '3.']),
  numbering('arabicParenR', ['1)', '2)', '3)']),
  numbering('romanUcPeriod', ['I.', 'II.', 'III.']),
  numbering('alphaUcPeriod', ['A.', 'B.', 'C.']),
  numbering('alphaLcParenR', ['a)', 'b)', 'c)']),
  numbering('alphaLcPeriod', ['a.', 'b.', 'c.']),
  numbering('romanLcPeriod', ['i.', 'ii.', 'iii.']),
];

// `setParagraphBullet` writes 'bullet' as • and 'number' as 1. 2. 3.
function normalize(style: BulletStyle | null): string {
  if (style === null || style === 'none') return 'none';
  if (style === 'bullet') return 'char:•';
  if (style === 'number') return 'autoNum:arabicPeriod';
  return 'char' in style ? `char:${style.char}` : `autoNum:${style.autoNum}`;
}

/** Whether every paragraph uses `entry`'s style (the tile the reference desktop app highlights). */
export function listStyleApplied(
  entry: ListGalleryEntry,
  bullets: readonly (BulletStyle | null)[],
): boolean {
  const key = normalize(entry.style);
  return bullets.length > 0 && bullets.every((style) => normalize(style) === key);
}
