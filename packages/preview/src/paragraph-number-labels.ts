import type { BulletStyle } from '@office-kit/pptx';

// Maps a `BulletStyle` value to the underlying `ST_TextAutoNumberScheme`
// token (or `null` when the paragraph isn't auto-numbered). `'number'`
// is the shorthand for arabicPeriod that setShapeBullets uses.
const bulletAutoNumType = (style: BulletStyle | null): string | null => {
  if (style === 'number') return 'arabicPeriod';
  if (style !== null && typeof style === 'object' && 'autoNum' in style) {
    return style.autoNum ?? null;
  }
  return null;
};

const toRoman = (n: number): string => {
  if (n <= 0) return String(n);
  const map: ReadonlyArray<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let out = '';
  let r = n;
  for (const [v, s] of map) {
    while (r >= v) {
      out += s;
      r -= v;
    }
  }
  return out;
};

const toAlpha = (n: number): string => {
  // 1 -> A, 26 -> Z, 27 -> AA, etc.
  if (n <= 0) return String(n);
  let r = n;
  let out = '';
  while (r > 0) {
    r -= 1;
    out = String.fromCharCode(65 + (r % 26)) + out;
    r = Math.floor(r / 26);
  }
  return out;
};

// Format an auto-number per ECMA-376 §17.18.96 `ST_TextAutoNumberScheme`.
// Only the most common variants are implemented; unknown tokens fall back
// to arabicPeriod-style formatting.
const formatAutoNum = (token: string, n: number): string => {
  const arabic = String(n);
  switch (token) {
    case 'arabicPlain':
      return arabic;
    case 'arabicPeriod':
      return `${arabic}.`;
    case 'arabicParenR':
      return `${arabic})`;
    case 'arabicParenBoth':
      return `(${arabic})`;
    case 'romanUcPeriod':
      return `${toRoman(n)}.`;
    case 'romanLcPeriod':
      return `${toRoman(n).toLowerCase()}.`;
    case 'romanUcParenR':
      return `${toRoman(n)})`;
    case 'romanLcParenR':
      return `${toRoman(n).toLowerCase()})`;
    case 'romanUcParenBoth':
      return `(${toRoman(n)})`;
    case 'romanLcParenBoth':
      return `(${toRoman(n).toLowerCase()})`;
    case 'alphaUcPeriod':
      return `${toAlpha(n)}.`;
    case 'alphaLcPeriod':
      return `${toAlpha(n).toLowerCase()}.`;
    case 'alphaUcParenR':
      return `${toAlpha(n)})`;
    case 'alphaLcParenR':
      return `${toAlpha(n).toLowerCase()})`;
    case 'alphaUcParenBoth':
      return `(${toAlpha(n)})`;
    case 'alphaLcParenBoth':
      return `(${toAlpha(n).toLowerCase()})`;
    default:
      return `${arabic}.`;
  }
};

/** Resolve list numbering consistently for rendered slides and editing surfaces. */
export const paragraphNumberLabels = (
  paraData: ReadonlyArray<{ bulletStyle: BulletStyle | null; level: number }>,
): Array<string | null> => {
  // Numbering pre-pass — assign an autonum index per paragraph. PowerPoint
  // keeps one counter per indent level: a nested list (level 1) between two
  // level-0 items does not restart the outer list, so "1. / a. / b. / 2."
  // renders as such. A paragraph resets the counters of every deeper level;
  // a non-numbered paragraph also resets its own level, and a different
  // numbering scheme at the same level starts over at 1.
  const numberLabels: Array<string | null> = Array.from({ length: paraData.length }, () => null);
  {
    const counters: number[] = [];
    const types: Array<string | null> = [];
    for (let i = 0; i < paraData.length; i++) {
      const para = paraData[i]!;
      const num = bulletAutoNumType(para.bulletStyle);
      const level = Math.max(0, para.level);
      for (let l = num === null ? level : level + 1; l < counters.length; l++) {
        counters[l] = 0;
        types[l] = null;
      }
      if (num === null) continue;
      if (types[level] !== num) {
        counters[level] = 1;
        types[level] = num;
      } else {
        counters[level] = (counters[level] ?? 0) + 1;
      }
      numberLabels[i] = formatAutoNum(num, counters[level]!);
    }
  }

  return numberLabels;
};
