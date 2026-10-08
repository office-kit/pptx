// DrawingML color transforms (ECMA-376 §20.1.2.3) as resolveDrawingColor
// applies them, and therefore as the preview paints them.
//
// The first table holds colors the reference desktop app itself rendered, so it pins the
// formulas to the reference desktop app rather than to a restatement of them:
// - 2007 default theme gradient stops (tint + satMod) read from the reference desktop app on Mac
//   16.113's PNG exports of the Themed Style 1 / 2 tables (sRGB PNGs), at
//   rows that show the table background unfilled, 3 px from the stop's end;
// - Medium Style 2 - Accent 1 band fills in the 2023 default theme, read from
//   a native screenshot of the Table Design tab (Display P3, converted to
//   sRGB).
// Tolerance: 2 levels per channel, the spread of the sampling (a gradient
// sample sits ~1% from its stop; the P3 → sRGB conversion rounds).

import { describe, expect, it } from 'vitest';
import { resolveDrawingColor } from '../src/api/index.ts';
import { parseXml } from '../src/internal/xml/index.ts';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const resolve = (base: string, transforms: string): string | null =>
  resolveDrawingColor(
    parseXml(`<a:srgbClr xmlns:a="${A}" val="${base}">${transforms}</a:srgbClr>`).root,
    null,
  );

const channelDistance = (a: string, b: string): number =>
  Math.max(
    ...[1, 3, 5].map((i) =>
      Math.abs(Number.parseInt(a.slice(i, i + 2), 16) - Number.parseInt(b.slice(i, i + 2), 16)),
    ),
  );

const THEME_2007_ACCENTS = ['4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646'] as const;
const STOPS = {
  // fillStyleLst[2], first stop (Themed Style 1's table background)
  '<a:tint val="50000"/><a:satMod val="300000"/>': [
    '#A3C4FF',
    '#FFA2A1',
    '#DAFDA7',
    '#C9B5E8',
    '#9EEAFF',
    '#FFBE87',
  ],
  // fillStyleLst[3], last stop (Themed Style 2's table background)
  '<a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/>': [
    '#9BC1FF',
    '#FF9A99',
    '#DCFFA0',
    '#C8B0ED',
    '#95EEFF',
    '#FFB977',
  ],
  // fillStyleLst[3], first stop
  '<a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/>': [
    '#4081CE',
    '#D1413D',
    '#A1CB4B',
    '#805CAC',
    '#3AB8D8',
    '#FF932C',
  ],
} as const;

const REFERENCE_APP_RENDERED: ReadonlyArray<readonly [string, string, string]> = [
  ...Object.entries(STOPS).flatMap(([transforms, expected]) =>
    THEME_2007_ACCENTS.map(
      (accent, i) => [accent, transforms, expected[i]!] as [string, string, string],
    ),
  ),
  ['156082', '<a:tint val="40000"/>', '#CCD2D8'],
  ['156082', '<a:tint val="20000"/>', '#E6EAED'],
];
const TOLERANCE = 2;

describe('color transforms against colors the reference desktop app rendered', () => {
  it.each(REFERENCE_APP_RENDERED)('%s %s → %s', (base, transforms, expected) => {
    const actual = resolve(base, transforms)!;
    expect(channelDistance(actual, expected), `${actual} vs ${expected}`).toBeLessThanOrEqual(
      TOLERANCE,
    );
  });
});

describe('color transforms', () => {
  it('tints and shades in linear light', () => {
    expect(resolve('FF0000', '<a:tint val="50000"/>')).toBe('#FFBCBC');
    expect(resolve('FF0000', '<a:shade val="50000"/>')).toBe('#BC0000');
    expect(resolve('808080', '<a:tint val="0"/>')).toBe('#FFFFFF');
    expect(resolve('808080', '<a:shade val="0"/>')).toBe('#000000');
  });

  it('does not cap satMod at 100% saturation; the channels clamp instead', () => {
    // #C0E0A0: HSL(90°, 50%, 75%). ×300% → S 150%: the blue channel goes to
    // 2L − (L + S − LS) = 0.375 → #60, where a capped S gives #81.
    expect(resolve('C0E0A0', '<a:satMod val="300000"/>')).toBe('#C0FF60');
    // #FFD0B0 is already at S 100%; capping would leave it unchanged.
    expect(resolve('FFD0B0', '<a:satMod val="300000"/>')).toBe('#FFC161');
    expect(resolve('C08040', '<a:satMod val="50000"/>')).toBe('#A08060');
  });

  it('applies satOff, sat and saturation floors at gray', () => {
    expect(resolve('C08040', '<a:satOff val="-80000"/>')).toBe('#808080');
    expect(resolve('C0E0A0', '<a:satOff val="100000"/>')).toBe('#C0FF61');
    expect(resolve('C08040', '<a:sat val="0"/>')).toBe('#808080');
    // Gray has hue 0 (red); #80 is L 50.2%, a hair above the pure hue.
    expect(resolve('808080', '<a:sat val="100000"/>')).toBe('#FF0101');
  });

  it('modulates and offsets luminance in HSL, clamped to black and white', () => {
    // The reference desktop app's "Accent 1, Lighter 60%" and "Darker 25%".
    expect(resolve('4472C4', '<a:lumMod val="40000"/><a:lumOff val="60000"/>')).toBe('#B4C7E7');
    expect(resolve('4472C4', '<a:lumMod val="75000"/>')).toBe('#2F5597');
    expect(resolve('4472C4', '<a:lumOff val="100000"/>')).toBe('#FFFFFF');
    expect(resolve('4472C4', '<a:lumMod val="0"/>')).toBe('#000000');
    expect(resolve('FF0000', '<a:lum val="25000"/>')).toBe('#800000');
  });

  it('sets, modulates and offsets the hue', () => {
    expect(resolve('FF0000', '<a:hue val="7200000"/>')).toBe('#00FF00');
    expect(resolve('FF0000', '<a:hueOff val="14400000"/>')).toBe('#0000FF');
    expect(resolve('FF0000', '<a:hueOff val="-7200000"/>')).toBe('#0000FF');
    expect(resolve('00FF00', '<a:hueMod val="200000"/>')).toBe('#0000FF');
  });

  it('complements, inverts and grays', () => {
    expect(resolve('FF0000', '<a:comp/>')).toBe('#00FFFF');
    expect(resolve('FF0000', '<a:inv/>')).toBe('#00FFFF');
    expect(resolve('336699', '<a:inv/>')).toBe('#CC9966');
    expect(resolve('FF0000', '<a:gray/>')).toBe('#4D4D4D');
  });

  it('sets, modulates and offsets RGB components in linear light', () => {
    // The spec's examples: (00, FF, 00) → (FF, FF, 00) and (00, FF, FF).
    expect(resolve('00FF00', '<a:red val="100000"/>')).toBe('#FFFF00');
    expect(resolve('00FF00', '<a:blue val="100000"/>')).toBe('#00FFFF');
    expect(resolve('FF0000', '<a:green val="50000"/>')).toBe('#FFBC00');
    expect(resolve('FFFFFF', '<a:redMod val="50000"/>')).toBe('#BCFFFF');
    expect(resolve('FFFFFF', '<a:greenMod val="0"/>')).toBe('#FF00FF');
    expect(resolve('000000', '<a:blueOff val="50000"/>')).toBe('#0000BC');
    expect(resolve('00FF00', '<a:greenOff val="-100000"/>')).toBe('#000000');
  });

  it('shifts gamma with the sRGB curve, and invGamma undoes it', () => {
    expect(resolve('808080', '<a:gamma/>')).toBe('#BCBCBC');
    expect(resolve('808080', '<a:invGamma/>')).toBe('#373737');
    expect(resolve('336699', '<a:gamma/><a:invGamma/>')).toBe('#336699');
    // Legacy gradients wrap a shade in gamma/invGamma: an sRGB-space shade.
    expect(resolve('FF0000', '<a:gamma/><a:shade val="50000"/><a:invGamma/>')).toBe('#800000');
  });

  it('applies transforms in document order', () => {
    expect(resolve('FF0000', '<a:shade val="50000"/><a:inv/>')).toBe('#43FFFF');
    expect(resolve('FF0000', '<a:inv/><a:shade val="50000"/>')).toBe('#00BCBC');
  });

  it('leaves RGB alone for the alpha transforms', () => {
    expect(resolve('336699', '<a:alpha val="50000"/><a:alphaMod val="50000"/>')).toBe('#336699');
  });
});
