// Mac PowerPoint's WordArt gallery: its twenty presets in gallery order, each
// the character payload PowerPoint wrote when the preset was applied
// (test/fixtures/native/wordart-*-shape.xml, catalogued in wordart-capture.md).

import {
  asColor,
  getShapeParagraphCount,
  getShapeRunCount,
  getShapeRunFormat,
  setShapeTextFormat,
  type Color,
  type ColorTransform,
  type PresentationTheme,
  type SlideShapeData,
  type TextFormat,
} from '@office-kit/pptx';
import { resolveColor } from './theme-color.ts';

/**
 * A color with PowerPoint's tint transforms. The text-format API takes
 * transforms only on gradient stops and pattern colors, so a transformed
 * solid fill, outline or effect color is written as the RGB it paints in this
 * deck's theme: the look matches native, the theme link does not.
 */
export type Tone = (color: Color, ...transforms: ColorTransform[]) => Color;

export interface WordArtPreset {
  /** PowerPoint's gallery label, also the swatch's tooltip and accessible name. */
  readonly label: string;
  readonly format: (tone: Tone) => TextFormat;
  /** Why the swatch is shown but cannot be applied here. */
  readonly unavailable?: string;
}

const lighter = (mod: number, off: number): ColorTransform[] => [
  { kind: 'lumMod', value: mod },
  { kind: 'lumOff', value: off },
];
const darker = (mod: number): ColorTransform[] => [{ kind: 'lumMod', value: mod }];

const NO_OUTLINE = { widthEmu: 0 } as const;
// The bevel presets write `<a:ln/>` with neither width nor color.
const EMPTY_OUTLINE = {} as const;
// The "Hard Shadow" presets: a short offset down and to the right, no blur.
const HARD = { offsetEmu: 38100, angleDeg: 45, alignment: 'tl', rotateWithShape: false } as const;
// The pattern presets' hard shadow points at 44°, not 45° (dir="2640000").
const PATTERN_HARD = {
  offsetEmu: 38100,
  angleDeg: 44,
  alignment: 'bl',
  rotateWithShape: false,
} as const;
const SOFT = {
  blurEmu: 38100,
  offsetEmu: 19050,
  angleDeg: 45,
  alignment: 'tl',
  rotateWithShape: false,
} as const;
const BEVEL_UNAVAILABLE = 'Text bevels are not supported by the library yet.';

export const WORDART_PRESETS: readonly WordArtPreset[] = [
  {
    // No fill: the text keeps the color it inherits.
    label: 'Fill: Black, Text color 1; Shadow',
    format: () => ({ outline: NO_OUTLINE, shadow: { ...SOFT, color: 'dk1', opacity: 0.4 } }),
  },
  {
    label: 'Fill: Blue, Accent color 1; Shadow',
    format: () => ({
      color: 'accent1',
      outline: NO_OUTLINE,
      shadow: {
        color: '#6E747A',
        opacity: 0.43,
        blurEmu: 38100,
        offsetEmu: 25400,
        angleDeg: 90,
        alignment: 'ctr',
        rotateWithShape: false,
      },
    }),
  },
  {
    label: 'Fill: Red, Accent color 2; Outline: Red, Accent color 2',
    format: (tone) => ({
      bold: true,
      color: tone('accent2', ...lighter(0.4, 0.6)),
      outline: { color: 'accent2', widthEmu: 22225 },
    }),
  },
  {
    label: 'Fill: White; Outline: Aqua, Accent color 5; Shadow',
    format: () => ({
      bold: true,
      color: '#FFFFFF',
      outline: { color: 'accent5', widthEmu: 10160 },
      shadow: {
        color: '#000000',
        opacity: 0.3,
        blurEmu: 38100,
        offsetEmu: 22860,
        angleDeg: 90,
        alignment: 'tl',
        rotateWithShape: false,
      },
    }),
  },
  {
    label: 'Gradient Fill, Gray',
    format: () => ({
      outline: NO_OUTLINE,
      textFill: {
        kind: 'gradient',
        angleDeg: 90,
        stops: [
          { offset: 0.21, color: '#53575C' },
          { offset: 0.88, color: '#C5C7CA' },
        ],
      },
    }),
  },
  {
    label: 'Fill: Purple, Accent color 4; Soft Bevel',
    format: () => ({ bold: true, color: 'accent4', outline: EMPTY_OUTLINE }),
    unavailable: BEVEL_UNAVAILABLE,
  },
  {
    label: 'Gradient Fill: Aqua, Accent color 5; Reflection',
    format: () => ({
      outline: NO_OUTLINE,
      textFill: {
        kind: 'gradient',
        angleDeg: 90,
        stops: [
          { offset: 0, color: 'accent5', colorTransforms: darker(0.5) },
          { offset: 0.5, color: 'accent5' },
          { offset: 1, color: 'accent5', colorTransforms: lighter(0.6, 0.4) },
        ],
      },
      reflection: {
        blurEmu: 6350,
        startOpacity: 0.53,
        opacity: 0.003,
        endPosition: 0.355,
        angleDeg: 90,
        scaleY: -0.9,
        alignment: 'bl',
        rotateWithShape: false,
      },
    }),
  },
  {
    label: 'Gradient Fill: Purple, Accent color 4; Outline: Purple, Accent color 4',
    format: () => ({
      bold: true,
      outline: { color: 'accent4', widthEmu: 12700 },
      textFill: {
        kind: 'gradient',
        angleDeg: 90,
        stops: [
          { offset: 0, color: 'accent4' },
          { offset: 0.04, color: 'accent4', colorTransforms: lighter(0.6, 0.4) },
          { offset: 0.87, color: 'accent4', colorTransforms: lighter(0.2, 0.8) },
        ],
      },
    }),
  },
  {
    label: 'Fill: White; Outline: Blue, Accent color 1; Glow: Blue, Accent color 1',
    format: (tone) => ({
      bold: true,
      spc: 50,
      // Native writes this literal green at a 1% tint (near white) despite the
      // "White" label; reproduced as captured.
      color: tone('#70AD47', { kind: 'tint', value: 0.01 }),
      outline: { color: 'accent1', widthEmu: 9525 },
      glow: { color: 'accent1', radiusEmu: 38100, opacity: 0.4 },
    }),
  },
  {
    label: 'Fill: Olive Green, Accent color 3; Sharp Bevel',
    format: () => ({ bold: true, color: 'accent3', outline: EMPTY_OUTLINE }),
    unavailable: BEVEL_UNAVAILABLE,
  },
  {
    // No fill, as in the first preset.
    label:
      'Fill: Black, Text color 1; Outline: White, Background color 1; Hard Shadow: White, Background color 1',
    format: (tone) => ({
      bold: true,
      outline: { color: 'bg1', widthEmu: 9525 },
      shadow: { ...HARD, blurEmu: 12700, color: tone('bg1', ...darker(0.5)) },
    }),
  },
  {
    label:
      'Fill: Black, Text color 1; Outline: White, Background color 1; Hard Shadow: Aqua, Accent color 5',
    format: (tone) => ({
      bold: true,
      color: tone('tx1', ...lighter(0.85, 0.15)),
      outline: { color: 'bg1', widthEmu: 13462 },
      shadow: { ...HARD, blurEmu: 0, alignment: 'bl', color: 'accent5' },
    }),
  },
  {
    label:
      'Fill: Aqua, Accent color 5; Outline: White, Background color 1; Hard Shadow: Aqua, Accent color 5',
    format: (tone) => ({
      bold: true,
      color: 'accent5',
      outline: { color: 'bg1', widthEmu: 9525 },
      shadow: { ...HARD, blurEmu: 12700, color: tone('accent5', ...lighter(0.6, 0.4)) },
    }),
  },
  {
    label: 'Fill: White; Outline: Red, Accent color 2; Hard Shadow: Red, Accent color 2',
    format: () => ({
      bold: true,
      color: '#FFFFFF',
      outline: { color: 'accent2', widthEmu: 6600 },
      shadow: { ...HARD, blurEmu: 0, color: 'accent2' },
    }),
  },
  {
    label: 'Fill: Tan, Background color 2; Inner Shadow',
    format: () => ({
      bold: true,
      spc: 50,
      color: 'bg2',
      outline: NO_OUTLINE,
      innerShadow: {
        color: '#000000',
        opacity: 0.5,
        blurEmu: 63500,
        offsetEmu: 50800,
        angleDeg: 225,
      },
    }),
  },
  {
    label: 'Pattern Fill: White; Dark Upward Diagonal Stripe; Shadow',
    format: (tone) => ({
      bold: true,
      outline: EMPTY_OUTLINE,
      textFill: {
        kind: 'pattern',
        preset: 'dkUpDiag',
        foreground: 'bg1',
        foregroundTransforms: darker(0.5),
        background: 'tx1',
        backgroundTransforms: lighter(0.75, 0.25),
      },
      shadow: { ...SOFT, color: tone('dk1', ...darker(0.5)), opacity: 0.4 },
    }),
  },
  {
    label: 'Pattern Fill: Olive Green, Accent color 3, Narrow Horizontal Stripe; Inner Shadow',
    format: (tone) => ({
      bold: true,
      outline: { color: tone('accent3', ...darker(0.5)), widthEmu: 12700 },
      textFill: {
        kind: 'pattern',
        preset: 'narHorz',
        foreground: 'accent3',
        background: 'accent3',
        backgroundTransforms: lighter(0.4, 0.6),
      },
      innerShadow: {
        color: tone('accent3', ...darker(0.5)),
        blurEmu: 177800,
        offsetEmu: 0,
        angleDeg: 0,
      },
    }),
  },
  {
    label: 'Pattern Fill: Blue, Accent color 1, 50%; Hard Shadow: Blue, Accent color 1',
    format: () => ({
      bold: true,
      outline: { color: 'accent1', widthEmu: 12700 },
      textFill: {
        kind: 'pattern',
        preset: 'pct50',
        foreground: 'accent1',
        background: 'accent1',
        backgroundTransforms: lighter(0.2, 0.8),
      },
      shadow: { ...PATTERN_HARD, blurEmu: 0, color: 'accent1' },
    }),
  },
  {
    label:
      'Pattern Fill: Aqua, Accent color 5, Light Downward Diagonal Stripe; Outline: Aqua, Accent color 5',
    format: () => ({
      bold: true,
      outline: { color: 'accent5', widthEmu: 12700 },
      textFill: {
        kind: 'pattern',
        preset: 'ltDnDiag',
        foreground: 'accent5',
        foregroundTransforms: lighter(0.6, 0.4),
        background: 'bg1',
      },
    }),
  },
  {
    label: 'Pattern Fill: Dark Blue, Dark Upward Diagonal Stripe; Hard Shadow',
    format: (tone) => ({
      bold: true,
      outline: { color: tone('tx2', ...darker(0.75)), widthEmu: 12700 },
      textFill: {
        kind: 'pattern',
        preset: 'dkUpDiag',
        foreground: 'tx2',
        background: 'tx2',
        backgroundTransforms: lighter(0.2, 0.8),
      },
      shadow: { ...PATTERN_HARD, blurEmu: 0, color: tone('tx2', ...darker(0.75)) },
    }),
  },
];

/** The preset's run format, with transformed colors resolved through `theme`. */
export function wordArtFormat(preset: WordArtPreset, theme: PresentationTheme | null): TextFormat {
  return preset.format(
    (color, ...transforms) => asColor(resolveColor(color, transforms, theme) ?? color) ?? color,
  );
}

// Everything a preset sets. Native replaces these rather than merging, so a
// white bold preset followed by the black shadow one leaves no white or bold.
const CLEARED: TextFormat = {
  color: null,
  outline: null,
  shadow: null,
  innerShadow: null,
  glow: null,
  reflection: null,
};

/** Applies a gallery preset to every run of `shape` and its paragraph ends. */
export function applyWordArtPreset(
  shape: SlideShapeData,
  preset: WordArtPreset,
  theme: PresentationTheme | null,
): void {
  const format = wordArtFormat(preset, theme);
  const runs = Array.from({ length: getShapeParagraphCount(shape) }, (_, paragraph) =>
    Array.from({ length: getShapeRunCount(shape, paragraph) }, (_, run) =>
      getShapeRunFormat(shape, paragraph, run),
    ),
  ).flat();
  // Native drops `b` and `spc`; the text-format API can only write b="0" and
  // spc="0", so it does so only over text that sets them itself.
  setShapeTextFormat(shape, {
    ...CLEARED,
    ...(!format.bold && runs.some((run) => run?.bold) ? { bold: false } : {}),
    ...(format.spc === undefined && runs.some((run) => run?.spc) ? { spc: 0 } : {}),
  });
  setShapeTextFormat(shape, format);
}
