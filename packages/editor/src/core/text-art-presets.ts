// The reference desktop app's (Mac) Text Art gallery: its twenty presets in gallery order, each
// the character payload the reference desktop app wrote when the preset was applied
// (test/fixtures/native/text-art-*-shape.xml, catalogued in text-art-capture.md).

import {
  setShapeText3D,
  setShapeTextFormat,
  setTableCellTextFormat,
  type ColorTransform,
  type SlideShapeData,
  type TableCellData,
  type Text3D,
  type TextFormat,
} from '@office-kit/pptx';

export interface TextArtPreset {
  /** The reference desktop app's gallery label, also the swatch's tooltip and accessible name. */
  readonly label: string;
  /** The run properties, written to every run and paragraph end. */
  readonly format: TextFormat;
  /** The bevel presets' text-body 3-D (`<a:scene3d>` / `<a:sp3d>` in `<a:bodyPr>`). */
  readonly text3D?: Text3D;
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
// Both bevels extrude the glyphs 4.5pt.
const BEVEL_EXTRUSION_EMU = 57150;

export const TEXT_ART_PRESETS: readonly TextArtPreset[] = [
  {
    // No fill: the text keeps the color it inherits.
    label: 'Fill: Black, Text color 1; Shadow',
    format: { outline: NO_OUTLINE, shadow: { ...SOFT, color: 'dk1', opacity: 0.4 } },
  },
  {
    label: 'Fill: Blue, Accent color 1; Shadow',
    format: {
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
    },
  },
  {
    label: 'Fill: Red, Accent color 2; Outline: Red, Accent color 2',
    format: {
      bold: true,
      color: 'accent2',
      colorTransforms: lighter(0.4, 0.6),
      outline: { color: 'accent2', widthEmu: 22225 },
    },
  },
  {
    label: 'Fill: White; Outline: Aqua, Accent color 5; Shadow',
    format: {
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
    },
  },
  {
    label: 'Gradient Fill, Gray',
    format: {
      outline: NO_OUTLINE,
      textFill: {
        kind: 'gradient',
        angleDeg: 90,
        stops: [
          { offset: 0.21, color: '#53575C' },
          { offset: 0.88, color: '#C5C7CA' },
        ],
      },
    },
  },
  {
    label: 'Fill: Purple, Accent color 4; Soft Bevel',
    format: { bold: true, color: 'accent4', outline: EMPTY_OUTLINE },
    text3D: {
      scene: {
        camera: 'orthographicFront',
        lightRig: {
          type: 'soft',
          direction: 't',
          rotation: { latitudeDeg: 0, longitudeDeg: 0, revolutionDeg: 260 },
        },
      },
      bevelTop: { widthEmu: 25400, heightEmu: 38100 },
      extrusionHeightEmu: BEVEL_EXTRUSION_EMU,
      material: 'softEdge',
    },
  },
  {
    label: 'Gradient Fill: Aqua, Accent color 5; Reflection',
    format: {
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
    },
  },
  {
    label: 'Gradient Fill: Purple, Accent color 4; Outline: Purple, Accent color 4',
    format: {
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
    },
  },
  {
    label: 'Fill: White; Outline: Blue, Accent color 1; Glow: Blue, Accent color 1',
    format: {
      bold: true,
      spc: 50,
      // Native writes this literal green at a 1% tint (near white) despite the
      // "White" label; reproduced as captured.
      color: '#70AD47',
      colorTransforms: [{ kind: 'tint', value: 0.01 }],
      outline: { color: 'accent1', widthEmu: 9525 },
      glow: { color: 'accent1', radiusEmu: 38100, opacity: 0.4 },
    },
  },
  {
    label: 'Fill: Olive Green, Accent color 3; Sharp Bevel',
    format: { bold: true, color: 'accent3', outline: EMPTY_OUTLINE },
    text3D: {
      scene: { camera: 'orthographicFront', lightRig: { type: 'harsh', direction: 't' } },
      bevelTop: { widthEmu: 63500, heightEmu: 12700, preset: 'angle' },
      extrusionHeightEmu: BEVEL_EXTRUSION_EMU,
      material: 'matte',
      // Native writes a contour color with no contour width; reproduced as captured.
      contourColor: 'bg1',
      contourColorTransforms: darker(0.65),
    },
  },
  {
    // No fill, as in the first preset.
    label:
      'Fill: Black, Text color 1; Outline: White, Background color 1; Hard Shadow: White, Background color 1',
    format: {
      bold: true,
      outline: { color: 'bg1', widthEmu: 9525 },
      shadow: { ...HARD, blurEmu: 12700, color: 'bg1', colorTransforms: darker(0.5) },
    },
  },
  {
    label:
      'Fill: Black, Text color 1; Outline: White, Background color 1; Hard Shadow: Aqua, Accent color 5',
    format: {
      bold: true,
      color: 'tx1',
      colorTransforms: lighter(0.85, 0.15),
      outline: { color: 'bg1', widthEmu: 13462 },
      shadow: { ...HARD, blurEmu: 0, alignment: 'bl', color: 'accent5' },
    },
  },
  {
    label:
      'Fill: Aqua, Accent color 5; Outline: White, Background color 1; Hard Shadow: Aqua, Accent color 5',
    format: {
      bold: true,
      color: 'accent5',
      outline: { color: 'bg1', widthEmu: 9525 },
      shadow: { ...HARD, blurEmu: 12700, color: 'accent5', colorTransforms: lighter(0.6, 0.4) },
    },
  },
  {
    label: 'Fill: White; Outline: Red, Accent color 2; Hard Shadow: Red, Accent color 2',
    format: {
      bold: true,
      color: '#FFFFFF',
      outline: { color: 'accent2', widthEmu: 6600 },
      shadow: { ...HARD, blurEmu: 0, color: 'accent2' },
    },
  },
  {
    label: 'Fill: Tan, Background color 2; Inner Shadow',
    format: {
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
    },
  },
  {
    label: 'Pattern Fill: White; Dark Upward Diagonal Stripe; Shadow',
    format: {
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
      shadow: { ...SOFT, color: 'dk1', colorTransforms: darker(0.5), opacity: 0.4 },
    },
  },
  {
    label: 'Pattern Fill: Olive Green, Accent color 3, Narrow Horizontal Stripe; Inner Shadow',
    format: {
      bold: true,
      outline: { color: 'accent3', colorTransforms: darker(0.5), widthEmu: 12700 },
      textFill: {
        kind: 'pattern',
        preset: 'narHorz',
        foreground: 'accent3',
        background: 'accent3',
        backgroundTransforms: lighter(0.4, 0.6),
      },
      innerShadow: {
        color: 'accent3',
        colorTransforms: darker(0.5),
        blurEmu: 177800,
        offsetEmu: 0,
        angleDeg: 0,
      },
    },
  },
  {
    label: 'Pattern Fill: Blue, Accent color 1, 50%; Hard Shadow: Blue, Accent color 1',
    format: {
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
    },
  },
  {
    label:
      'Pattern Fill: Aqua, Accent color 5, Light Downward Diagonal Stripe; Outline: Aqua, Accent color 5',
    format: {
      bold: true,
      outline: { color: 'accent5', widthEmu: 12700 },
      textFill: {
        kind: 'pattern',
        preset: 'ltDnDiag',
        foreground: 'accent5',
        foregroundTransforms: lighter(0.6, 0.4),
        background: 'bg1',
      },
    },
  },
  {
    label: 'Pattern Fill: Dark Blue, Dark Upward Diagonal Stripe; Hard Shadow',
    format: {
      bold: true,
      outline: { color: 'tx2', colorTransforms: darker(0.75), widthEmu: 12700 },
      textFill: {
        kind: 'pattern',
        preset: 'dkUpDiag',
        foreground: 'tx2',
        background: 'tx2',
        backgroundTransforms: lighter(0.2, 0.8),
      },
      shadow: { ...PATTERN_HARD, blurEmu: 0, color: 'tx2', colorTransforms: darker(0.75) },
    },
  },
];

// Everything a preset sets. Native replaces these rather than merging, so a
// white bold preset followed by the black shadow one leaves no white or bold:
// The reference desktop app removes `b` and `spc` rather than writing them off.
const CLEARED: TextFormat = {
  color: null,
  bold: null,
  spc: null,
  outline: null,
  shadow: null,
  innerShadow: null,
  glow: null,
  reflection: null,
};

/** Applies a gallery preset to every run of `shape` and its paragraph ends. */
export function applyTextArtPreset(shape: SlideShapeData, preset: TextArtPreset): void {
  setShapeTextFormat(shape, CLEARED);
  setShapeTextFormat(shape, preset.format);
  // The bevel belongs to the preset like the run effects do, so a preset
  // without one removes it. (Unlike the run properties, this is not yet
  // confirmed against a native capture of one preset applied over another.)
  setShapeText3D(shape, preset.text3D ?? null);
}

/**
 * Table Design ▸ Text Art Styles ▸ Quick Styles over table cells: the same run
 * properties as on a shape. Cells have no text bevel, so a preset's bevel is
 * left out.
 */
export function applyTableCellTextArtPreset(cell: TableCellData, preset: TextArtPreset): void {
  setTableCellTextFormat(cell, CLEARED);
  setTableCellTextFormat(cell, preset.format);
}
