// Human-authored refinements over the generated capability manifest.
//
// The generator gives every capability a working default (a humanized label and
// a parsed parameter schema). This file upgrades the ones that benefit from a
// hand-tuned schema, a bilingual label, a ribbon home, or `primary` prominence.
// Entries here NEVER remove capabilities — coverage stays exhaustive regardless
// of how much of this file is filled in. Unlisted capabilities simply use their
// generated defaults, which are still fully reachable via the command palette.

import { generatedOverrides } from './overrides.generated.ts';
import type { CapabilityOverride } from './types.ts';

// Hand-authored refinements. Merged on top of `generatedOverrides` (the
// workflow-enriched field schemas), so a hand entry wins for the same id.
const handOverrides: Record<string, CapabilityOverride> = {
  // --- Slides ------------------------------------------------------------
  addBlankSlide: {
    labelEn: 'Blank slide',
    labelJa: '白紙のスライド',
    ribbonGroup: 'slides',
    primary: true,
    params: [],
  },
  addTitleSlide: {
    labelEn: 'Title slide',
    labelJa: 'タイトルスライド',
    ribbonGroup: 'slides',
    primary: true,
    params: [{ name: 'title', type: 'string', kind: 'string', optional: false, label: 'Title' }],
  },
  addContentSlide: {
    labelEn: 'Title + content slide',
    labelJa: 'タイトルとコンテンツ',
    ribbonGroup: 'slides',
    primary: true,
    params: [
      { name: 'title', type: 'string', kind: 'string', optional: true, label: 'Title' },
      { name: 'body', type: 'string', kind: 'string', optional: true, label: 'Body' },
    ],
  },
  duplicateSlide: {
    labelEn: 'Duplicate slide',
    labelJa: 'スライドの複製',
    ribbonGroup: 'slides',
    primary: true,
    params: [],
  },
  removeSlide: {
    labelEn: 'Delete slide',
    labelJa: 'スライドの削除',
    ribbonGroup: 'slides',
    primary: true,
    params: [],
  },
  moveSlide: { labelEn: 'Move slide', labelJa: 'スライドの移動', ribbonGroup: 'slides' },
  setSlideHidden: {
    labelEn: 'Hide slide',
    labelJa: 'スライドを非表示',
    ribbonGroup: 'slides',
    params: [
      { name: 'hidden', type: 'boolean', kind: 'boolean', optional: false, label: 'Hidden' },
    ],
  },
  setSlideSize: { labelEn: 'Slide size', labelJa: 'スライドのサイズ', ribbonGroup: 'design' },

  // --- Shapes ------------------------------------------------------------
  setShapeFill: {
    labelEn: 'Shape fill',
    labelJa: '図形の塗りつぶし',
    ribbonGroup: 'shape-styles',
    primary: true,
    params: [{ name: 'color', type: 'string', kind: 'color', optional: false, label: 'Color' }],
  },
  setShapeStroke: {
    labelEn: 'Shape outline',
    labelJa: '図形の枠線',
    ribbonGroup: 'shape-styles',
    primary: true,
    // The library takes a single `options` object — flattening it would pass
    // the wrong positional args. Render it as a field group.
    params: [
      {
        name: 'options',
        type: '{ color?: string; widthEmu?: number }',
        kind: 'object',
        optional: false,
        fields: [
          { name: 'color', type: 'string', kind: 'color', optional: true, label: 'Color' },
          { name: 'widthEmu', type: 'number', kind: 'emu', optional: true, label: 'Width' },
        ],
      },
    ],
  },
  setShapeNoFill: {
    labelEn: 'No fill',
    labelJa: '塗りつぶしなし',
    ribbonGroup: 'shape-styles',
    params: [],
  },
  setShapeNoStroke: {
    labelEn: 'No outline',
    labelJa: '枠線なし',
    ribbonGroup: 'shape-styles',
    params: [],
  },
  removeShape: {
    labelEn: 'Delete shape',
    labelJa: '図形の削除',
    ribbonGroup: 'arrange',
    primary: true,
    params: [],
  },
  bringShapeToFront: {
    labelEn: 'Bring to front',
    labelJa: '最前面へ',
    ribbonGroup: 'arrange',
    params: [],
  },
  sendShapeToBack: {
    labelEn: 'Send to back',
    labelJa: '最背面へ',
    ribbonGroup: 'arrange',
    params: [],
  },
  bringShapeForward: {
    labelEn: 'Bring forward',
    labelJa: '前面へ',
    ribbonGroup: 'arrange',
    params: [],
  },
  sendShapeBackward: {
    labelEn: 'Send backward',
    labelJa: '背面へ',
    ribbonGroup: 'arrange',
    params: [],
  },
  setShapeRotation: {
    labelEn: 'Rotation',
    labelJa: '回転',
    ribbonGroup: 'arrange',
    params: [
      {
        name: 'degrees',
        type: 'number',
        kind: 'number',
        optional: false,
        label: 'Degrees',
        default: '0',
      },
    ],
  },

  // --- Effects / advanced fills (field-based nested dialogs) --------------
  setShapeGradientFill: {
    labelEn: 'Gradient fill',
    labelJa: 'グラデーション',
    ribbonGroup: 'shape-styles',
    params: [
      {
        name: 'options',
        type: 'GradientFillOptions',
        kind: 'object',
        optional: false,
        fields: [
          {
            name: 'stops',
            type: 'GradientStop[]',
            kind: 'array',
            optional: false,
            label: 'Color stops',
            item: {
              name: 'stop',
              type: 'GradientStop',
              kind: 'object',
              optional: false,
              fields: [
                {
                  name: 'offset',
                  type: 'number',
                  kind: 'number',
                  optional: false,
                  label: 'Offset (0–1)',
                },
                { name: 'color', type: 'string', kind: 'color', optional: false, label: 'Color' },
              ],
            },
          },
          {
            name: 'angleDeg',
            type: 'number',
            kind: 'number',
            optional: true,
            label: 'Angle (°)',
            default: '90',
          },
          {
            name: 'path',
            type: "'linear'|'circle'|'rect'|'shape'",
            kind: 'enum',
            optional: true,
            label: 'Path',
            enumValues: ['linear', 'circle', 'rect', 'shape'],
          },
        ],
      },
    ],
  },
  setShapePatternFill: {
    labelEn: 'Pattern fill',
    labelJa: 'パターン',
    ribbonGroup: 'shape-styles',
    params: [
      {
        name: 'options',
        type: 'PatternFillOptions',
        kind: 'object',
        optional: false,
        fields: [
          {
            name: 'preset',
            type: 'PatternPreset',
            kind: 'string',
            optional: false,
            label: 'Preset (e.g. pct50, dkUpDiag, wave)',
          },
          {
            name: 'foreground',
            type: 'string',
            kind: 'color',
            optional: false,
            label: 'Foreground',
          },
          {
            name: 'background',
            type: 'string',
            kind: 'color',
            optional: false,
            label: 'Background',
          },
        ],
      },
    ],
  },
  setShapeShadow: {
    labelEn: 'Shadow',
    labelJa: '影',
    ribbonGroup: 'effects',
    params: [
      {
        name: 'options',
        type: 'ShadowOptions',
        kind: 'object',
        optional: true,
        fields: [
          { name: 'color', type: 'string', kind: 'color', optional: true, label: 'Color' },
          { name: 'blurEmu', type: 'Emu', kind: 'emu', optional: true, label: 'Blur' },
          { name: 'offsetEmu', type: 'Emu', kind: 'emu', optional: true, label: 'Offset' },
          {
            name: 'angleDeg',
            type: 'number',
            kind: 'number',
            optional: true,
            label: 'Angle (°)',
            default: '45',
          },
          {
            name: 'opacity',
            type: 'number',
            kind: 'number',
            optional: true,
            label: 'Opacity (0–1)',
          },
        ],
      },
    ],
  },
  setShapeGlow: {
    labelEn: 'Glow',
    labelJa: '光彩',
    ribbonGroup: 'effects',
    params: [
      {
        name: 'options',
        type: 'GlowOptions',
        kind: 'object',
        optional: false,
        fields: [
          { name: 'color', type: 'string', kind: 'color', optional: false, label: 'Color' },
          { name: 'radiusEmu', type: 'Emu', kind: 'emu', optional: true, label: 'Radius' },
        ],
      },
    ],
  },
  setSlideTransition: {
    labelEn: 'Transition',
    labelJa: '画面切り替え',
    ribbonGroup: 'transition',
    primary: true,
    params: [
      {
        name: 'options',
        type: 'TransitionOptions',
        kind: 'object',
        optional: false,
        fields: [
          {
            name: 'effect',
            type: 'TransitionEffect',
            kind: 'enum',
            optional: false,
            label: 'Effect',
            enumValues: [
              'none',
              'fade',
              'push',
              'cover',
              'wipe',
              'split',
              'cut',
              'dissolve',
              'checker',
              'blinds',
              'randomBar',
              'zoom',
              'circle',
              'diamond',
              'plus',
              'wedge',
              'newsflash',
            ],
          },
          {
            name: 'speed',
            type: "'slow'|'med'|'fast'",
            kind: 'enum',
            optional: true,
            label: 'Speed',
            enumValues: ['slow', 'med', 'fast'],
          },
          {
            name: 'direction',
            type: 'string',
            kind: 'string',
            optional: true,
            label: 'Direction (effect-specific, e.g. l/r/u/d)',
          },
          {
            name: 'thruBlack',
            type: 'boolean',
            kind: 'boolean',
            optional: true,
            label: 'Through black (fade)',
          },
        ],
      },
    ],
  },
  setShapeAnimation: {
    labelEn: 'Animation',
    labelJa: 'アニメーション',
    ribbonGroup: 'animation',
    primary: true,
    params: [
      {
        name: 'opts',
        type: 'AnimationOptions',
        kind: 'object',
        optional: false,
        fields: [
          {
            name: 'effect',
            type: 'AnimationEffect',
            kind: 'enum',
            optional: false,
            label: 'Effect',
            enumValues: ['fadeIn', 'fadeOut', 'appear', 'disappear'],
          },
          {
            name: 'durationMs',
            type: 'number',
            kind: 'number',
            optional: true,
            label: 'Duration (ms)',
            default: '500',
          },
        ],
      },
    ],
  },
  setShapeImageCrop: {
    labelEn: 'Crop image',
    labelJa: '画像のトリミング',
    ribbonGroup: 'picture',
    params: [
      {
        name: 'crop',
        type: 'ImageCrop',
        kind: 'object',
        optional: false,
        fields: [
          { name: 'left', type: 'number', kind: 'number', optional: true, label: 'Left (0–1)' },
          { name: 'top', type: 'number', kind: 'number', optional: true, label: 'Top (0–1)' },
          { name: 'right', type: 'number', kind: 'number', optional: true, label: 'Right (0–1)' },
          { name: 'bottom', type: 'number', kind: 'number', optional: true, label: 'Bottom (0–1)' },
        ],
      },
    ],
  },
};

// The effective override map: workflow-enriched field schemas as the base,
// hand-authored refinements layered on top (hand wins per id).
export const overrides: Record<string, CapabilityOverride> = {
  ...generatedOverrides,
  ...handOverrides,
};
