// The Format pane's Effects sections (Shadow, Reflection, Glow, Soft Edges,
// 3-D Format, 3-D Rotation) edit either a shape's own effects or the effects
// on its text runs. Both go through one `EffectTarget` so the sections are
// written once. Preset values follow what Mac PowerPoint writes for its
// galleries; the ones not compared against a native capture are listed in
// POWERPOINT_PARITY.md.

import {
  asColor,
  getShape3D,
  getShapeEffects,
  getShapeKind,
  getShapeText,
  getShapeText3D,
  setShape3D,
  setShapeGlow,
  setShapeInnerShadow,
  setShapeReflection,
  setShapeShadow,
  setShapeSoftEdge,
  setShapeText3D,
  setShapeTextFormat,
  type BevelPreset,
  type CameraPreset,
  type Color,
  type ColorTransform,
  type LightRigType,
  type PresentationData,
  type PresetMaterial,
  type ReadText3D,
  type ReflectionOptions,
  type SlideShapeData,
  type Text3D,
} from '@office-kit/pptx';
import { textFormatsInRange } from '../core/text-format-selection.ts';

export const EMU_PER_POINT = 12700;

type Alignment = 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';

/** One shadow, outer or inner: PowerPoint's Shadow section holds exactly one. */
export interface ShadowValue {
  readonly kind: 'outer' | 'inner';
  readonly color: Color;
  readonly colorTransforms?: readonly ColorTransform[];
  readonly opacity: number;
  readonly blurEmu: number;
  readonly offsetEmu: number;
  readonly angleDeg: number;
  readonly scaleX?: number;
  readonly scaleY?: number;
  readonly skewX?: number;
  readonly skewY?: number;
  readonly alignment?: Alignment;
}

export interface GlowValue {
  readonly color: Color;
  readonly colorTransforms?: readonly ColorTransform[];
  readonly opacity: number;
  readonly radiusEmu: number;
}

export interface EffectsState {
  readonly shadow: ShadowValue | null;
  readonly reflection: ReflectionOptions | null;
  readonly glow: GlowValue | null;
  readonly softEdgeEmu: number | null;
  readonly threeD: ReadText3D | null;
}

export interface EffectTarget {
  readonly kind: 'shape' | 'text';
  read(pres: PresentationData, shape: SlideShapeData): EffectsState;
  shadow(shape: SlideShapeData, value: ShadowValue | null): void;
  reflection(shape: SlideShapeData, value: ReflectionOptions | null): void;
  glow(shape: SlideShapeData, value: GlowValue | null): void;
  /** Absent where the library cannot write a soft edge (text runs). */
  softEdge?: (shape: SlideShapeData, radiusEmu: number | null) => void;
  threeD(shape: SlideShapeData, value: Text3D | null): void;
}

const color = (value: string | undefined, fallback: Color): Color =>
  (value === undefined ? undefined : asColor(value)) ?? fallback;

export const shapeEffects: EffectTarget = {
  kind: 'shape',
  read(pres, shape) {
    let shadow: ShadowValue | null = null;
    let reflection: ReflectionOptions | null = null;
    let glow: GlowValue | null = null;
    let softEdgeEmu: number | null = null;
    for (const effect of getShapeEffects(pres, shape)) {
      if (effect.kind === 'outerShdw' || (effect.kind === 'innerShdw' && !shadow)) {
        shadow = {
          kind: effect.kind === 'outerShdw' ? 'outer' : 'inner',
          color: color(effect.color, '#000000'),
          opacity: effect.opacity ?? 1,
          blurEmu: effect.blurEmu,
          offsetEmu: effect.distEmu,
          angleDeg: effect.angleDeg,
          ...(effect.kind === 'outerShdw'
            ? {
                ...(effect.scaleX === undefined ? {} : { scaleX: effect.scaleX }),
                ...(effect.scaleY === undefined ? {} : { scaleY: effect.scaleY }),
                ...(effect.skewX === undefined ? {} : { skewX: effect.skewX }),
                ...(effect.skewY === undefined ? {} : { skewY: effect.skewY }),
                ...(effect.alignment === undefined ? {} : { alignment: effect.alignment }),
              }
            : {}),
        };
      } else if (effect.kind === 'glow') {
        glow = {
          color: color(effect.color, 'accent1'),
          opacity: effect.opacity ?? 1,
          radiusEmu: effect.radiusEmu,
        };
      } else if (effect.kind === 'reflection') {
        const { kind: _kind, distEmu, ...rest } = effect;
        reflection = { ...rest, offsetEmu: distEmu };
      } else if (effect.kind === 'softEdge') {
        softEdgeEmu = effect.radiusEmu;
      }
    }
    return { shadow, reflection, glow, softEdgeEmu, threeD: getShape3D(shape) };
  },
  shadow(shape, value) {
    if (value?.kind === 'inner') {
      setShapeShadow(shape, null);
      setShapeInnerShadow(shape, innerOptions(value));
    } else {
      setShapeInnerShadow(shape, null);
      setShapeShadow(shape, value && outerOptions(value));
    }
  },
  reflection: setShapeReflection,
  glow: setShapeGlow,
  softEdge: setShapeSoftEdge,
  threeD: setShape3D,
};

const outerOptions = (value: ShadowValue) => ({
  color: value.color,
  ...(value.colorTransforms ? { colorTransforms: value.colorTransforms } : {}),
  opacity: value.opacity,
  blurEmu: value.blurEmu,
  offsetEmu: value.offsetEmu,
  angleDeg: value.angleDeg,
  alignment: value.alignment ?? 'tl',
  ...(value.scaleX === undefined ? {} : { scaleX: value.scaleX }),
  ...(value.scaleY === undefined ? {} : { scaleY: value.scaleY }),
  ...(value.skewX === undefined ? {} : { skewX: value.skewX }),
  ...(value.skewY === undefined ? {} : { skewY: value.skewY }),
});

const innerOptions = (value: ShadowValue) => ({
  color: value.color,
  ...(value.colorTransforms ? { colorTransforms: value.colorTransforms } : {}),
  opacity: value.opacity,
  blurEmu: value.blurEmu,
  offsetEmu: value.offsetEmu,
  angleDeg: value.angleDeg,
});

export const textEffects: EffectTarget = {
  kind: 'text',
  read(pres, shape) {
    const format =
      getShapeKind(shape) === 'shape'
        ? textFormatsInRange(shape, { start: 0, end: getShapeText(shape).length }, undefined, {
            pres,
          })[0]
        : undefined;
    const outer = format?.shadow;
    const inner = format?.innerShadow;
    const shadow: ShadowValue | null = outer
      ? {
          kind: 'outer',
          color: outer.color ?? '#000000',
          opacity: outer.opacity ?? 1,
          blurEmu: outer.blurEmu ?? 0,
          offsetEmu: outer.offsetEmu ?? 0,
          angleDeg: outer.angleDeg ?? 0,
          ...(outer.scaleX === undefined ? {} : { scaleX: outer.scaleX }),
          ...(outer.scaleY === undefined ? {} : { scaleY: outer.scaleY }),
          ...(outer.skewX === undefined ? {} : { skewX: outer.skewX }),
          ...(outer.skewY === undefined ? {} : { skewY: outer.skewY }),
          ...(outer.alignment === undefined ? {} : { alignment: outer.alignment }),
        }
      : inner
        ? {
            kind: 'inner',
            color: inner.color ?? '#000000',
            opacity: inner.opacity ?? 1,
            blurEmu: inner.blurEmu ?? 0,
            offsetEmu: inner.offsetEmu ?? 0,
            angleDeg: inner.angleDeg ?? 0,
          }
        : null;
    const glow = format?.glow
      ? {
          color: format.glow.color,
          opacity: format.glow.opacity ?? 1,
          radiusEmu: format.glow.radiusEmu ?? 0,
        }
      : null;
    return {
      shadow,
      reflection: format?.reflection ?? null,
      glow,
      softEdgeEmu: null,
      threeD: getShapeKind(shape) === 'shape' ? getShapeText3D(shape) : null,
    };
  },
  shadow(shape, value) {
    setShapeTextFormat(
      shape,
      value?.kind === 'inner'
        ? { shadow: null, innerShadow: innerOptions(value) }
        : { innerShadow: null, shadow: value && outerOptions(value) },
    );
  },
  reflection(shape, value) {
    setShapeTextFormat(shape, { reflection: value });
  },
  glow(shape, value) {
    setShapeTextFormat(shape, { glow: value });
  },
  threeD: setShapeText3D,
};

// ---------------------------------------------------------------------------
// Presets.

export interface Preset<T> {
  readonly label: string;
  readonly ja: string;
  readonly value: T;
}

const black = { color: '#000000' as Color };
const outer = (
  label: string,
  ja: string,
  angleDeg: number,
  alignment: Alignment,
): Preset<ShadowValue> => ({
  label,
  ja,
  value: {
    kind: 'outer',
    ...black,
    opacity: 0.4,
    blurEmu: 50800,
    offsetEmu: 38100,
    angleDeg,
    alignment,
  },
});
const inner = (label: string, ja: string, angleDeg: number): Preset<ShadowValue> => ({
  label,
  ja,
  value: { kind: 'inner', ...black, opacity: 0.5, blurEmu: 63500, offsetEmu: 50800, angleDeg },
});

export const SHADOW_PRESETS: readonly {
  readonly heading: string;
  readonly ja: string;
  readonly items: readonly Preset<ShadowValue>[];
}[] = [
  {
    heading: 'Outer',
    ja: '外側',
    items: [
      outer('Offset: Bottom Right', 'オフセット: 右下', 45, 'tl'),
      outer('Offset: Bottom', 'オフセット: 下', 90, 't'),
      outer('Offset: Bottom Left', 'オフセット: 左下', 135, 'tr'),
      outer('Offset: Right', 'オフセット: 右', 0, 'l'),
      {
        label: 'Offset: Center',
        ja: 'オフセット: 中央',
        value: {
          kind: 'outer',
          ...black,
          opacity: 0.4,
          blurEmu: 63500,
          offsetEmu: 0,
          angleDeg: 0,
          scaleX: 1.02,
          scaleY: 1.02,
          alignment: 'ctr',
        },
      },
      outer('Offset: Left', 'オフセット: 左', 180, 'r'),
      outer('Offset: Top Right', 'オフセット: 右上', 315, 'bl'),
      outer('Offset: Top', 'オフセット: 上', 270, 'b'),
      outer('Offset: Top Left', 'オフセット: 左上', 225, 'br'),
    ],
  },
  {
    heading: 'Inner',
    ja: '内側',
    items: [
      inner('Inside: Top Left', '内側: 左上', 225),
      inner('Inside: Top', '内側: 上', 270),
      inner('Inside: Top Right', '内側: 右上', 315),
      inner('Inside: Left', '内側: 左', 180),
      {
        label: 'Inside: Center',
        ja: '内側: 中央',
        value: {
          kind: 'inner',
          ...black,
          opacity: 0.5,
          blurEmu: 114300,
          offsetEmu: 0,
          angleDeg: 0,
        },
      },
      inner('Inside: Right', '内側: 右', 0),
      inner('Inside: Bottom Left', '内側: 左下', 135),
      inner('Inside: Bottom', '内側: 下', 90),
      inner('Inside: Bottom Right', '内側: 右下', 45),
    ],
  },
  {
    heading: 'Perspective',
    ja: '透視投影',
    items: [
      {
        label: 'Perspective: Upper Left',
        ja: '透視投影: 左上',
        value: {
          kind: 'outer',
          ...black,
          opacity: 0.2,
          blurEmu: 76200,
          offsetEmu: 0,
          angleDeg: 225,
          scaleY: 0.23,
          skewX: 20,
          alignment: 'br',
        },
      },
      {
        label: 'Perspective: Upper Right',
        ja: '透視投影: 右上',
        value: {
          kind: 'outer',
          ...black,
          opacity: 0.2,
          blurEmu: 76200,
          offsetEmu: 0,
          angleDeg: 315,
          scaleY: 0.23,
          skewX: -20,
          alignment: 'bl',
        },
      },
      {
        label: 'Perspective: Below',
        ja: '透視投影: 下',
        value: {
          kind: 'outer',
          ...black,
          opacity: 0.15,
          blurEmu: 152400,
          offsetEmu: 317500,
          angleDeg: 90,
          scaleX: 0.9,
          scaleY: -0.19,
          alignment: 'ctr',
        },
      },
      {
        label: 'Perspective: Lower Left',
        ja: '透視投影: 左下',
        value: {
          kind: 'outer',
          ...black,
          opacity: 0.2,
          blurEmu: 76200,
          offsetEmu: 12700,
          angleDeg: 135,
          scaleY: -0.23,
          skewX: 13.34,
          alignment: 'bl',
        },
      },
      {
        label: 'Perspective: Lower Right',
        ja: '透視投影: 右下',
        value: {
          kind: 'outer',
          ...black,
          opacity: 0.2,
          blurEmu: 76200,
          offsetEmu: 12700,
          angleDeg: 45,
          scaleY: -0.23,
          skewX: -13.34,
          alignment: 'br',
        },
      },
    ],
  },
];

const reflectionPreset = (
  size: 'Tight' | 'Half' | 'Full',
  offsetPt: 0 | 4 | 8,
): Preset<ReflectionOptions> => {
  const ja = { Tight: '反射 (弱)', Half: '反射 (中)', Full: '反射 (強)' }[size];
  return {
    label: `${size} Reflection: ${offsetPt ? `${offsetPt} point offset` : 'Touching'}`,
    ja: `${ja}: ${offsetPt ? `${offsetPt} pt オフセット` : '接触'}`,
    value: {
      blurEmu: 6350,
      offsetEmu: offsetPt * EMU_PER_POINT,
      angleDeg: 90,
      startOpacity: size === 'Tight' ? 0.52 : 0.5,
      opacity: 0.003,
      endPosition: { Tight: 0.35, Half: 0.55, Full: 0.9 }[size],
      scaleY: -1,
      alignment: 'bl',
      rotateWithShape: false,
    },
  };
};
export const REFLECTION_PRESETS: readonly Preset<ReflectionOptions>[] = (
  [0, 4, 8] as const
).flatMap((offset) =>
  (['Tight', 'Half', 'Full'] as const).map((size) => reflectionPreset(size, offset)),
);

export const GLOW_PRESETS: readonly Preset<GlowValue>[] = [5, 8, 11, 18].flatMap((points) =>
  [1, 2, 3, 4, 5, 6].map((accent) => ({
    label: `Glow: ${points} point; Accent color ${accent}`,
    ja: `光彩: ${points} pt; アクセント カラー ${accent}`,
    value: {
      color: `accent${accent}` as Color,
      colorTransforms: [{ kind: 'satMod', value: 1.75 }],
      opacity: 0.4,
      radiusEmu: points * EMU_PER_POINT,
    },
  })),
);

export const SOFT_EDGE_PRESETS: readonly Preset<number>[] = [1, 2.5, 5, 10, 25, 50].map(
  (points) => ({
    label: `${points} Point`,
    ja: `${points} pt`,
    value: points * EMU_PER_POINT,
  }),
);

export const BEVEL_PRESETS: readonly Preset<BevelPreset>[] = [
  { label: 'Circle', ja: '丸', value: 'circle' },
  { label: 'Relaxed Inset', ja: '額縁風', value: 'relaxedInset' },
  { label: 'Cross', ja: 'クロス', value: 'cross' },
  { label: 'Cool Slant', ja: '斜面', value: 'coolSlant' },
  { label: 'Angle', ja: '角度', value: 'angle' },
  { label: 'Soft Round', ja: 'ソフト ラウンド', value: 'softRound' },
  { label: 'Convex', ja: '凸レンズ', value: 'convex' },
  { label: 'Slope', ja: 'スロープ', value: 'slope' },
  { label: 'Divot', ja: '切り込み', value: 'divot' },
  { label: 'Riblet', ja: 'リブ', value: 'riblet' },
  { label: 'Hard Edge', ja: 'ハード エッジ', value: 'hardEdge' },
  { label: 'Art Deco', ja: 'アール デコ', value: 'artDeco' },
];

export const MATERIAL_PRESETS: readonly {
  readonly heading: string;
  readonly ja: string;
  readonly items: readonly Preset<PresetMaterial>[];
}[] = [
  {
    heading: 'Standard',
    ja: '標準',
    items: [
      { label: 'Matte', ja: 'マット', value: 'matte' },
      { label: 'Warm Matte', ja: 'マット (暖色)', value: 'warmMatte' },
      { label: 'Plastic', ja: 'プラスチック', value: 'plastic' },
      { label: 'Metal', ja: 'メタル', value: 'metal' },
    ],
  },
  {
    heading: 'Special Effect',
    ja: '特殊効果',
    items: [
      { label: 'Dark Edge', ja: '暗いエッジ', value: 'dkEdge' },
      { label: 'Soft Edge', ja: 'ソフト エッジ', value: 'softEdge' },
      { label: 'Flat', ja: 'フラット', value: 'flat' },
      { label: 'Wireframe', ja: 'ワイヤーフレーム', value: 'legacyWireframe' },
    ],
  },
  {
    heading: 'Translucent',
    ja: '半透明',
    items: [
      { label: 'Powder', ja: 'パウダー', value: 'powder' },
      { label: 'Translucent Powder', ja: '半透明パウダー', value: 'translucentPowder' },
      { label: 'Clear', ja: 'クリア', value: 'clear' },
    ],
  },
];

export const LIGHTING_PRESETS: readonly {
  readonly heading: string;
  readonly ja: string;
  readonly items: readonly Preset<LightRigType>[];
}[] = [
  {
    heading: 'Neutral',
    ja: '標準',
    items: [
      { label: 'Three Point', ja: '3 点', value: 'threePt' },
      { label: 'Balance', ja: 'バランス', value: 'balanced' },
      { label: 'Soft', ja: 'ソフト', value: 'soft' },
      { label: 'Harsh', ja: 'ハード', value: 'harsh' },
      { label: 'Flood', ja: 'フラッド', value: 'flood' },
      { label: 'Contrasting', ja: '対照', value: 'contrasting' },
    ],
  },
  {
    heading: 'Warm',
    ja: '暖色',
    items: [
      { label: 'Morning', ja: '朝', value: 'morning' },
      { label: 'Sunrise', ja: '日の出', value: 'sunrise' },
      { label: 'Sunset', ja: '夕焼け', value: 'sunset' },
    ],
  },
  {
    heading: 'Cool',
    ja: '寒色',
    items: [
      { label: 'Chilly', ja: '冷光', value: 'chilly' },
      { label: 'Freezing', ja: '氷点', value: 'freezing' },
    ],
  },
  {
    heading: 'Special',
    ja: '特殊効果',
    items: [
      { label: 'Flat', ja: 'フラット', value: 'flat' },
      { label: 'Two Point', ja: '2 点', value: 'twoPt' },
      { label: 'Glow', ja: '光彩', value: 'glow' },
      { label: 'Bright Room', ja: '明るい部屋', value: 'brightRoom' },
    ],
  },
];

export const ROTATION_PRESETS: readonly {
  readonly heading: string;
  readonly ja: string;
  readonly items: readonly Preset<CameraPreset>[];
}[] = [
  {
    heading: 'Parallel',
    ja: '平行投影',
    items: [
      { label: 'Isometric Left Down', ja: '等角投影: 左下', value: 'isometricLeftDown' },
      { label: 'Isometric Right Up', ja: '等角投影: 右上', value: 'isometricRightUp' },
      { label: 'Isometric Top Up', ja: '等角投影: 上', value: 'isometricTopUp' },
      { label: 'Isometric Bottom Down', ja: '等角投影: 下', value: 'isometricBottomDown' },
      { label: 'Off Axis 1 Left', ja: '不等角投影 1: 左', value: 'isometricOffAxis1Left' },
      { label: 'Off Axis 1 Right', ja: '不等角投影 1: 右', value: 'isometricOffAxis1Right' },
      { label: 'Off Axis 1 Top', ja: '不等角投影 1: 上', value: 'isometricOffAxis1Top' },
      { label: 'Off Axis 2 Left', ja: '不等角投影 2: 左', value: 'isometricOffAxis2Left' },
      { label: 'Off Axis 2 Right', ja: '不等角投影 2: 右', value: 'isometricOffAxis2Right' },
      { label: 'Off Axis 2 Top', ja: '不等角投影 2: 上', value: 'isometricOffAxis2Top' },
    ],
  },
  {
    heading: 'Perspective',
    ja: '透視投影',
    items: [
      { label: 'Perspective Front', ja: '透視投影: 正面', value: 'perspectiveFront' },
      { label: 'Perspective Left', ja: '透視投影: 左', value: 'perspectiveLeft' },
      { label: 'Perspective Right', ja: '透視投影: 右', value: 'perspectiveRight' },
      { label: 'Perspective Below', ja: '透視投影: 下', value: 'perspectiveBelow' },
      { label: 'Perspective Above', ja: '透視投影: 上', value: 'perspectiveAbove' },
      {
        label: 'Perspective Relaxed Moderately',
        ja: '透視投影: 緩やか (中)',
        value: 'perspectiveRelaxedModerately',
      },
      { label: 'Perspective Relaxed', ja: '透視投影: 緩やか', value: 'perspectiveRelaxed' },
      {
        label: 'Perspective Contrasting Left',
        ja: '透視投影: 対照 (左)',
        value: 'perspectiveContrastingLeftFacing',
      },
      {
        label: 'Perspective Contrasting Right',
        ja: '透視投影: 対照 (右)',
        value: 'perspectiveContrastingRightFacing',
      },
      {
        label: 'Perspective Heroic Extreme Left',
        ja: '透視投影: 強調 (左)',
        value: 'perspectiveHeroicExtremeLeftFacing',
      },
      {
        label: 'Perspective Heroic Extreme Right',
        ja: '透視投影: 強調 (右)',
        value: 'perspectiveHeroicExtremeRightFacing',
      },
    ],
  },
  {
    heading: 'Oblique',
    ja: '斜投影',
    items: [
      { label: 'Oblique Top Left', ja: '斜投影: 左上', value: 'obliqueTopLeft' },
      { label: 'Oblique Top Right', ja: '斜投影: 右上', value: 'obliqueTopRight' },
      { label: 'Oblique Bottom Left', ja: '斜投影: 左下', value: 'obliqueBottomLeft' },
      { label: 'Oblique Bottom Right', ja: '斜投影: 右下', value: 'obliqueBottomRight' },
    ],
  },
];

/** What PowerPoint writes when a 3-D field is edited on a shape without 3-D. */
export const DEFAULT_SCENE: NonNullable<Text3D['scene']> = {
  camera: 'orthographicFront',
  lightRig: { type: 'threePt', direction: 't' },
};

export const DEFAULT_BEVEL_EMU = 6 * EMU_PER_POINT;

/** A writable copy of a read-back 3-D: its colors narrow to `Color` or drop. */
export const writable3D = (value: ReadText3D | null): Text3D => {
  if (!value) return {};
  const {
    contourColor,
    extrusionColor,
    contourColorTransforms,
    extrusionColorTransforms,
    ...rest
  } = value;
  const contour = contourColor === undefined ? undefined : asColor(contourColor);
  const extrusion = extrusionColor === undefined ? undefined : asColor(extrusionColor);
  return {
    ...rest,
    ...(contour
      ? { contourColor: contour, ...(contourColorTransforms ? { contourColorTransforms } : {}) }
      : {}),
    ...(extrusion
      ? {
          extrusionColor: extrusion,
          ...(extrusionColorTransforms ? { extrusionColorTransforms } : {}),
        }
      : {}),
  };
};
