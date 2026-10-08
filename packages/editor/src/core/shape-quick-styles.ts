import {
  setShapeStyle,
  setShapeFill,
  setShapeGradientFill,
  setShapeNoFill,
  setShapeNoStroke,
  setShapeStroke,
  setShapeStrokeCap,
  setShapeStrokeCompound,
  setShapeStrokeDash,
  setShapeStrokeJoin,
  setShapeStrokeArrow,
  type ShapeStyleOptions,
  type SlideShapeData,
} from '@office-kit/pptx';

export const quickStyleColors = [
  'dk1',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
] as const;
export type QuickStyleColor = (typeof quickStyleColors)[number];

export const themeQuickStyles = [
  'Colored Outline',
  'Colored Fill',
  'Light 1 Outline, Colored Fill',
  'Subtle Effect',
  'Moderate Effect',
  'Intense Effect',
] as const;
export const presetQuickStyles = [
  'Transparent',
  'Transparent, Colored Outline',
  'Semitransparent',
  'Colored Fill, No Outline',
  'Gradient Fill, No Outline',
] as const;
export type ShapeQuickStyle =
  | (typeof themeQuickStyles)[number]
  | (typeof presetQuickStyles)[number];

// The reference desktop app's (Mac) gallery stores theme styles as references, rather than
// copying the current theme's paint. Keeping those references lets a later
// theme change update the shape in the same way.
export function quickStyleReferences(
  style: ShapeQuickStyle,
  color: QuickStyleColor,
): ShapeStyleOptions {
  const line = (idx: number, slot: QuickStyleColor | 'lt1' = color) => ({ idx, color: slot });
  const fill = (idx: number, slot: QuickStyleColor | 'lt1' = color) => ({ idx, color: slot });
  const effect = (idx: number) => ({ idx, color });
  const font = (slot: QuickStyleColor | 'lt1' | 'dk1') => ({ idx: 'minor' as const, color: slot });
  switch (style) {
    case 'Colored Outline':
      return { line: line(2), fill: fill(1, 'lt1'), effect: effect(0), font: font('dk1') };
    case 'Colored Fill':
      return {
        line: { ...line(2), colorTransforms: [{ kind: 'shade', value: 0.15 }] },
        fill: fill(1),
        effect: effect(0),
        font: font('lt1'),
      };
    case 'Light 1 Outline, Colored Fill':
      return { line: line(3, 'lt1'), fill: fill(1), effect: effect(1), font: font('lt1') };
    case 'Subtle Effect':
      return { line: line(1), fill: fill(2), effect: effect(1), font: font('dk1') };
    case 'Moderate Effect':
      return { line: line(1), fill: fill(3), effect: effect(2), font: font('lt1') };
    case 'Intense Effect':
      return { line: line(0), fill: fill(3), effect: effect(3), font: font('lt1') };
    default:
      return {
        line: { idx: 0, color: '#000000' },
        fill: { idx: 0, color: '#000000' },
        effect: { idx: 0, color: '#000000' },
        font: font(
          style === 'Transparent' || style === 'Transparent, Colored Outline' ? color : 'lt1',
        ),
      };
  }
}

export function applyShapeQuickStyle(
  shape: SlideShapeData,
  style: ShapeQuickStyle,
  color: QuickStyleColor,
): void {
  setShapeStyle(shape, quickStyleReferences(style, color));
  switch (style) {
    case 'Transparent':
    case 'Transparent, Colored Outline':
      setShapeNoFill(shape);
      break;
    case 'Semitransparent':
      setShapeFill(shape, { color, opacity: 0.5 });
      break;
    case 'Colored Fill, No Outline':
      setShapeFill(shape, color);
      break;
    case 'Gradient Fill, No Outline':
      setShapeGradientFill(shape, {
        angleDeg: 270,
        scaled: true,
        rotateWithShape: true,
        tileRect: { left: 0, top: 0, right: 0, bottom: 0 },
        stops: [
          { offset: 0, color, colorTransforms: [{ kind: 'lumMod', value: 0.67 }] },
          {
            offset: 0.48,
            color,
            colorTransforms: [
              { kind: 'lumMod', value: 0.97 },
              { kind: 'lumOff', value: 0.03 },
            ],
          },
          {
            offset: 1,
            color,
            colorTransforms: [
              { kind: 'lumMod', value: 0.6 },
              { kind: 'lumOff', value: 0.4 },
            ],
          },
        ],
      });
      break;
    default:
      return;
  }
  if (style === 'Transparent, Colored Outline') {
    setShapeStroke(shape, { color, widthEmu: 9525 });
    setShapeStrokeCap(shape, 'flat');
    setShapeStrokeCompound(shape, 'sng');
    setShapeStrokeDash(shape, 'solid');
    setShapeStrokeJoin(shape, 'round');
    setShapeStrokeArrow(shape, 'head', { type: 'none', width: 'med', length: 'med' });
    setShapeStrokeArrow(shape, 'tail', { type: 'none', width: 'med', length: 'med' });
  } else setShapeNoStroke(shape);
}
