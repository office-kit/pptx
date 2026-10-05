// The RGB a color paints in this deck once PowerPoint's tint/shade transforms
// apply. Swatches and gallery previews need the same answer the renderer gives,
// so this goes through the library's resolver rather than its own HSL math.

import {
  resolveDrawingColor,
  type Color,
  type ColorTransform,
  type PresentationTheme,
} from '@office-kit/pptx';

const DML = 'http://schemas.openxmlformats.org/drawingml/2006/main';
// ST_Percentage is in 1/1000 percent; hue angles are in 1/60000 degree.
const PERCENTAGE_UNITS = 100000;
const ANGLE_UNITS = 60000;

type ColorElement = Parameters<typeof resolveDrawingColor>[0];

const element = (
  localName: string,
  value: string | null,
  children: ColorElement[] = [],
): ColorElement => ({
  kind: 'element',
  name: { prefix: 'a', localName, namespaceURI: DML },
  attrs:
    value === null ? [] : [{ name: { prefix: '', localName: 'val', namespaceURI: '' }, value }],
  prefixDecls: new Map<string, string>(),
  children,
});

/** `#RRGGBB`, or `null` when a theme slot has no theme to resolve against. */
export function resolveColor(
  color: Color,
  transforms: readonly ColorTransform[],
  theme: PresentationTheme | null,
): string | null {
  const adjustments = transforms.map((transform) =>
    element(
      transform.kind,
      'value' in transform
        ? String(
            Math.round(
              transform.value *
                (transform.kind === 'hue' || transform.kind === 'hueOff'
                  ? ANGLE_UNITS
                  : PERCENTAGE_UNITS),
            ),
          )
        : null,
    ),
  );
  const base = color.startsWith('#')
    ? element('srgbClr', color.slice(1), adjustments)
    : element('schemeClr', color.replace(/^scheme:/, ''), adjustments);
  return resolveDrawingColor(base, theme);
}
