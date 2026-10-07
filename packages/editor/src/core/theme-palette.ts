// The swatches of PowerPoint's colour palette: the ten theme colours, their
// five rows of tints and shades, and the ten standard colours. Shared by the
// colour picker button and the palette embedded in the Animations tab's
// Effect Options, so both offer — and write — the same colours.

import type { ColorTransform, PresentationTheme } from '@office-kit/pptx';
import { resolveColor } from './theme-color.ts';

export interface PaletteSwatch {
  /** A theme slot (`accent1`) or `#RRGGBB`. */
  readonly color: string;
  /** What it paints in this deck. */
  readonly paint: string;
  /** The untranslated name, `Accent 1` or `Dark Red`. */
  readonly name: string;
  readonly theme: boolean;
  readonly transforms?: readonly ColorTransform[];
  readonly shade?: 'Darker' | 'Lighter';
  readonly shadePercent?: number;
}

const THEME_SLOTS = [
  ['bg1', 'light1', 'Background 1'],
  ['tx1', 'dark1', 'Text 1'],
  ['bg2', 'light2', 'Background 2'],
  ['tx2', 'dark2', 'Text 2'],
  ['accent1', 'accent1', 'Accent 1'],
  ['accent2', 'accent2', 'Accent 2'],
  ['accent3', 'accent3', 'Accent 3'],
  ['accent4', 'accent4', 'Accent 4'],
  ['accent5', 'accent5', 'Accent 5'],
  ['accent6', 'accent6', 'Accent 6'],
] as const;

export const STANDARD_COLORS = [
  ['#C00000', 'Dark Red'],
  ['#FF0000', 'Red'],
  ['#FFC000', 'Orange'],
  ['#FFFF00', 'Yellow'],
  ['#92D050', 'Light Green'],
  ['#00B050', 'Green'],
  ['#00B0F0', 'Light Blue'],
  ['#0070C0', 'Blue'],
  ['#002060', 'Dark Blue'],
  ['#7030A0', 'Purple'],
] as const;

type ShadeTransform = Extract<ColorTransform, { value: number }>;
const lumMod = (value: number): ShadeTransform => ({ kind: 'lumMod', value });
const lumOff = (value: number): ShadeTransform => ({ kind: 'lumOff', value });

// PowerPoint's five rows under each theme colour, as `lumMod` / `lumOff`.
const DARKER_ROWS = [0.95, 0.85, 0.75, 0.65, 0.5].map((v) => [lumMod(v)]);
const LIGHTER_ROWS = [
  [lumMod(0.2), lumOff(0.8)],
  [lumMod(0.4), lumOff(0.6)],
  [lumMod(0.6), lumOff(0.4)],
  [lumMod(0.75)],
  [lumMod(0.5)],
];
const TEXT_ROWS = [
  [lumMod(0.5), lumOff(0.5)],
  [lumMod(0.65), lumOff(0.35)],
  [lumMod(0.75), lumOff(0.25)],
  [lumMod(0.85), lumOff(0.15)],
  [lumMod(0.95), lumOff(0.05)],
];
const BACKGROUND_2_ROWS = [0.9, 0.75, 0.5, 0.25, 0.1].map((v) => [lumMod(v)]);

const rowsFor = (color: string): readonly (readonly ShadeTransform[])[] =>
  color === 'bg1'
    ? DARKER_ROWS
    : color === 'tx1'
      ? TEXT_ROWS
      : color === 'bg2'
        ? BACKGROUND_2_ROWS
        : LIGHTER_ROWS;

const shadeOf = (
  color: string,
  row: number,
): { shade: 'Darker' | 'Lighter'; shadePercent: number } => {
  if (color === 'bg1') return { shade: 'Darker', shadePercent: [5, 15, 25, 35, 50][row]! };
  if (color === 'tx1') return { shade: 'Lighter', shadePercent: [50, 35, 25, 15, 5][row]! };
  if (color === 'bg2') return { shade: 'Darker', shadePercent: [10, 25, 50, 75, 90][row]! };
  return row < 3
    ? { shade: 'Lighter', shadePercent: [80, 60, 40][row]! }
    : { shade: 'Darker', shadePercent: [25, 50][row - 3]! };
};

/** The theme colours (with their tint and shade rows when asked) and the standard colours. */
export const paletteSwatches = (
  theme: PresentationTheme | null,
  withShades: boolean,
): PaletteSwatch[] => {
  const base: PaletteSwatch[] = [];
  if (theme) {
    for (const [color, slot, name] of THEME_SLOTS) {
      if (theme[slot]) base.push({ color, paint: theme[slot], name, theme: true });
    }
    if (withShades) {
      for (const row of [0, 1, 2, 3, 4]) {
        for (const [color, slot, name] of THEME_SLOTS) {
          if (!theme[slot]) continue;
          const transforms = rowsFor(color)[row]!;
          base.push({
            color,
            paint: resolveColor(color, transforms, theme) ?? '#000000',
            name,
            theme: true,
            transforms,
            ...shadeOf(color, row),
          });
        }
      }
    }
  }
  return [
    ...base,
    ...STANDARD_COLORS.map(([color, name]) => ({ color, paint: color, name, theme: false })),
  ];
};

/** Whether two lists of colour transforms say the same thing. */
export const sameColorTransforms = (
  left: readonly ColorTransform[] | undefined,
  right: readonly ColorTransform[] | undefined,
): boolean => {
  const a = left ?? [];
  const b = right ?? [];
  return (
    a.length === b.length &&
    a.every((transform, index) => {
      const other = b[index]!;
      if (transform.kind !== other.kind) return false;
      return 'value' in transform
        ? 'value' in other && transform.value === other.value
        : !('value' in other);
    })
  );
};
