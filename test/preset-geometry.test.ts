// `getPresetGeometry`: preset outlines evaluated from ECMA-376's own
// presetShapeDefinitions.xml.

import { describe, expect, it } from 'vitest';
import { getPresetGeometry, inches, type CustomGeometry } from '../src/api/index.ts';
import { SHAPE_PRESETS } from '../src/internal/enum-values.ts';

const SIZES = [
  { w: inches(1), h: inches(1) },
  { w: inches(4), h: inches(1) },
  { w: inches(1), h: inches(4) },
];

const numbers = (geometry: CustomGeometry): number[] =>
  geometry.paths.flatMap((path) =>
    path.commands.flatMap((command) => {
      switch (command.kind) {
        case 'moveTo':
        case 'lnTo':
          return [command.pt.x, command.pt.y];
        case 'arcTo':
          return [command.wR, command.hR, command.stAng, command.swAng];
        case 'quadBezTo':
        case 'cubicBezTo':
          return command.pts.flatMap((pt) => [pt.x, pt.y]);
        case 'close':
          return [];
      }
    }),
  );

describe('getPresetGeometry', () => {
  it.each(SHAPE_PRESETS)('evaluates %s to finite paths at any aspect ratio', (preset) => {
    for (const size of SIZES) {
      const geometry = getPresetGeometry(preset, size);
      expect(geometry).not.toBeNull();
      expect(geometry!.paths.length).toBeGreaterThan(0);
      expect(numbers(geometry!).every(Number.isFinite)).toBe(true);
      if (geometry!.textRect)
        expect(Object.values(geometry!.textRect).every(Number.isFinite)).toBe(true);
    }
  });

  it('defines upArrow, which the published definitions omit, as downArrow upside down', () => {
    const size = { w: inches(2), h: inches(3) };
    const flip = (geometry: CustomGeometry) =>
      geometry.paths[0]!.commands.flatMap((command) =>
        command.kind === 'moveTo' || command.kind === 'lnTo'
          ? [`${Math.round(command.pt.x)},${Math.round(size.h - command.pt.y)}`]
          : [],
      );
    const down = getPresetGeometry('downArrow', size)!;
    const up = getPresetGeometry('upArrow', size)!;
    expect(new Set(flip(up))).toEqual(
      new Set(
        down.paths[0]!.commands.flatMap((command) =>
          command.kind === 'moveTo' || command.kind === 'lnTo'
            ? [`${Math.round(command.pt.x)},${Math.round(command.pt.y)}`]
            : [],
        ),
      ),
    );
  });

  it("lets the shape's adjust values replace the preset defaults", () => {
    const size = { w: inches(2), h: inches(1) };
    const start = (adjust?: Record<string, number>) => {
      const command = getPresetGeometry('roundRect', size, adjust)!.paths[0]!.commands[0]!;
      return command.kind === 'moveTo' ? command.pt : null;
    };
    // The outline starts where the top-left corner's arc ends: 16667/100000
    // of the short side down by default, at the very corner with no rounding.
    expect(start()).toEqual({ x: 0, y: (inches(1) * 16667) / 100000 });
    expect(start({ adj: 0 })).toEqual({ x: 0, y: 0 });
  });

  it('keeps the shading of 3-D presets and their text rectangle', () => {
    const cube = getPresetGeometry('cube', { w: inches(2), h: inches(2) })!;
    expect(cube.paths.map((path) => path.fill)).toEqual([
      'norm',
      'darkenLess',
      'lightenLess',
      'none',
    ]);
    expect(cube.textRect).toEqual({ l: 0, t: inches(0.5), r: inches(1.5), b: inches(2) });
  });

  it('returns null for names ECMA-376 does not define', () => {
    expect(getPresetGeometry('notAShape', SIZES[0]!)).toBeNull();
    expect(getPresetGeometry('constructor', SIZES[0]!)).toBeNull();
  });
});
