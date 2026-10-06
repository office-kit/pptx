// Preset geometry (`<a:prstGeom>`, ECMA-376 §20.1.9.18) evaluated from the
// spec's own definitions (Part 1, Annex D, presetShapeDefinitions.xml), with
// the same guide evaluator `<a:custGeom>` uses: a preset is a custom geometry
// whose formulas the standard wrote down once for every file.
//
// Allowed imports: internal/xml (through custom-geometry.ts).

import {
  type CustomGeometry,
  type GeomCommand,
  type GeomPath,
  type GeomPoint,
  type GeomTextRect,
  type PathFillMode,
  GeomEvalError,
  builtinGuides,
  evalFormula,
  resolveToken,
} from './custom-geometry.ts';
import { PRESET_GEOMETRY } from './preset-geometry.generated.ts';

/** One pen command, as written in the definitions: an op letter, then guide names or literals. */
export type PresetPathCommand = readonly [string, ...string[]];

export interface PresetPathDefinition {
  readonly w?: number;
  readonly h?: number;
  /** Omitted for the schema default, `norm`. */
  readonly fill?: PathFillMode;
  /** Omitted for the schema default, `true`. */
  readonly stroke?: false;
  readonly commands: readonly PresetPathCommand[];
}

export interface PresetGeometryDefinition {
  /** Adjust values with their defaults, as `[name, fmla]`. */
  readonly av: readonly (readonly [string, string])[];
  readonly gd: readonly (readonly [string, string])[];
  readonly rect?: readonly [string, string, string, string];
  readonly paths: readonly PresetPathDefinition[];
}

const point = (x: string, y: string, guides: Map<string, number>): GeomPoint => ({
  x: resolveToken(x, guides),
  y: resolveToken(y, guides),
});

const evalCommand = (command: PresetPathCommand, guides: Map<string, number>): GeomCommand => {
  const [op, ...args] = command;
  const arg = (index: number): string => {
    const value = args[index];
    if (value === undefined) throw new GeomEvalError(`preset command ${op} is missing operands`);
    return value;
  };
  switch (op) {
    case 'M':
      return { kind: 'moveTo', pt: point(arg(0), arg(1), guides) };
    case 'L':
      return { kind: 'lnTo', pt: point(arg(0), arg(1), guides) };
    case 'A':
      return {
        kind: 'arcTo',
        wR: resolveToken(arg(0), guides),
        hR: resolveToken(arg(1), guides),
        stAng: resolveToken(arg(2), guides),
        swAng: resolveToken(arg(3), guides),
      };
    case 'Q':
      return {
        kind: 'quadBezTo',
        pts: [point(arg(0), arg(1), guides), point(arg(2), arg(3), guides)],
      };
    case 'C':
      return {
        kind: 'cubicBezTo',
        pts: [
          point(arg(0), arg(1), guides),
          point(arg(2), arg(3), guides),
          point(arg(4), arg(5), guides),
        ],
      };
    case 'Z':
      return { kind: 'close' };
    default:
      throw new GeomEvalError(`unknown preset command: ${op}`);
  }
};

/**
 * Evaluates `preset` at `w` × `h` EMU. `adjust` holds the shape's own
 * `<a:avLst>` values, which replace the preset's defaults of the same name.
 * Returns `null` for a name the standard does not define.
 */
export const evaluatePresetGeometry = (
  preset: string,
  w: number,
  h: number,
  adjust: Readonly<Record<string, number>>,
): CustomGeometry | null => {
  // `Object.hasOwn` keeps a name like `constructor` from reaching the prototype.
  if (!Object.hasOwn(PRESET_GEOMETRY, preset)) return null;
  const definition = PRESET_GEOMETRY[preset]!;
  const guides = builtinGuides(w, h);
  for (const [name, fmla] of definition.av) {
    guides.set(name, Object.hasOwn(adjust, name) ? adjust[name]! : evalFormula(fmla, guides));
  }
  for (const [name, fmla] of definition.gd) guides.set(name, evalFormula(fmla, guides));
  let textRect: GeomTextRect | null = null;
  if (definition.rect) {
    const [l, t, r, b] = definition.rect.map((token) => resolveToken(token, guides));
    textRect = { l: l!, t: t!, r: r!, b: b! };
  }
  const paths: GeomPath[] = definition.paths.map((path) => ({
    w: path.w ?? null,
    h: path.h ?? null,
    fill: path.fill ?? 'norm',
    stroke: path.stroke ?? true,
    commands: path.commands.map((command) => evalCommand(command, guides)),
  }));
  return { paths, textRect };
};
