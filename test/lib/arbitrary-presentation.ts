// Property-test data model + builder for randomized presentations.
//
// `presentationArbitrary()` yields a `DeckSpec` — a plain, JSON-shaped
// description of a deck. `buildPresentation(spec)` realizes that spec
// through the public fn-API (`createPresentation` + `addSlide` + the
// `add*` / `set*` authoring calls) and returns the live
// `PresentationData`.
//
// Why split the random *spec* from the *builder*: fast-check shrinks the
// value the arbitrary produces. Keeping that value a tree of plain data
// (not an opaque PptxKit handle) means a failing case prints — and
// shrinks down to — a minimal, human-readable object the maintainer can
// paste straight into a regression test.
//
// Everything here stays strictly inside the documented public API and
// never constructs a state the types forbid: positions/sizes are whole
// EMU, colors are 6-hex or scheme tokens, presets/patterns are real
// ECMA-376 tokens, and text excludes the characters XML 1.0 cannot carry
// (control chars, and CR which XML parsers normalize to LF — both would
// make round-trip equality a lie rather than a bug).

import fc from 'fast-check';
import {
  addSlide,
  addSlideLine,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  emu,
  findSlideLayout,
  findSlideLayoutByType,
  type PatternPreset,
  type PresentationData,
  setShapeBullets,
  setShapeFill,
  setShapeFlip,
  setShapeGradientFill,
  setShapeNoFill,
  setShapePatternFill,
  setShapeRotation,
  setShapeRunFormat,
  type SlideData,
  type SlideShapeData,
} from '../../src/api/index.ts';

// Fixed seed so CI runs are reproducible. Bump deliberately when you want
// to explore a fresh region of the input space; a flake here is a real
// bug, not noise.
export const PROPERTY_SEED = 0x70_74_78_6b; // "ptxk"

// ---------------------------------------------------------------------------
// Spec model — the shrinkable description of a deck.

export type FillSpec =
  | { readonly kind: 'inherit' }
  | { readonly kind: 'none' }
  | { readonly kind: 'solid'; readonly hex: string }
  | {
      readonly kind: 'gradient';
      readonly fromHex: string;
      readonly toHex: string;
      readonly midHex: string | null;
      readonly angleDeg: number;
    }
  | {
      readonly kind: 'pattern';
      readonly preset: PatternPreset;
      readonly fgHex: string;
      readonly bgHex: string;
    };

export type BulletSpec =
  | 'bullet'
  | 'number'
  | 'none'
  | { readonly char: string }
  | { readonly autoNum: string };

export interface RunFormatSpec {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly sizePt: number;
  readonly font: string;
}

interface TextShapeCommon {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly text: string;
  readonly rotationDeg: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  readonly fill: FillSpec;
  readonly format: RunFormatSpec;
  readonly bullet: BulletSpec | null;
}

export interface TextBoxSpec extends TextShapeCommon {
  readonly kind: 'textbox';
}

export interface PresetSpec extends TextShapeCommon {
  readonly kind: 'preset';
  readonly preset: string;
}

export interface TableSpec {
  readonly kind: 'table';
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly rows: ReadonlyArray<ReadonlyArray<string>>;
  readonly firstRow: boolean;
  readonly bandRow: boolean;
}

export interface LineSpec {
  readonly kind: 'line';
  readonly fromX: number;
  readonly fromY: number;
  readonly toX: number;
  readonly toY: number;
}

export type ShapeSpec = TextBoxSpec | PresetSpec | TableSpec | LineSpec;

export interface SlideSpec {
  readonly shapes: ReadonlyArray<ShapeSpec>;
}

export interface DeckSpec {
  readonly size: '16:9' | '4:3';
  readonly slides: ReadonlyArray<SlideSpec>;
}

// ---------------------------------------------------------------------------
// Token pools — every value here is a real ECMA-376 token, so the builder
// never has to validate; the spec is correct by construction.

// A spread of preset autoshape geometries (ST_ShapeType, §20.1.10.55).
const PRESETS = [
  'rect',
  'roundRect',
  'ellipse',
  'triangle',
  'diamond',
  'parallelogram',
  'trapezoid',
  'pentagon',
  'hexagon',
  'star5',
  'rightArrow',
  'leftArrow',
  'chevron',
  'cloud',
  'heart',
  'plus',
  'cube',
  'can',
] as const;

// Preset fill patterns (ST_PresetPatternVal, §20.1.10.49).
const PATTERNS = [
  'pct50',
  'pct25',
  'dkUpDiag',
  'ltDnDiag',
  'cross',
  'wave',
  'dkHorz',
  'smGrid',
] as const;

const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Verdana', 'Georgia'] as const;

// Bullet variants chosen so each reads back unambiguously through
// `getParagraphBullet`: the '•' char maps to 'bullet' and the
// 'arabicPeriod' autonum maps to 'number', so the explicit { char } /
// { autoNum } cases deliberately avoid those two sentinels.
const BULLETS: ReadonlyArray<BulletSpec> = [
  'bullet',
  'number',
  'none',
  { char: '-' },
  { char: '*' },
  { autoNum: 'romanLcPeriod' },
  { autoNum: 'alphaUcParenR' },
];

// Single-code-point strings spanning ASCII, the five XML metacharacters
// (the interesting escaping fuzz), accented Latin, CJK, and an emoji
// (surrogate pair). No control chars and no '\r' / '\n' / '\t': the first
// three are illegal or whitespace-normalized in XML, and we keep shapes
// single-paragraph so text round-trips exactly.
const TEXT_CHARS = [
  ...'abcdefghijklmnopqrstuvwxyz',
  ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
  ...'0123456789',
  ' ',
  ...'!?.,;:-_/()[]@#%*+=',
  '&',
  '<',
  '>',
  '"',
  "'",
  'é',
  'ñ',
  'ü',
  'ç',
  '©',
  '日',
  '本',
  '語',
  '😀',
  '🚀',
] as const;

const TEXT_CHARS_NON_SPACE = TEXT_CHARS.filter((c) => c !== ' ');

// ---------------------------------------------------------------------------
// Leaf arbitraries.

// 6-digit uppercase hex. `getShapeFill` reports solid colors as
// `#` + uppercase, and `parseColor` uppercases on the way in, so
// generating uppercase keeps the round-trip assertion an exact equality.
const hexArb: fc.Arbitrary<string> = fc
  .array(fc.constantFrom(...'0123456789ABCDEF'), { minLength: 6, maxLength: 6 })
  .map((cs) => cs.join(''));

// Non-empty text with no leading/trailing whitespace: first char is
// always non-space and trailing spaces are stripped, so we never depend
// on `xml:space="preserve"` to survive the trip.
const textArb: fc.Arbitrary<string> = fc
  .tuple(
    fc.constantFrom(...TEXT_CHARS_NON_SPACE),
    fc.array(fc.constantFrom(...TEXT_CHARS), { maxLength: 20 }),
  )
  .map(([first, rest]) => `${first}${rest.join('')}`.replace(/\s+$/u, ''));

// Whole-EMU coordinates and extents. EMU is integer-valued (ST_Coordinate
// is xsd:long); a fractional value would trip PowerPoint's repair, so the
// public `emu()` helper rounds — we feed it integers to begin with. Bounds
// keep shapes on a 4:3 canvas (the smaller of the two we emit).
const posArb: fc.Arbitrary<number> = fc.integer({ min: 0, max: 6_000_000 });
const sizeArb: fc.Arbitrary<number> = fc.integer({ min: 12_700, max: 3_000_000 });
const rotationArb: fc.Arbitrary<number> = fc.integer({ min: 0, max: 359 });
const sizePtArb: fc.Arbitrary<number> = fc.integer({ min: 6, max: 72 });

const fillArb: fc.Arbitrary<FillSpec> = fc.oneof(
  fc.constant<FillSpec>({ kind: 'inherit' }),
  fc.constant<FillSpec>({ kind: 'none' }),
  hexArb.map<FillSpec>((hex) => ({ kind: 'solid', hex })),
  fc
    .record({ fromHex: hexArb, toHex: hexArb, midHex: fc.option(hexArb), angleDeg: rotationArb })
    .map<FillSpec>((g) => ({ kind: 'gradient', ...g })),
  fc
    .record({ preset: fc.constantFrom(...PATTERNS), fgHex: hexArb, bgHex: hexArb })
    .map<FillSpec>((p) => ({ kind: 'pattern', ...p })),
);

const formatArb: fc.Arbitrary<RunFormatSpec> = fc.record({
  bold: fc.boolean(),
  italic: fc.boolean(),
  sizePt: sizePtArb,
  font: fc.constantFrom(...FONTS),
});

const bulletArb: fc.Arbitrary<BulletSpec | null> = fc.option(fc.constantFrom(...BULLETS));

const textShapeCommonArb = fc.record<TextShapeCommon>({
  x: posArb,
  y: posArb,
  w: sizeArb,
  h: sizeArb,
  text: textArb,
  rotationDeg: rotationArb,
  flipH: fc.boolean(),
  flipV: fc.boolean(),
  fill: fillArb,
  format: formatArb,
  bullet: bulletArb,
});

const textBoxArb: fc.Arbitrary<ShapeSpec> = textShapeCommonArb.map((c) => ({
  kind: 'textbox',
  ...c,
}));

const presetArb: fc.Arbitrary<ShapeSpec> = fc
  .tuple(fc.constantFrom(...PRESETS), textShapeCommonArb)
  .map(([preset, c]) => ({ kind: 'preset', preset, ...c }));

// Rectangular table: a fixed column count drives every row, so
// `getTableDimensions` reads back the exact shape we authored.
const tableArb: fc.Arbitrary<ShapeSpec> = fc
  .record({
    x: posArb,
    y: posArb,
    w: sizeArb,
    h: sizeArb,
    rowCount: fc.integer({ min: 1, max: 3 }),
    colCount: fc.integer({ min: 1, max: 3 }),
    firstRow: fc.boolean(),
    bandRow: fc.boolean(),
  })
  .chain((t) =>
    fc
      .array(fc.array(textArb, { minLength: t.colCount, maxLength: t.colCount }), {
        minLength: t.rowCount,
        maxLength: t.rowCount,
      })
      .map<ShapeSpec>((rows) => ({
        kind: 'table',
        x: t.x,
        y: t.y,
        w: t.w,
        h: t.h,
        rows,
        firstRow: t.firstRow,
        bandRow: t.bandRow,
      })),
  );

const lineArb: fc.Arbitrary<ShapeSpec> = fc
  .record({ fromX: posArb, fromY: posArb, toX: posArb, toY: posArb })
  .map<ShapeSpec>((l) => ({ kind: 'line', ...l }));

const shapeArb: fc.Arbitrary<ShapeSpec> = fc.oneof(
  { weight: 4, arbitrary: textBoxArb },
  { weight: 4, arbitrary: presetArb },
  { weight: 2, arbitrary: tableArb },
  { weight: 1, arbitrary: lineArb },
);

const slideArb: fc.Arbitrary<SlideSpec> = fc.record({
  shapes: fc.array(shapeArb, { maxLength: 6 }),
});

/**
 * The top-level arbitrary: a deck of 1–3 slides, each carrying 0–6
 * randomized shapes. Pass to `fc.assert(fc.asyncProperty(...))`.
 */
export const presentationArbitrary = (): fc.Arbitrary<DeckSpec> =>
  fc.record({
    size: fc.constantFrom<'16:9' | '4:3'>('16:9', '4:3'),
    slides: fc.array(slideArb, { minLength: 1, maxLength: 3 }),
  });

// ---------------------------------------------------------------------------
// Builder — realizes a spec through the public API.

const applyFill = (shape: SlideShapeData, fill: FillSpec): void => {
  switch (fill.kind) {
    case 'inherit':
      return; // leave the shape's default fill choice in place
    case 'none':
      setShapeNoFill(shape);
      return;
    case 'solid':
      setShapeFill(shape, `#${fill.hex}`);
      return;
    case 'gradient': {
      const stops = [
        { offset: 0, color: `#${fill.fromHex}` },
        ...(fill.midHex !== null ? [{ offset: 0.5, color: `#${fill.midHex}` }] : []),
        { offset: 1, color: `#${fill.toHex}` },
      ];
      setShapeGradientFill(shape, { stops, angleDeg: fill.angleDeg });
      return;
    }
    case 'pattern':
      setShapePatternFill(shape, {
        preset: fill.preset,
        foreground: `#${fill.fgHex}`,
        background: `#${fill.bgHex}`,
      });
      return;
  }
};

const applyTextShape = (shape: SlideShapeData, spec: TextShapeCommon): void => {
  // Rotation and flip are always written explicitly (0 / false clears the
  // attribute) so the reload reads back a deterministic value.
  setShapeRotation(shape, spec.rotationDeg);
  setShapeFlip(shape, { horizontal: spec.flipH, vertical: spec.flipV });
  applyFill(shape, spec.fill);
  setShapeRunFormat(shape, 0, 0, {
    bold: spec.format.bold,
    italic: spec.format.italic,
    size: spec.format.sizePt,
    font: spec.format.font,
  });
  if (spec.bullet !== null) setShapeBullets(shape, spec.bullet);
};

const applyShape = (slide: SlideData, spec: ShapeSpec): void => {
  switch (spec.kind) {
    case 'textbox': {
      const shape = addSlideTextBox(slide, {
        x: emu(spec.x),
        y: emu(spec.y),
        w: emu(spec.w),
        h: emu(spec.h),
        text: spec.text,
      });
      applyTextShape(shape, spec);
      return;
    }
    case 'preset': {
      const shape = addSlideShape(slide, {
        preset: spec.preset,
        x: emu(spec.x),
        y: emu(spec.y),
        w: emu(spec.w),
        h: emu(spec.h),
        text: spec.text,
      });
      applyTextShape(shape, spec);
      return;
    }
    case 'table':
      addSlideTable(slide, {
        x: emu(spec.x),
        y: emu(spec.y),
        w: emu(spec.w),
        h: emu(spec.h),
        rows: spec.rows,
        firstRow: spec.firstRow,
        bandRow: spec.bandRow,
      });
      return;
    case 'line':
      addSlideLine(slide, {
        from: { x: emu(spec.fromX), y: emu(spec.fromY) },
        to: { x: emu(spec.toX), y: emu(spec.toY) },
      });
      return;
  }
};

/**
 * Realizes a `DeckSpec` into a live `PresentationData`. Every slide is
 * built on the deck's `Blank` layout, which carries no placeholders — so
 * the shapes the spec describes are exactly the trailing shapes of each
 * slide's shape tree, in authoring order.
 */
export const buildPresentation = (spec: DeckSpec): PresentationData => {
  const pres = createPresentation({ size: spec.size });
  const layout = findSlideLayoutByType(pres, 'blank') ?? findSlideLayout(pres, 'Blank');
  if (layout === null) throw new Error('buildPresentation: blank layout missing from fresh deck');
  for (const slideSpec of spec.slides) {
    const slide = addSlide(pres, { layout });
    for (const shapeSpec of slideSpec.shapes) applyShape(slide, shapeSpec);
  }
  return pres;
};
