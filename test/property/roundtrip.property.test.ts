// Property: round-trip safety.
//
// The project's core invariant is "parse(serialize(x)) is structurally
// equal to x for everything we support." This drives it with fast-check:
// for a randomized-but-valid deck, `savePresentation` → `loadPresentation`
// must preserve every structurally-relevant fact the public API exposes —
// slide count, shape order, kind, geometry, fill choice, text, run
// formatting, bullets, and table contents.
//
// Assertions are scoped to what the library actually *claims* round-trips:
// e.g. `getShapeFill` reports a gradient/pattern only by `kind` (not its
// stops), so we assert the kind, not the parameters. Solid fills and text
// it claims verbatim, so we assert exact equality there.

import { describe, expect, it } from 'vitest';
import fc from 'fast-check';
import {
  type BulletSpec,
  buildPresentation,
  type DeckSpec,
  type FillSpec,
  type PresetSpec,
  PROPERTY_SEED,
  presentationArbitrary,
  type RunFormatSpec,
  type ShapeSpec,
  type TextBoxSpec,
} from '../lib/arbitrary-presentation.ts';
import {
  getParagraphBullet,
  getShapeFill,
  getShapeFlip,
  getShapeKind,
  getShapePatternFill,
  getShapePosition,
  getShapePreset,
  getShapeRotation,
  getShapeRunFormat,
  getShapeSize,
  getShapeText,
  getSlides,
  getSlideShapes,
  getTableCell,
  getTableCellText,
  getTableDimensions,
  loadPresentation,
  type PresentationData,
  savePresentation,
  type SlideShapeData,
} from '../../src/api/index.ts';

const NUM_RUNS = 60;

const assertFill = (pres: PresentationData, shape: SlideShapeData, fill: FillSpec): void => {
  const got = getShapeFill(shape);
  switch (fill.kind) {
    case 'inherit':
      // The shape kept its builder default; we don't pin its exact value
      // (textbox = noFill, preset = inherit), only that nothing we set
      // turned it into a solid/gradient/pattern.
      expect(['inherit', 'none']).toContain(got.kind);
      return;
    case 'none':
      expect(got.kind).toBe('none');
      return;
    case 'solid':
      expect(got).toEqual({ kind: 'solid', color: `#${fill.hex}` });
      return;
    case 'gradient':
      expect(got.kind).toBe('gradient');
      return;
    case 'pattern':
      expect(got.kind).toBe('pattern');
      expect(getShapePatternFill(pres, shape)?.preset).toBe(fill.preset);
      return;
  }
};

const assertRunFormat = (shape: SlideShapeData, format: RunFormatSpec): void => {
  const got = getShapeRunFormat(shape, 0, 0);
  expect(got).not.toBeNull();
  expect(got?.size).toBe(format.sizePt);
  expect(got?.font).toBe(format.font);
  // Normalize: a cleared flag may read back as `false` or `undefined`.
  expect(Boolean(got?.bold)).toBe(format.bold);
  expect(Boolean(got?.italic)).toBe(format.italic);
};

const assertBullet = (shape: SlideShapeData, bullet: BulletSpec | null): void => {
  if (bullet === null) return;
  expect(getParagraphBullet(shape, 0)).toEqual(bullet);
};

const assertTextShape = (
  pres: PresentationData,
  shape: SlideShapeData,
  spec: TextBoxSpec | PresetSpec,
): void => {
  expect(getShapeKind(shape)).toBe('shape');
  if (spec.kind === 'preset') expect(getShapePreset(shape)).toBe(spec.preset);
  expect(getShapeText(shape)).toBe(spec.text);
  expect(getShapePosition(shape)).toEqual({ x: spec.x, y: spec.y });
  expect(getShapeSize(shape)).toEqual({ w: spec.w, h: spec.h });
  expect(getShapeRotation(shape)).toBe(spec.rotationDeg);
  expect(getShapeFlip(shape)).toEqual({ horizontal: spec.flipH, vertical: spec.flipV });
  assertFill(pres, shape, spec.fill);
  assertRunFormat(shape, spec.format);
  assertBullet(shape, spec.bullet);
};

const assertShape = (pres: PresentationData, shape: SlideShapeData, spec: ShapeSpec): void => {
  switch (spec.kind) {
    case 'textbox':
    case 'preset':
      assertTextShape(pres, shape, spec);
      return;
    case 'table': {
      expect(getShapeKind(shape)).toBe('graphicFrame');
      const dims = getTableDimensions(shape);
      expect(dims.rows).toBe(spec.rows.length);
      expect(dims.cols).toBe(spec.rows[0]?.length ?? 0);
      for (let r = 0; r < spec.rows.length; r++) {
        const row = spec.rows[r]!;
        for (let c = 0; c < row.length; c++) {
          expect(getTableCellText(getTableCell(shape, r, c))).toBe(row[c]);
        }
      }
      return;
    }
    case 'line':
      expect(getShapeKind(shape)).toBe('connector');
      return;
  }
};

const checkRoundTrip = async (spec: DeckSpec): Promise<void> => {
  const pres = buildPresentation(spec);
  const bytes = await savePresentation(pres);
  const reloaded = await loadPresentation(bytes);

  const slides = getSlides(reloaded);
  expect(slides.length).toBe(spec.slides.length);

  for (let i = 0; i < spec.slides.length; i++) {
    const slideSpec = spec.slides[i]!;
    const shapes = getSlideShapes(slides[i]!);
    const expected = slideSpec.shapes;
    // The Blank layout adds no placeholders, so the authored shapes are
    // the trailing N — slicing from the end stays correct even if a
    // future layout change prepends inherited shapes.
    const authored = shapes.slice(shapes.length - expected.length);
    expect(authored.length).toBe(expected.length);
    for (let j = 0; j < expected.length; j++) {
      assertShape(reloaded, authored[j]!, expected[j]!);
    }
  }
};

describe('property: presentation round-trip', () => {
  it('preserves structurally-relevant data through save → load', async () => {
    await fc.assert(fc.asyncProperty(presentationArbitrary(), checkRoundTrip), {
      seed: PROPERTY_SEED,
      numRuns: NUM_RUNS,
    });
  });
});
