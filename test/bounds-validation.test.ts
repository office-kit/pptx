// Regression net for the boundary-validation bug class: a batch of authoring
// setters used to serialize caller-supplied numbers/strings straight into
// constrained OOXML attributes, so an out-of-range value produced a `.pptx`
// PowerPoint marks corrupt. Each authoring entry point now validates at the
// boundary (src/internal/bounds.ts). These tests pin that an out-of-range value
// THROWS (instead of emitting invalid XML), and that a valid extreme stays
// schema-valid. See the fuzz-sweep findings these came from.

import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addBlankSlide,
  addSlideChart,
  addSlideImage,
  addSlideLine,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  getTableCell,
  inches,
  loadPresentation,
  pt,
  savePresentation,
  setShapeAnimation,
  setShapePatternFill,
  setShapeRunFormat,
  setShapeStroke,
  setShapeTextBodyRotationDeg,
  setShapeTextColumns,
  setShapeTextMargins,
  setSlideTransition,
  setTableCellBorders,
  setTableCellMargins,
  setTableColumnWidth,
  setTableRowHeight,
  setTableStyleId,
} from '../src/api/index.ts';
import { buildPng } from './lib/build-png.ts';
import {
  expectSchemaValid,
  isSchemaValidationAvailable,
  type SchemaKind,
} from './lib/expect-schema-valid.ts';

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;
const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

const kindFor = (name: string): SchemaKind | null => {
  if (/\/(slides|notesSlides)\/[^/]*\.xml$/.test(name)) return 'pml';
  if (/\/charts\/chart\d+\.xml$/.test(name)) return 'chart';
  return null;
};

// Saves + reloads `pres` and validates every authored part against its schema.
const expectDeckSchemaValid = async (
  pres: ReturnType<typeof createPresentation>,
): Promise<void> => {
  const bytes = await savePresentation(pres);
  const pkg = _internalPackageOf(await loadPresentation(bytes));
  for (const part of pkg.parts) {
    const kind = kindFor(part.name);
    if (kind) expectSchemaValid(decode(part.data), kind);
  }
};

const rect = (pres: ReturnType<typeof createPresentation>) =>
  addSlideShape(addBlankSlide(pres), {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(2),
    text: 'X',
  });

const table = (pres: ReturnType<typeof createPresentation>) =>
  addSlideTable(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(6),
    h: inches(3),
    rows: [
      ['a', 'b'],
      ['c', 'd'],
    ],
  });

describe('bounds: text run formatting', () => {
  it('rejects out-of-range font size and char spacing', () => {
    const s = rect(createPresentation());
    expect(() => setShapeRunFormat(s, 0, 0, { size: 0.5 })).toThrow(RangeError);
    expect(() => setShapeRunFormat(s, 0, 0, { size: 5000 })).toThrow(RangeError);
    expect(() => setShapeRunFormat(s, 0, 0, { spc: 9_999_999 })).toThrow(RangeError);
    expect(() => setShapeRunFormat(s, 0, 0, { spc: -500_000 })).toThrow(RangeError);
  });

  it('accepts the 3-digit hex shorthand for run color and highlight', () => {
    const s = rect(createPresentation());
    expect(() => setShapeRunFormat(s, 0, 0, { color: '#f00' })).not.toThrow();
    expect(() => setShapeRunFormat(s, 0, 0, { highlight: '#abc' })).not.toThrow();
  });

  skipIfNoXmllint('a valid extreme font size stays schema-valid', async () => {
    const pres = createPresentation();
    setShapeRunFormat(rect(pres), 0, 0, { size: 4000 });
    await expectDeckSchemaValid(pres);
  });
});

describe('bounds: tables', () => {
  it('normalizes a lowercase GUID and rejects non-GUID style ids', () => {
    const t = table(createPresentation());
    // Lowercase is accepted and normalized (no throw).
    expect(() => setTableStyleId(t, '{c8de2e6a-6fb6-4bfa-acec-8d87b36ff2c3}')).not.toThrow();
    expect(() => setTableStyleId(t, 'MyCoolStyle')).toThrow(RangeError);
    expect(() => setTableStyleId(t, '')).toThrow(RangeError);
  });

  it('rejects out-of-range cell border width, column width, row height, margins', () => {
    const t = table(createPresentation());
    const c = getTableCell(t, 0, 0);
    expect(() => setTableCellBorders(c, { left: { color: '#000000', widthEmu: -5 } })).toThrow(
      RangeError,
    );
    expect(() =>
      setTableCellBorders(c, { left: { color: '#000000', widthEmu: 99_999_999 } }),
    ).toThrow(RangeError);
    expect(() => setTableColumnWidth(t, 0, 3e13 as ReturnType<typeof inches>)).toThrow(RangeError);
    expect(() => setTableRowHeight(t, 0, 3e13 as ReturnType<typeof inches>)).toThrow(RangeError);
    expect(() => setTableCellMargins(c, { left: 3e9 })).toThrow(RangeError);
  });

  it('rejects a table with negative width/height', () => {
    const pres = createPresentation();
    expect(() =>
      addSlideTable(addBlankSlide(pres), {
        x: inches(1),
        y: inches(1),
        w: -inches(3) as ReturnType<typeof inches>,
        h: inches(2),
        rows: [['a', 'b']],
      }),
    ).toThrow(RangeError);
  });

  skipIfNoXmllint('a normalized lowercase GUID is schema-valid', async () => {
    const pres = createPresentation();
    setTableStyleId(table(pres), '{c8de2e6a-6fb6-4bfa-acec-8d87b36ff2c3}');
    await expectDeckSchemaValid(pres);
  });
});

describe('bounds: charts', () => {
  it('rejects out-of-range overlap and series line width', () => {
    const pres = createPresentation();
    expect(() =>
      addSlideChart(addBlankSlide(pres), {
        x: inches(0.5),
        y: inches(0.5),
        w: inches(8),
        h: inches(4.5),
        spec: {
          kind: 'column',
          categories: ['A', 'B'],
          series: [{ name: 'S', values: [1, 2] }],
          overlapPct: 101,
        },
      }),
    ).toThrow(RangeError);
    expect(() =>
      addSlideChart(addBlankSlide(pres), {
        x: inches(0.5),
        y: inches(0.5),
        w: inches(8),
        h: inches(4.5),
        spec: {
          kind: 'line',
          categories: ['A', 'B'],
          series: [{ name: 'S', values: [1, 2], lineWidthEmu: -5 }],
        },
      }),
    ).toThrow(RangeError);
  });
});

describe('bounds: animations and transitions', () => {
  it('rounds fractional but rejects negative/huge animation duration', () => {
    const s = rect(createPresentation());
    // A fractional ms is rounded to a valid integer (dur is unsignedInt), not rejected.
    expect(() => setShapeAnimation(s, { effect: 'fadeIn', durationMs: 500.5 })).not.toThrow();
    expect(() => setShapeAnimation(s, { effect: 'fadeIn', durationMs: -100 })).toThrow(RangeError);
    expect(() => setShapeAnimation(s, { effect: 'fadeIn', durationMs: 5e9 })).toThrow(RangeError);
  });

  it('rejects out-of-range advance time and unknown/empty transition effect', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    expect(() => setSlideTransition(slide, { effect: 'fade', advanceAfterMs: -100 })).toThrow(
      RangeError,
    );
    expect(() => setSlideTransition(slide, { effect: 'fade', advanceAfterMs: 5e9 })).toThrow(
      RangeError,
    );
    expect(() => setSlideTransition(slide, { effect: '' })).toThrow(RangeError);
    expect(() => setSlideTransition(slide, { effect: 'bogusEffect' })).toThrow(RangeError);
  });
});

describe('bounds: connectors and strokes', () => {
  it('rejects out-of-range line / stroke width', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    expect(() =>
      addSlideLine(slide, {
        from: { x: inches(1), y: inches(1) },
        to: { x: inches(5), y: inches(3) },
        widthEmu: -pt(2),
      }),
    ).toThrow(RangeError);
    expect(() =>
      addSlideLine(slide, {
        from: { x: inches(1), y: inches(1) },
        to: { x: inches(5), y: inches(3) },
        widthEmu: 20_116_801,
      }),
    ).toThrow(RangeError);
    const sh = rect(pres);
    expect(() => setShapeStroke(sh, { color: '#000', widthEmu: -5 })).toThrow(RangeError);
  });
});

describe('bounds: text box columns / margins / rotation', () => {
  const textBox = (pres: ReturnType<typeof createPresentation>) =>
    addSlideTextBox(addBlankSlide(pres), {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      text: 'Hi',
    });

  it('rejects column count > 16, fractional count, negative gap', () => {
    const s = textBox(createPresentation());
    expect(() => setShapeTextColumns(s, { count: 20 })).toThrow(RangeError);
    expect(() => setShapeTextColumns(s, { count: 2, gapEmu: -500 })).toThrow(RangeError);
  });

  it('rejects out-of-range text margins and body rotation', () => {
    const s = textBox(createPresentation());
    expect(() => setShapeTextMargins(s, { left: 3_000_000_000 })).toThrow(RangeError);
    expect(() => setShapeTextBodyRotationDeg(s, 40_000)).toThrow(RangeError);
  });
});

describe('bounds: fills and shape geometry', () => {
  it('rejects an out-of-enum pattern preset', () => {
    const s = rect(createPresentation());
    expect(() =>
      setShapePatternFill(s, {
        preset: 'bogusPreset' as never,
        foreground: '#000000',
        background: '#FFFFFF',
      }),
    ).toThrow(RangeError);
  });

  it('rejects a shape positioned/sized out of range', () => {
    const pres = createPresentation();
    expect(() =>
      addSlideShape(addBlankSlide(pres), {
        preset: 'rect',
        x: inches(1),
        y: inches(1),
        w: 3e13 as ReturnType<typeof inches>,
        h: inches(2),
      }),
    ).toThrow(RangeError);
  });

  it('rejects an image sized out of range', () => {
    const pres = createPresentation();
    expect(() =>
      addSlideImage(addBlankSlide(pres), buildPng(8, 8, [1, 2, 3]), {
        x: inches(1),
        y: inches(1),
        w: -inches(3) as ReturnType<typeof inches>,
        h: inches(2),
      }),
    ).toThrow(RangeError);
  });
});
