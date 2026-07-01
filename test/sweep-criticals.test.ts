// Regression net for the final generative-sweep findings (the batch after the
// boundary-validation pass): each test pins one defect where the writer emitted
// a `.pptx` PowerPoint marks corrupt — XML-illegal control characters, chart
// percentages outside their ECMA-376 simple-type ranges, unvalidated effect
// EMU, a scheme-color round-trip the setter rejected, importSlide leaving a
// dangling chart relationship, and invalid `ST_ShapeType` math tokens.

import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addBlankSlide,
  addSlideChart,
  addSlideShape,
  createPresentation,
  findSlideLayout,
  getShapeFillColor,
  getSlides,
  importSlide,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
  setShapeGlow,
  setShapeShadow,
  setSlideNotes,
} from '../src/api/index.ts';
import {
  expectSchemaValid,
  isSchemaValidationAvailable,
  type SchemaKind,
} from './lib/expect-schema-valid.ts';

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;
const decode = (b: Uint8Array): string => new TextDecoder().decode(b);
// Build control-char strings from char codes rather than embedding raw illegal
// bytes in this source file (which tooling and editors mangle).
const ctrl = (code: number): string => String.fromCharCode(code);

const rect = (pres: ReturnType<typeof createPresentation>, text?: string) =>
  addSlideShape(addBlankSlide(pres), {
    preset: 'rect',
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(2),
    ...(text === undefined ? {} : { text }),
  });

const validateParts = (pres: ReturnType<typeof createPresentation>): void => {
  const pkg = _internalPackageOf(pres);
  for (const part of pkg.parts) {
    let kind: SchemaKind | null = null;
    if (/\/slides\/[^/]*\.xml$/.test(part.name)) kind = 'pml';
    else if (/\/charts\/chart\d+\.xml$/.test(part.name)) kind = 'chart';
    if (kind) expectSchemaValid(decode(part.data), kind);
  }
};

describe('sweep: XML-illegal control characters', () => {
  // Serialization is synchronous (an authoring call re-serializes its part, and
  // `savePresentation` wraps a sync `.save()`), so an illegal character surfaces
  // as a synchronous throw. Wrap author-then-save so the assertion holds wherever
  // the serializer first sees the character.
  it('rejects a NUL (and other C0 controls) in shape text', () => {
    expect(() => {
      const pres = createPresentation();
      rect(pres, `before${ctrl(0x00)}after`);
      return savePresentation(pres);
    }).toThrow(/control character/i);
  });

  it('rejects a control character in slide notes', () => {
    expect(() => {
      const pres = createPresentation();
      setSlideNotes(addBlankSlide(pres), `note${ctrl(0x08)}body`); // backspace (U+0008)
      return savePresentation(pres);
    }).toThrow(/control character/i);
  });

  it('allows the three XML-legal whitespace controls (tab / LF / CR)', async () => {
    const pres = createPresentation();
    rect(pres, `a${ctrl(0x09)}b${ctrl(0x0a)}c${ctrl(0x0d)}d`);
    await expect(savePresentation(pres)).resolves.toBeInstanceOf(Uint8Array);
  });
});

describe('sweep: chart percentage ranges', () => {
  const chart = (
    pres: ReturnType<typeof createPresentation>,
    spec: Parameters<typeof addSlideChart>[1]['spec'],
  ) =>
    addSlideChart(addBlankSlide(pres), {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(8),
      h: inches(4.5),
      spec,
    });

  it('rejects gapWidthPct above 500 (ST_GapAmount)', () => {
    expect(() =>
      chart(createPresentation(), {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'S', values: [1, 2] }],
        gapWidthPct: 600,
      }),
    ).toThrow(RangeError);
  });

  it('rejects holeSizePct outside 1..90 (ST_HoleSize)', () => {
    expect(() =>
      chart(createPresentation(), {
        kind: 'doughnut',
        categories: ['A', 'B'],
        series: [{ name: 'S', values: [1, 2] }],
        holeSizePct: 95,
      }),
    ).toThrow(RangeError);
    expect(() =>
      chart(createPresentation(), {
        kind: 'doughnut',
        categories: ['A', 'B'],
        series: [{ name: 'S', values: [1, 2] }],
        holeSizePct: 0,
      }),
    ).toThrow(RangeError);
  });

  it('rejects firstSliceAngleDeg outside 0..360 (ST_FirstSliceAng)', () => {
    expect(() =>
      chart(createPresentation(), {
        kind: 'pie',
        categories: ['A', 'B'],
        series: [{ name: 'S', values: [1, 2] }],
        firstSliceAngleDeg: 361,
      }),
    ).toThrow(RangeError);
  });

  skipIfNoXmllint('valid extremes (gap 500, hole 90, angle 360) stay schema-valid', async () => {
    const pres = createPresentation();
    chart(pres, {
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2] }],
      gapWidthPct: 500,
    });
    chart(pres, {
      kind: 'doughnut',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2] }],
      holeSizePct: 90,
      firstSliceAngleDeg: 360,
    });
    validateParts(await loadPresentation(await savePresentation(pres)));
  });
});

describe('sweep: shape effect EMU validation', () => {
  it('rejects negative / non-finite / over-max shadow EMU', () => {
    expect(() => setShapeShadow(rect(createPresentation()), { blurEmu: -1 })).toThrow(RangeError);
    expect(() => setShapeShadow(rect(createPresentation()), { offsetEmu: Number.NaN })).toThrow(
      RangeError,
    );
    expect(() => setShapeShadow(rect(createPresentation()), { blurEmu: 27273042316901 })).toThrow(
      RangeError,
    );
    // Fractional EMU is rounded to a valid integer, not rejected.
    expect(() => setShapeShadow(rect(createPresentation()), { blurEmu: 50800.5 })).not.toThrow();
  });

  it('rejects out-of-range glow radius', () => {
    expect(() =>
      setShapeGlow(rect(createPresentation()), { color: '#FF0000', radiusEmu: -5 }),
    ).toThrow(RangeError);
  });
});

describe('sweep: scheme-color round-trip', () => {
  it('accepts the scheme:<token> string the getter returns', () => {
    const r = rect(createPresentation());
    setShapeFill(r, 'accent1');
    const read = getShapeFillColor(r);
    expect(read).toBe('scheme:accent1');
    // The whole point: feeding the getter's output back to the setter must work.
    expect(() => setShapeFill(r, read!)).not.toThrow();
    expect(getShapeFillColor(r)).toBe('scheme:accent1');
  });

  it('still rejects an unknown scheme token', () => {
    expect(() => setShapeFill(rect(createPresentation()), 'scheme:bogus')).toThrow();
  });
});

describe('sweep: importSlide drops charts without a dangling relationship', () => {
  it('produces a valid package when the source slide has a chart', async () => {
    const source = createPresentation();
    addSlideChart(addBlankSlide(source), {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'bar',
        categories: ['A', 'B'],
        series: [{ name: 'S', values: [1, 2] }],
      },
    });
    const sourceReloaded = await loadPresentation(await savePresentation(source));

    const target = createPresentation();
    const layout = findSlideLayout(target, 'Blank') ?? findSlideLayout(target, 'Title and Content');
    expect(layout).not.toBeNull();
    importSlide(target, getSlides(sourceReloaded)[0]!, layout!);

    const bytes = await savePresentation(target);
    const reloaded = await loadPresentation(bytes);
    expect(getSlides(reloaded)).toHaveLength(1);

    // No body r:id may reference a relationship the imported slide doesn't carry.
    const pkg = _internalPackageOf(reloaded);
    const slidePart = pkg.parts.find((p) => /\/ppt\/slides\/slide\d+\.xml$/.test(p.name))!;
    const relsPart = pkg.parts.find((p) =>
      /\/ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(p.name),
    );
    const relIds = new Set(
      relsPart ? [...decode(relsPart.data).matchAll(/Id="([^"]+)"/g)].map((m) => m[1]) : [],
    );
    const bodyRefs = [...decode(slidePart.data).matchAll(/r:(?:id|embed|link)="([^"]+)"/g)].map(
      (m) => m[1],
    );
    for (const ref of bodyRefs) expect(relIds.has(ref)).toBe(true);
  });
});

describe('sweep: ST_ShapeType math presets', () => {
  skipIfNoXmllint('emits schema-valid prstGeom for the math operator presets', async () => {
    const pres = createPresentation();
    for (const preset of [
      'mathPlus',
      'mathMinus',
      'mathMultiply',
      'mathDivide',
      'mathEqual',
      'mathNotEqual',
    ] as const) {
      addSlideShape(addBlankSlide(pres), {
        preset,
        x: inches(1),
        y: inches(1),
        w: inches(2),
        h: inches(2),
      });
    }
    validateParts(await loadPresentation(await savePresentation(pres)));
  });
});
