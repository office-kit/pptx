// Schema coverage for the from-scratch authoring API.
//
// Regression net for a batch of "generates but is schema-invalid / wrong"
// defects: every probe builds a deck through the public API, saves + reloads,
// and runs each authored part (slides / charts / notes) through `xmllint`
// against the ECMA-376 XSDs. Spec-deviations the schema can't catch (a radial
// gradient silently emitted as linear, etc.) are pinned with content checks.

import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addBlankSlide,
  addSlideChart,
  addSlideImage,
  addSlideLine,
  addSlideShape,
  addSlideTable,
  createPresentation,
  getParagraphLineSpacing,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  pt,
  savePresentation,
  setParagraphLineSpacing,
  setParagraphSpacing,
  setShapeAnimation,
  setShapeFill,
  setShapeGradientFill,
  setShapeImageBrightness,
  setShapeImageContrast,
  setShapeRunFormat,
  setShapeStroke,
  setShapeStrokeArrow,
  setShapeStrokeDash,
  setShapeStrokeJoin,
  setShapeText,
  setSlideNotes,
  setSlideTransition,
  getTableCell,
  setTableCellBorders,
  setTableCellFill,
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

// Saves + reloads `pres`, validates each authored part against its schema, and
// returns the concatenated XML of those parts for content assertions.
const authoredXml = async (pres: ReturnType<typeof createPresentation>): Promise<string> => {
  const bytes = await savePresentation(pres);
  const pkg = _internalPackageOf(await loadPresentation(bytes));
  let combined = '';
  for (const part of pkg.parts) {
    const kind = kindFor(part.name);
    if (kind === null) continue;
    const xml = decode(part.data);
    expectSchemaValid(xml, kind);
    combined += `\n<!-- ${part.name} -->\n${xml}`;
  }
  return combined;
};

const shape = (pres: ReturnType<typeof createPresentation>, preset = 'rect') =>
  addSlideShape(addBlankSlide(pres), {
    preset,
    x: inches(1),
    y: inches(1),
    w: inches(3),
    h: inches(2),
    text: 'X',
  });

describe('schema coverage: text run + paragraph ordering', () => {
  skipIfNoXmllint('color + highlight on one run stays in schema order', async () => {
    const pres = createPresentation();
    setShapeRunFormat(shape(pres), 0, 0, { color: '#112233', highlight: '#FFFF00', font: 'Arial' });
    await authoredXml(pres);
  });

  skipIfNoXmllint('paragraph spacing after a bullet stays in schema order', async () => {
    const pres = createPresentation();
    const s = shape(pres);
    setShapeText(s, 'a\nb\nc', { bullets: 'bullet' });
    setParagraphSpacing(s, 0, { beforePts: 6, afterPts: 6 });
    await authoredXml(pres);
  });

  skipIfNoXmllint('setParagraphLineSpacing writes lnSpc first', async () => {
    const pres = createPresentation();
    const s = shape(pres);
    setShapeText(s, 'a\nb', { bullets: 'bullet' });
    setParagraphLineSpacing(s, 0, { kind: 'pct', value: 1.5 });
    const xml = await authoredXml(pres);
    expect(xml).toContain('<a:lnSpc>');
    // 1.5x → `<a:spcPct val="150000"/>` (val is 1000ths of a percent).
    expect(xml).toContain('<a:spcPct val="150000"/>');
  });

  it('setParagraphLineSpacing converts both modes, clears, and rejects bad input', () => {
    const pres = createPresentation();
    const s = shape(pres);
    setShapeText(s, 'a\nb', { bullets: 'bullet' });

    // pct mode round-trips through the getter (1.5x).
    setParagraphLineSpacing(s, 0, { kind: 'pct', value: 1.5 });
    expect(getParagraphLineSpacing(s, 0)).toEqual({ kind: 'pct', value: 1.5 });

    // pts mode: 24pt is stored as `<a:spcPts val="2400">` (×100) and reads back
    // as 24 — a wrong write factor would fail this round-trip.
    setParagraphLineSpacing(s, 0, { kind: 'pts', value: 24 });
    expect(getParagraphLineSpacing(s, 0)).toEqual({ kind: 'pts', value: 24 });

    // null clears the override.
    setParagraphLineSpacing(s, 0, null);
    expect(getParagraphLineSpacing(s, 0)).toBeNull();

    // Negative / non-finite values are rejected at the boundary.
    expect(() => setParagraphLineSpacing(s, 0, { kind: 'pct', value: -1 })).toThrow(RangeError);
    expect(() => setParagraphLineSpacing(s, 0, { kind: 'pts', value: Number.NaN })).toThrow(
      RangeError,
    );
  });
});

describe('schema coverage: stroke sub-element ordering', () => {
  skipIfNoXmllint('dash + arrowheads + join combine in schema order', async () => {
    const pres = createPresentation();
    const s = shape(pres);
    setShapeStroke(s, { color: '#000000', widthEmu: pt(2) });
    setShapeStrokeDash(s, 'dash');
    setShapeStrokeArrow(s, 'tail', { type: 'triangle', width: 'med', length: 'med' });
    setShapeStrokeArrow(s, 'head', { type: 'oval' });
    setShapeStrokeJoin(s, 'round');
    await authoredXml(pres);
  });
});

describe('schema coverage: fills', () => {
  skipIfNoXmllint('3-digit hex shorthand is accepted', async () => {
    const pres = createPresentation();
    setShapeFill(shape(pres), '#f0a');
    await authoredXml(pres);
  });

  skipIfNoXmllint('radial gradient emits <a:path>, not a downgraded <a:lin>', async () => {
    const pres = createPresentation();
    setShapeGradientFill(shape(pres), {
      stops: [
        { offset: 0, color: '#FFFFFF' },
        { offset: 1, color: '#000000' },
      ],
      path: 'circle',
      focus: { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 },
    });
    const xml = await authoredXml(pres);
    expect(xml).toContain('<a:path');
    expect(xml).toContain('<a:fillToRect');
  });
});

describe('schema coverage: tables', () => {
  skipIfNoXmllint('whitespace-padded / multi-line cells stay valid', async () => {
    const pres = createPresentation();
    addSlideTable(addBlankSlide(pres), {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(3),
      rows: [
        [' leading', 'trailing '],
        ['mid\tdle', 'two\nlines'],
      ],
    });
    const xml = await authoredXml(pres);
    expect(xml).not.toContain('xml:space');
    // A `\n` splits the cell into separate paragraphs — each line is its own
    // `<a:t>`, not one run carrying an embedded newline (the old behavior).
    expect(xml).toContain('<a:t>two</a:t>');
    expect(xml).toContain('<a:t>lines</a:t>');
  });

  skipIfNoXmllint('cell fill then border keeps tcPr children ordered', async () => {
    const pres = createPresentation();
    const t = addSlideTable(addBlankSlide(pres), {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(3),
      rows: [
        ['a', 'b'],
        ['1', '2'],
      ],
    });
    const c = getTableCell(t, 0, 0)!;
    setTableCellFill(c, '#FFEEAA');
    setTableCellBorders(c, {
      left: { color: '#000000', widthEmu: pt(1) },
      bottom: { color: '#FF0000', widthEmu: pt(2) },
    });
    await authoredXml(pres);
  });
});

describe('schema coverage: transitions', () => {
  skipIfNoXmllint("effect 'none' emits no effect child", async () => {
    const pres = createPresentation();
    setSlideTransition(addBlankSlide(pres), { effect: 'none' });
    const xml = await authoredXml(pres);
    expect(xml).not.toContain('<p:none');
  });

  skipIfNoXmllint('an incompatible attribute is dropped, not emitted', async () => {
    const pres = createPresentation();
    // `dir` is invalid on a fade (CT_OptionalBlackTransition) — must be dropped.
    setSlideTransition(addBlankSlide(pres), { effect: 'fade', direction: 'l' });
    const xml = await authoredXml(pres);
    expect(xml).not.toContain('<p:fade dir');
  });
});

describe('schema coverage: charts', () => {
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

  skipIfNoXmllint('valueAxis min/max emit in max-before-min order', async () => {
    const pres = createPresentation();
    chart(pres, {
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2] }],
      valueAxis: { min: 0, max: 10 },
    });
    const xml = await authoredXml(pres);
    expect(xml.indexOf('<c:max')).toBeGreaterThan(-1);
    expect(xml.indexOf('<c:max')).toBeLessThan(xml.indexOf('<c:min'));
  });

  skipIfNoXmllint('marker/smooth/invertIfNegative are gated to valid kinds', async () => {
    const pres = createPresentation();
    chart(pres, {
      kind: 'area',
      categories: ['A', 'B'],
      series: [
        {
          name: 'S',
          values: [1, -2],
          smooth: true,
          invertIfNegative: true,
          markerSymbol: 'circle',
        },
      ],
    });
    const xml = await authoredXml(pres);
    expect(xml).not.toContain('<c:smooth');
    expect(xml).not.toContain('invertIfNegative');
    expect(xml).not.toContain('<c:marker');
  });

  skipIfNoXmllint('trendline is dropped on pie series (CT_PieSer has no trendline)', async () => {
    const pres = createPresentation();
    chart(pres, {
      kind: 'pie',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2], trendline: { type: 'linear' } }],
    });
    const xml = await authoredXml(pres);
    // CT_PieSer permits no <c:trendline>; emitting it would fail dml-chart.xsd.
    expect(xml).not.toContain('<c:trendline');
  });
});

describe('schema coverage: images, notes, connectors, animation', () => {
  skipIfNoXmllint('brightness + contrast share one valid <a:lum>', async () => {
    const pres = createPresentation();
    const p = addSlideImage(addBlankSlide(pres), buildPng(32, 32, [9, 9, 9]), {
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(2),
    });
    setShapeImageBrightness(p, 0.2);
    setShapeImageContrast(p, -0.3);
    const xml = await authoredXml(pres);
    expect(xml).toContain('<a:lum');
    expect(xml).not.toContain('lumOff');
    expect(xml).not.toContain('lumMod');
  });

  skipIfNoXmllint('notes slide part root is <p:notes>', async () => {
    const pres = createPresentation();
    setSlideNotes(addBlankSlide(pres), 'speaker notes');
    const xml = await authoredXml(pres);
    expect(xml).toContain('<p:notes');
    expect(xml).not.toContain('<p:notesSlide');
  });

  skipIfNoXmllint('an unstyled connector is visible (lnRef style)', async () => {
    const pres = createPresentation();
    addSlideLine(addBlankSlide(pres), {
      from: { x: inches(1), y: inches(1) },
      to: { x: inches(4), y: inches(2) },
    });
    const xml = await authoredXml(pres);
    expect(xml).toContain('<a:lnRef');
  });

  skipIfNoXmllint('animating two shapes merges into one valid timing tree', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'A',
    });
    const b = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(4),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'B',
    });
    setShapeAnimation(a, { effect: 'fadeIn' });
    setShapeAnimation(b, { effect: 'fadeOut' });
    expect(getSlideShapes(slide).length).toBe(2);
    expect(getSlides(pres).length).toBe(1);
    const xml = await authoredXml(pres);
    // The merge must keep BOTH shapes' effects — a replace-on-write
    // implementation would drop shape A's entrance and still be schema-valid,
    // so assert both build entries and both preset classes survive.
    expect((xml.match(/<p:bldP/g) ?? []).length).toBe(2);
    expect(xml).toContain('presetClass="entr"');
    expect(xml).toContain('presetClass="exit"');
  });
});
