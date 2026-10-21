// `RenderSlideOptions.resolveFamily` lets an svg-mode host keep font families
// it can supply itself (e.g. through embedded `@font-face` rules) instead of
// the bundled substitutes `substituteFamily` maps everything else to.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTable,
  addSlideTextBox,
  findSlideLayout,
  getTableCell,
  inches,
  loadPresentation,
  setShapeTextAutoFit,
  setShapeTextFormat,
  setTableCellTextFormat,
} from '@office-kit/pptx';
import {
  auditTextLayout,
  defaultMeasurer,
  renderSlideToSvg,
  substituteFamily,
  type FontSpec,
  type TextMeasurer,
} from '../packages/preview/src/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const CUSTOM = 'Lora';
// Far more than a 2×1 in box holds, so a normAutofit body runs the shrink search.
const LINES = Array.from({ length: 8 }, (_, i) => `Line ${i + 1}`).join('\n');
const keepCustom = (family: string | null): string =>
  family === CUSTOM ? family : substituteFamily(family);

const deck = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  return { pres, slide };
};

const textFamilies = (svg: string): string[] => [
  ...new Set([...svg.matchAll(/<(?:text|tspan)[^>]*font-family="([^"]*)"/g)].map((m) => m[1]!)),
];

const recordingMeasurer = (): { measure: TextMeasurer; families: Set<string> } => {
  const families = new Set<string>();
  return {
    families,
    measure: (text: string, spec: FontSpec) => {
      families.add(spec.family);
      return defaultMeasurer(text, spec);
    },
  };
};

describe('RenderSlideOptions.resolveFamily', () => {
  it('substitutes unknown families by default in svg mode', async () => {
    const { pres, slide } = await deck();
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(1),
      text: 'Hello',
    });
    setShapeTextFormat(shape, { font: CUSTOM });

    expect(textFamilies(renderSlideToSvg(pres, slide, { textLayout: 'svg' }))).toEqual(['Carlito']);
  });

  it('writes and measures the family the host keeps', async () => {
    const { pres, slide } = await deck();
    const custom = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(1),
      text: 'Hello',
    });
    setShapeTextFormat(custom, { font: CUSTOM });
    const arial = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(3),
      w: inches(5),
      h: inches(1),
      text: 'World',
    });
    setShapeTextFormat(arial, { font: 'Arial' });
    const { measure, families } = recordingMeasurer();

    const svg = renderSlideToSvg(pres, slide, {
      textLayout: 'svg',
      measureText: measure,
      resolveFamily: keepCustom,
    });

    expect(textFamilies(svg).sort()).toEqual(['Liberation Sans', CUSTOM]);
    expect([...families].sort()).toEqual(['Liberation Sans', CUSTOM]);
  });

  it('applies to table cell text', async () => {
    const { pres, slide } = await deck();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      rows: [['a', 'b']],
    });
    setTableCellTextFormat(getTableCell(table, 0, 0), { font: CUSTOM });

    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg', resolveFamily: keepCustom });

    expect(textFamilies(svg)).toContain(CUSTOM);
  });

  it('measures the shrink-to-fit search with the resolved family', async () => {
    const { pres, slide } = await deck();
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: LINES,
    });
    setShapeTextFormat(shape, { font: CUSTOM });
    setShapeTextAutoFit(shape, 'normal');
    const { measure, families } = recordingMeasurer();

    renderSlideToSvg(pres, slide, {
      textLayout: 'svg',
      measureText: measure,
      resolveFamily: () => CUSTOM,
    });

    expect([...families]).toEqual([CUSTOM]);
  });

  it('quotes a kept family that is not a plain CSS identifier sequence', async () => {
    const { pres, slide } = await deck();
    addSlideTextBox(slide, { x: inches(1), y: inches(1), w: inches(5), h: inches(1), text: 'x' });
    const { measure, families } = recordingMeasurer();

    const svg = renderSlideToSvg(pres, slide, {
      textLayout: 'svg',
      measureText: measure,
      resolveFamily: () => 'M PLUS 1p',
    });

    // Unquoted, "1p" is not an identifier and the browser drops the whole value.
    expect(svg).toContain('font-family="&quot;M PLUS 1p&quot;"');
    expect([...families]).toEqual(['M PLUS 1p']);
  });

  it('is ignored by the foreignObject path, which keeps the authored name itself', async () => {
    const { pres, slide } = await deck();
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: LINES,
    });
    setShapeTextFormat(shape, { font: CUSTOM });
    // The shrink-to-fit search measures with the SVG engine in this mode too.
    setShapeTextAutoFit(shape, 'normal');
    const { measure, families } = recordingMeasurer();

    const plain = renderSlideToSvg(pres, slide, { measureText: measure });
    const withOption = renderSlideToSvg(pres, slide, {
      measureText: measure,
      resolveFamily: () => 'Other',
    });

    expect(withOption).toBe(plain);
    expect(families.has('Other')).toBe(false);
    expect(plain).toContain(`font-family:&quot;${CUSTOM}&quot;, Calibri`);
  });
});

describe('auditTextLayout', () => {
  it('fits a normAutofit body with the authored family, as it lays it out', async () => {
    const { pres, slide } = await deck();
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: LINES,
    });
    setShapeTextFormat(shape, { font: CUSTOM });
    setShapeTextAutoFit(shape, 'normal');
    const { measure, families } = recordingMeasurer();

    auditTextLayout(pres, { measureText: measure });

    expect([...families]).toEqual([CUSTOM]);
  });
});
