// `RenderSlideOptions.eastAsianFamily` names the face for a run's East Asian
// characters in svg mode. resvg's fallback for glyphs the run's family lacks
// ignores the weight, so without it bold Japanese text rasterized thin.

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
  setShapeAlignment,
  setShapeTextFormat,
  setTableCellText,
  setTableCellTextFormat,
} from '@office-kit/pptx';
import {
  defaultMeasurer,
  renderSlideToSvg,
  type FontSpec,
  type TextMeasurer,
} from '../packages/preview/src/index.ts';
import { renderSlideToImage } from '../packages/preview/src/node.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const EAST_ASIAN = 'Test Sans JP';

const deck = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  return { pres, slide };
};

const boldBox = (slide: Parameters<typeof addSlideTextBox>[0], text: string) => {
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(8),
    h: inches(1),
    text,
  });
  setShapeTextFormat(shape, { bold: true });
  return shape;
};

const tspans = (svg: string) =>
  [...svg.matchAll(/<tspan([^>]*)>([^<]*)<\/tspan>/g)].map((m) => ({
    attrs: m[1]!,
    text: m[2]!,
    family: /font-family="([^"]*)"/.exec(m[1]!)?.[1],
  }));

const recordingMeasurer = (): { measure: TextMeasurer; calls: Array<[string, string]> } => {
  const calls: Array<[string, string]> = [];
  return {
    calls,
    measure: (text: string, spec: FontSpec) => {
      calls.push([text, spec.family]);
      return defaultMeasurer(text, spec);
    },
  };
};

describe('RenderSlideOptions.eastAsianFamily', () => {
  it('writes East Asian characters, punctuation included, in the East Asian family', async () => {
    const { pres, slide } = await deck();
    boldBox(slide, '着手 A1・A4、次に（全角）');

    const spans = tspans(
      renderSlideToSvg(pres, slide, { textLayout: 'svg', eastAsianFamily: EAST_ASIAN }),
    );

    expect(spans.map((s) => [s.text, s.family])).toEqual([
      ['着手', EAST_ASIAN],
      [' A1', 'Carlito'],
      ['・', EAST_ASIAN],
      ['A4', 'Carlito'],
      ['、次に（全角）', EAST_ASIAN],
    ]);
    // The run's weight carries over to the East Asian face.
    expect(spans.every((s) => s.attrs.includes('font-weight="700"'))).toBe(true);
  });

  it('starts a new text chunk at each family change, from the line start', async () => {
    const { pres, slide } = await deck();
    boldBox(slide, '着手 A1');

    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg', eastAsianFamily: EAST_ASIAN });
    const spans = tspans(svg);

    // resvg paints a chunk that mixes families with one face, so the Latin run
    // has its own absolute x; the first run starts at the <text> x.
    expect(spans[0]!.attrs).not.toMatch(/\bx="/);
    expect(spans[1]!.attrs).toMatch(/\bx="[\d.]+"/);
    expect(svg).toMatch(/<text [^>]*text-anchor="start"[^>]*><tspan [^>]*>着手<\/tspan>/);
  });

  it('keeps a centered mixed line centered once it is split into chunks', async () => {
    const { pres, slide } = await deck();
    const shape = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(8),
      h: inches(1),
      text: '着手 A1',
    });
    setShapeTextFormat(shape, { bold: true });
    setShapeAlignment(shape, 'center');

    const centered = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const split = renderSlideToSvg(pres, slide, {
      textLayout: 'svg',
      eastAsianFamily: EAST_ASIAN,
    });
    const centerX = Number(/<text x="([\d.]+)"[^>]*text-anchor="middle"/.exec(centered)![1]);
    const startX = Number(/<text x="([\d.]+)"[^>]*text-anchor="start"/.exec(split)![1]);
    const latinX = Number(/<tspan x="([\d.]+)"/.exec(split)![1]);

    // defaultMeasurer: two CJK glyphs at 1 em, " A1" at 0.55 em, 18 pt.
    const em = 18 * (96 / 72);
    const width = 2 * em + 3 * 0.55 * em;
    expect(startX).toBeCloseTo(centerX - width / 2, 1);
    expect(latinX).toBeCloseTo(startX + 2 * em, 0);
  });

  it('measures the East Asian characters with the East Asian family', async () => {
    const { pres, slide } = await deck();
    boldBox(slide, '着手 A1');
    const { measure, calls } = recordingMeasurer();

    renderSlideToSvg(pres, slide, {
      textLayout: 'svg',
      measureText: measure,
      eastAsianFamily: EAST_ASIAN,
    });

    expect(calls).toContainEqual(['着', EAST_ASIAN]);
    expect(calls).toContainEqual(['A1', 'Carlito']);
    expect(calls.some(([text, family]) => text.includes('着') && family !== EAST_ASIAN)).toBe(
      false,
    );
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
    const cell = getTableCell(table, 0, 0);
    setTableCellText(cell, '合計 12');
    setTableCellTextFormat(cell, { bold: true });

    const spans = tspans(
      renderSlideToSvg(pres, slide, { textLayout: 'svg', eastAsianFamily: EAST_ASIAN }),
    );

    expect(spans.find((s) => s.text === '合計')?.family).toBe(EAST_ASIAN);
  });

  it('leaves the svg unchanged when unset, and is ignored by the foreignObject path', async () => {
    const { pres, slide } = await deck();
    boldBox(slide, '着手 A1');

    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    expect(tspans(svg).map((s) => [s.text, s.family])).toEqual([['着手 A1', 'Carlito']]);
    expect(svg).not.toMatch(/<tspan x=/);

    expect(renderSlideToSvg(pres, slide, { eastAsianFamily: EAST_ASIAN })).toBe(
      renderSlideToSvg(pres, slide),
    );
  });
});

describe('renderSlideToImage East Asian face', () => {
  it('skips a fontFiles entry that is not a font, as resvg does', async () => {
    const { pres, slide } = await deck();
    boldBox(slide, '着手 A1');

    expect(() => renderSlideToImage(pres, slide, { fontFiles: [fixturePath] })).not.toThrow();
  });
});
