// Chart titles leave size and rotation to the application unless authored;
// `valueAxisTickLabelPos` mirrors the category-axis field on the value axis.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  type ChartSpec,
  addSlideChart,
  getSlideCharts,
  getSlides,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decoder = new TextDecoder();
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const roundTrip = async (
  spec: ChartSpec,
): Promise<{ readonly spec: ChartSpec; readonly xml: string }> => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  addSlideChart(getSlides(pres)[0]!, {
    x: inches(0.5),
    y: inches(0.5),
    w: inches(8),
    h: inches(4.5),
    spec,
  });
  const reloaded = await loadPresentation(await savePresentation(pres));
  const xml = decoder.decode(readPackagePart(reloaded, '/ppt/charts/chart1.xml')!);
  return { spec: getSlideCharts(getSlides(reloaded)[0]!)[0]!.spec!, xml };
};

const sliceElement = (xml: string, tag: string): string =>
  xml.slice(xml.indexOf(`<c:${tag}>`), xml.indexOf(`</c:${tag}>`));

const titleIn = (xml: string, tag: string): string => sliceElement(sliceElement(xml, tag), 'title');

const base: ChartSpec = {
  kind: 'column',
  categories: ['A', 'B'],
  series: [{ name: 'Count', values: [1, 2] }],
};

describe('fn API: chart title defaults', () => {
  skipIfNoXmllint('leaves the title size to the application when sizePt is unset', async () => {
    const { spec, xml } = await roundTrip({ ...base, title: 'T', titleStyle: { bold: true } });
    const title = sliceElement(xml, 'title');
    expect(title).toContain('<a:defRPr/>');
    expect(title).not.toContain('sz=');
    expect(title).not.toContain('rot=');
    expect(spec.titleStyle).toMatchObject({ bold: true });
    expect(spec.titleStyle?.sizePt).toBeUndefined();
    expectSchemaValid(xml, 'chart');
  });

  it('writes the authored size on the run only', async () => {
    const { spec, xml } = await roundTrip({ ...base, title: 'T', titleStyle: { sizePt: 20 } });
    const title = sliceElement(xml, 'title');
    expect(title).toContain('<a:defRPr/>');
    expect(title).toContain('<a:rPr lang="en-US" sz="2000"');
    expect(spec.titleStyle?.sizePt).toBe(20);
  });
});

describe('fn API: axis title rotation', () => {
  skipIfNoXmllint('omits rot on a value-axis title without titleRotationDeg', async () => {
    const { spec, xml } = await roundTrip({ ...base, valueAxisTitle: 'Score' });
    // PowerPoint reads `vert="horz"` without `rot` as "not rotated", so both
    // stay out and its vertical default applies.
    expect(titleIn(xml, 'valAx')).toContain('<a:bodyPr spcFirstLastPara="1"');
    expect(titleIn(xml, 'valAx')).not.toContain('rot=');
    expect(titleIn(xml, 'valAx')).not.toContain('vert=');
    expect(spec.valueAxisTitle).toBe('Score');
    expect(spec.valueAxisTitleRotationDeg).toBeUndefined();
    expectSchemaValid(xml, 'chart');
  });

  it('writes rot="0" when the value-axis title is pinned horizontal', async () => {
    const { spec, xml } = await roundTrip({
      ...base,
      valueAxisTitle: 'Score',
      valueAxisTitleRotationDeg: 0,
    });
    expect(titleIn(xml, 'valAx')).toContain(
      '<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz"',
    );
    expect(spec.valueAxisTitleRotationDeg).toBe(0);
  });

  it('omits rot on a category-axis title without titleRotationDeg', async () => {
    const { spec, xml } = await roundTrip({ ...base, categoryAxisTitle: 'Group' });
    expect(titleIn(xml, 'catAx')).not.toContain('rot=');
    expect(titleIn(xml, 'catAx')).not.toContain('vert=');
    expect(spec.categoryAxisTitleRotationDeg).toBeUndefined();
  });

  it('omits rot on a secondary value-axis title', async () => {
    const { xml } = await roundTrip({
      ...base,
      series: [...base.series, { name: 'Share', values: [3, 4], secondaryAxis: true }],
      secondaryValueAxis: { title: 'Share' },
    });
    const secondary = xml.slice(xml.lastIndexOf('<c:valAx>'));
    expect(secondary).toContain('<a:t>Share</a:t>');
    expect(sliceElement(secondary, 'title')).not.toContain('rot=');
    expect(sliceElement(secondary, 'title')).not.toContain('vert=');
  });
});

describe('fn API: valueAxisTickLabelPos', () => {
  skipIfNoXmllint('writes tickLblPos on the primary value axis and reads it back', async () => {
    const { spec, xml } = await roundTrip({
      ...base,
      valueAxisMinorTickMark: 'none',
      valueAxisTickLabelPos: 'low',
      valueAxisLineColor: '#888888',
    });
    const valAx = sliceElement(xml, 'valAx');
    expect(valAx).toContain('<c:minorTickMark val="none"/><c:tickLblPos val="low"/><c:spPr>');
    expect(spec.valueAxisTickLabelPos).toBe('low');
    expectSchemaValid(xml, 'chart');
  });

  it('writes nothing when unset', async () => {
    const { spec, xml } = await roundTrip(base);
    expect(sliceElement(xml, 'valAx')).not.toContain('tickLblPos');
    expect(spec.valueAxisTickLabelPos).toBeUndefined();
  });
});
