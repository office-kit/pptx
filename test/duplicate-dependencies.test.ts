import { expect, it } from 'vitest';
import {
  _internalPackageOf,
  createPresentation,
  addBlankSlide,
  addSlideChart,
  duplicateSlide,
  copyShape,
  importShape,
  getSlideCharts,
  getShapeChartSeriesValues,
  setChartSpec,
  getSlidePartName,
  getSlideNotes,
  setSlideNotes,
  savePresentation,
  loadPresentation,
  getSlides,
  inches,
  validatePresentation,
} from '../src/api/index.ts';
import { partName, resolveTarget } from '../src/internal/opc/index.ts';

it('duplicates charts, workbooks and notes without sharing their editable state', async () => {
  const pres = createPresentation();
  const original = addBlankSlide(pres);
  addSlideChart(original, {
    x: inches(1),
    y: inches(1),
    w: inches(5),
    h: inches(3),
    spec: { kind: 'column', categories: ['Q1'], series: [{ name: 'Revenue', values: [10] }] },
  });
  setSlideNotes(original, 'Original notes');
  const duplicate = duplicateSlide(pres, original);
  setChartSpec(getSlideCharts(duplicate)[0]!, {
    kind: 'column',
    categories: ['Q1'],
    series: [{ name: 'Revenue', values: [99] }],
  });
  setSlideNotes(duplicate, 'New notes');
  const reloaded = await loadPresentation(await savePresentation(pres));
  const [first, second] = getSlides(reloaded);
  expect(getShapeChartSeriesValues(getSlideCharts(first!)[0]!.shape, 'Revenue')).toEqual([10]);
  expect(getShapeChartSeriesValues(getSlideCharts(second!)[0]!.shape, 'Revenue')).toEqual([99]);
  expect(getSlideNotes(first!)).toBe('Original notes');
  expect(getSlideNotes(second!)).toBe('New notes');
  expect(validatePresentation(reloaded).filter((issue) => issue.severity === 'error')).toEqual([]);
});

it('copies unknown dependency bytes and cycles while keeping relationship IDs', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const pkg = _internalPackageOf(pres);
  const source = partName(getSlidePartName(slide));
  const unknown = partName('/custom/unknown.bin');
  const bytes = new Uint8Array([0, 17, 255, 9]);
  pkg.addPart(unknown, 'application/x-private', bytes);
  const rels = pkg.getRels(source)!;
  rels.items.push({
    id: 'rId99',
    type: 'urn:private:extension',
    target: unknown,
    targetMode: 'Internal',
  });
  pkg.setRels(source, rels);
  pkg.setRels(unknown, {
    items: [{ id: 'rId1', type: 'urn:private:back', target: source, targetMode: 'Internal' }],
  });
  const copy = partName(getSlidePartName(duplicateSlide(pres, slide)));
  const copiedRel = pkg.getRels(copy)!.items.find((rel) => rel.id === 'rId99')!;
  const copiedPart = resolveTarget(copy, copiedRel.target);
  expect(copiedPart).not.toBe(unknown);
  expect(pkg.getPart(copiedPart)!.data).toEqual(bytes);
  expect(resolveTarget(copiedPart, pkg.getRels(copiedPart)!.items[0]!.target)).toBe(copy);
});

it('fails atomically when an owned dependency is missing', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const pkg = _internalPackageOf(pres);
  const source = partName(getSlidePartName(slide));
  const rels = pkg.getRels(source)!;
  rels.items.push({
    id: 'rId99',
    type: 'urn:private:extension',
    target: '/missing.bin',
    targetMode: 'Internal',
  });
  pkg.setRels(source, rels);
  const before = pkg.parts.map((part) => ({ name: part.name, data: new Uint8Array(part.data) }));
  expect(() => duplicateSlide(pres, slide)).toThrow('missing dependency');
  expect(pkg.parts.map(({ name, data }) => ({ name, data }))).toEqual(before);
  expect(getSlides(pres)).toHaveLength(1);
});

it('copies charts with independent editable data on the same and another slide', async () => {
  const pres = createPresentation();
  const source = addBlankSlide(pres);
  const destination = addBlankSlide(pres);
  addSlideChart(source, {
    x: inches(1),
    y: inches(1),
    w: inches(5),
    h: inches(3),
    spec: { kind: 'column', categories: ['Q1'], series: [{ name: 'Revenue', values: [10] }] },
  });
  const original = getSlideCharts(source)[0]!.shape;
  copyShape(source, original);
  copyShape(destination, original);
  setChartSpec(getSlideCharts(source)[1]!, {
    kind: 'column',
    categories: ['Q1'],
    series: [{ name: 'Revenue', values: [20] }],
  });
  setChartSpec(getSlideCharts(destination)[0]!, {
    kind: 'column',
    categories: ['Q1'],
    series: [{ name: 'Revenue', values: [30] }],
  });
  const reloaded = await loadPresentation(await savePresentation(pres));
  const values = getSlides(reloaded).flatMap((slide) =>
    getSlideCharts(slide).map((chart) => getShapeChartSeriesValues(chart.shape, 'Revenue')),
  );
  expect(values).toEqual([[10], [20], [30]]);
  expect(validatePresentation(reloaded).filter((issue) => issue.severity === 'error')).toEqual([]);
});

it('imports chart dependency graphs from a frozen copy without sharing later edits', async () => {
  const original = createPresentation();
  const originalSlide = addBlankSlide(original);
  addSlideChart(originalSlide, {
    x: inches(1),
    y: inches(1),
    w: inches(5),
    h: inches(3),
    spec: { kind: 'column', categories: ['Q1'], series: [{ name: 'Revenue', values: [10] }] },
  });
  const frozen = await loadPresentation(await savePresentation(original));
  setChartSpec(getSlideCharts(originalSlide)[0]!, {
    kind: 'column',
    categories: ['Q1'],
    series: [{ name: 'Revenue', values: [40] }],
  });
  importShape(originalSlide, getSlideCharts(getSlides(frozen)[0]!)[0]!.shape);
  setChartSpec(getSlideCharts(originalSlide)[1]!, {
    kind: 'column',
    categories: ['Q1'],
    series: [{ name: 'Revenue', values: [70] }],
  });
  importShape(originalSlide, getSlideCharts(getSlides(frozen)[0]!)[0]!.shape);
  const reloaded = await loadPresentation(await savePresentation(original));
  expect(
    getSlideCharts(getSlides(reloaded)[0]!).map((chart) =>
      getShapeChartSeriesValues(chart.shape, 'Revenue'),
    ),
  ).toEqual([[40], [70], [10]]);
  expect(
    getShapeChartSeriesValues(getSlideCharts(getSlides(frozen)[0]!)[0]!.shape, 'Revenue'),
  ).toEqual([10]);
  expect(validatePresentation(reloaded).filter((issue) => issue.severity === 'error')).toEqual([]);
});
