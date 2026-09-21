// copyShape — clone a shape onto another slide in the same deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  _internalPackageOf,
  getSlidePartName,
  addSlideChart,
  getSlideCharts,
  getShapeChartSeriesValues,
  setChartSpec,
  validatePresentation,
  addSlideShape,
  copyShape,
  findShapesByKind,
  getShapeKind,
  getShapeImageBytes,
  groupShapes,
  savePresentation,
  readPackagePart,
  getShapeText,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

import { partName, resolveTarget } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const tinyPng = (): Uint8Array =>
  new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);

describe('fn API: copyShape', () => {
  it('copies a text shape onto another slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    const sp = addSlideShape(slideA!, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'COPY ME',
      name: 'Tagged',
    });
    const beforeBCount = getSlideShapes(slideB!).length;

    const copied = copyShape(slideB!, sp);
    expect(getShapeKind(copied)).toBe('shape');
    expect(getShapeText(copied)).toBe('COPY ME');
    expect(getSlideShapes(slideB!).length).toBe(beforeBCount + 1);
  });

  it('copies a picture shape and the new shape still references valid media', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    const pic = addSlideImage(slideA!, tinyPng(), {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      format: 'png',
    });
    const copied = copyShape(slideB!, pic);
    expect(getShapeKind(copied)).toBe('picture');
    // The target slide now has at least one picture.
    expect(findShapesByKind(slideB!, 'picture').length).toBeGreaterThan(0);
  });

  it('copies across presentations with independent media, surviving a save and reload', async () => {
    const presA = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const presB = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const pic = addSlideImage(getSlides(presA)[0]!, tinyPng(), {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      format: 'png',
    });
    const copied = copyShape(getSlides(presB)[0]!, pic);
    expect(getShapeImageBytes(copied)).toEqual(tinyPng());
    const restored = await loadPresentation(await savePresentation(presB));
    const pictures = findShapesByKind(getSlides(restored)[0]!, 'picture');
    expect(getShapeImageBytes(pictures.at(-1)!)).toEqual(tinyPng());
  });

  it('copies chart dependencies without sharing the workbook or overwriting target parts', async () => {
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const target = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const chart = (pres: typeof source, value: number) =>
      addSlideChart(getSlides(pres)[0]!, {
        x: inches(1),
        y: inches(1),
        w: inches(5),
        h: inches(3),
        spec: {
          kind: 'column',
          categories: ['Q1'],
          series: [{ name: 'Revenue', values: [value] }],
        },
      });
    chart(target, 50);
    copyShape(getSlides(target)[0]!, chart(source, 10));
    setChartSpec(getSlideCharts(getSlides(source)[0]!)[0]!, {
      kind: 'column',
      categories: ['Q1'],
      series: [{ name: 'Revenue', values: [99] }],
    });
    const restored = await loadPresentation(await savePresentation(target));
    expect(
      getSlideCharts(getSlides(restored)[0]!).map((chart) =>
        getShapeChartSeriesValues(chart.shape, 'Revenue'),
      ),
    ).toEqual([[50], [10]]);
    expect(validatePresentation(restored).filter((issue) => issue.severity === 'error')).toEqual(
      [],
    );
  });

  it('retains unknown dependency bodies and cycles and rejects missing parts before mutation', async () => {
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const target = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(source)[0]!;
    const targetSlide = getSlides(target)[0]!;
    const pic = addSlideImage(slide, tinyPng(), {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      format: 'png',
    });
    const pkg = _internalPackageOf(source);
    const slidePart = partName(getSlidePartName(slide));
    const imageRel = pkg.getRels(slidePart)!.items.find((rel) => rel.type.endsWith('/image'))!;
    const imagePart = resolveTarget(slidePart, imageRel.target);
    const unknown = partName('/custom/data.bin');
    const bytes = new Uint8Array([0, 255, 17]);
    pkg.addPart(unknown, 'application/x-private', bytes);
    pkg.setRels(imagePart, {
      items: [{ id: 'rId99', type: 'urn:custom', target: unknown, targetMode: 'Internal' }],
    });
    pkg.setRels(unknown, {
      items: [{ id: 'rId1', type: 'urn:back', target: slidePart, targetMode: 'Internal' }],
    });
    copyShape(targetSlide, pic);
    const targetPkg = _internalPackageOf(target);
    const copiedPart = targetPkg.parts.find(
      (part) => part.contentType === 'application/x-private',
    )!;
    expect(copiedPart.data).toEqual(bytes);
    expect(
      resolveTarget(copiedPart.name, targetPkg.getRels(copiedPart.name)!.items[0]!.target),
    ).toBe(getSlidePartName(targetSlide));
    pkg.removePart(unknown);
    const before = targetPkg.parts.map((part) => ({
      name: part.name,
      data: new Uint8Array(part.data),
    }));
    expect(() => copyShape(targetSlide, pic)).toThrow('missing dependency');
    expect(targetPkg.parts.map(({ name, data }) => ({ name, data }))).toEqual(before);
  });

  it('allocates distinct ids for every nested shape when copying groups repeatedly', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const box = (text: string) =>
      addSlideShape(slide, {
        preset: 'rect',
        x: inches(1),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text,
      });
    const group = groupShapes([box('A'), box('B')]);
    copyShape(slide, group);
    copyShape(slide, group);
    box('After copies');
    const xml = new TextDecoder().decode(readPackagePart(pres, '/ppt/slides/slide1.xml')!);
    const ids = [...xml.matchAll(/<p:cNvPr[^>]* id="(\d+)"/g)].map((m) => m[1]);
    expect(ids.length).toBeGreaterThan(8);
    expect(new Set(ids).size).toBe(ids.length);
  });
});
