// Direct API-contract tests for the readers the preview workstreams added.
// The renderer test suites exercise these end-to-end; this file pins down the
// readers' own contracts — return shapes, null cases, and cascade behavior —
// against the public API alone.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addSlide,
  addSlideTable,
  addSlideTextBox,
  findSlideLayout,
  getParagraphBulletImageBytes,
  getParagraphPropertiesEffective,
  getShapeGradientFillEffective,
  getSlides,
  getSlideTables,
  getTableCell,
  getTableCellParagraphs,
  inches,
  loadPresentation,
  savePresentation,
  setShapeGradientFill,
  setTableCellAlignment,
  setTableCellTextFormat,
  setTableCellText,
  findSlidePlaceholder,
  setShapeText,
} from '../src/api/index.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const blankSlide = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  return { pres, slide: addSlide(pres, { layout }) };
};

describe('getTableCellParagraphs', () => {
  it('returns alignment and per-run format for a styled cell', async () => {
    const { slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      rows: [['styled', 'plain']],
    });
    const styled = getTableCell(table, 0, 0);
    if (!styled) throw new Error('cell (0,0) missing');
    setTableCellTextFormat(styled, { size: 28, bold: true, color: '#FF0000' });
    setTableCellAlignment(styled, 'center');

    const paras = getTableCellParagraphs(styled);
    expect(paras).toHaveLength(1);
    expect(paras[0]!.align).toBe('center');
    const run = paras[0]!.elements.find((e) => e.kind === 'r');
    if (!run || run.kind !== 'r') throw new Error('run missing');
    expect(run.text).toBe('styled');
    expect(run.format?.size).toBe(28);
    expect(run.format?.bold).toBe(true);

    // The sibling cell carries no explicit format: alignment unset, format
    // fields left undefined for the caller's cascade to resolve.
    const plain = getTableCell(table, 0, 1);
    if (!plain) throw new Error('cell (0,1) missing');
    const plainParas = getTableCellParagraphs(plain);
    expect(plainParas[0]!.align).toBeNull();
    const plainRun = plainParas[0]!.elements.find((e) => e.kind === 'r');
    if (!plainRun || plainRun.kind !== 'r') throw new Error('plain run missing');
    expect(plainRun.format?.size).toBeUndefined();
  });

  it('multi-paragraph cell text comes back one entry per paragraph', async () => {
    const { slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(1),
      rows: [['a']],
    });
    const cell = getTableCell(table, 0, 0);
    if (!cell) throw new Error('cell missing');
    setTableCellText(cell, 'first\nsecond');
    expect(getTableCellParagraphs(cell)).toHaveLength(2);
  });
});

describe('getParagraphBulletImageBytes', () => {
  it('returns null for a paragraph without a picture bullet', async () => {
    const { slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(1),
      text: 'no picture bullet here',
    });
    expect(getParagraphBulletImageBytes(box, 0)).toBeNull();
  });
});

describe('getShapeGradientFillEffective', () => {
  it("returns the shape's own gradient verbatim", async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'g',
    });
    setShapeGradientFill(box, {
      stops: [
        { offset: 0, color: '#112233' },
        { offset: 1, color: '#445566' },
      ],
      angleDeg: 45,
    });
    const g = getShapeGradientFillEffective(pres, box);
    expect(g).not.toBeNull();
    expect(g!.stops.map((s) => s.color)).toEqual(['#112233', '#445566']);
    expect(g!.angleDeg).toBe(45);
  });

  it('returns null for a non-placeholder shape with no gradient anywhere', async () => {
    const { pres, slide } = await blankSlide();
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'solid',
    });
    expect(getShapeGradientFillEffective(pres, box)).toBeNull();
  });

  it.each([
    ['noFill', '<a:noFill/>'],
    ['solidFill', '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>'],
  ])('does not inherit a placeholder gradient through an explicit %s', async (_kind, fill) => {
    const source = unzipSync(
      await readFile(
        fileURLToPath(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
      ),
    );
    const gradient =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill>';
    const slideXml = strFromU8(source['ppt/slides/slide1.xml']!);
    source['ppt/slides/slide1.xml'] = strToU8(
      slideXml.replace('<p:spPr/>', `<p:spPr>${fill}</p:spPr>`),
    );
    const layoutXml = strFromU8(source['ppt/slideLayouts/slideLayout6.xml']!);
    source['ppt/slideLayouts/slideLayout6.xml'] = strToU8(
      layoutXml.replace('<p:spPr/>', `<p:spPr>${gradient}</p:spPr>`),
    );
    const pres = await loadPresentation(zipSync(source));
    const shape = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
    if (!shape) throw new Error('expected title placeholder');

    expect(getShapeGradientFillEffective(pres, shape)).toBeNull();
  });

  it('inherits a literal layout gradient when the shape has no fill choice', async () => {
    const source = unzipSync(
      await readFile(
        fileURLToPath(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
      ),
    );
    const gradient =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill>';
    const layoutXml = strFromU8(source['ppt/slideLayouts/slideLayout6.xml']!);
    source['ppt/slideLayouts/slideLayout6.xml'] = strToU8(
      layoutXml.replace('<p:spPr/>', `<p:spPr>${gradient}</p:spPr>`),
    );
    const pres = await loadPresentation(zipSync(source));
    const shape = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
    if (!shape) throw new Error('expected title placeholder');

    expect(getShapeGradientFillEffective(pres, shape)?.stops.map((stop) => stop.color)).toEqual([
      '#FF0000',
      '#0000FF',
    ]);
  });

  it('stops at an explicit layout fill instead of inheriting a master gradient', async () => {
    const source = unzipSync(
      await readFile(
        fileURLToPath(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
      ),
    );
    const gradient =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill>';
    const layoutXml = strFromU8(source['ppt/slideLayouts/slideLayout6.xml']!);
    source['ppt/slideLayouts/slideLayout6.xml'] = strToU8(
      layoutXml.replace(
        '<p:spPr/>',
        '<p:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></p:spPr>',
      ),
    );
    const masterXml = strFromU8(source['ppt/slideMasters/slideMaster1.xml']!);
    source['ppt/slideMasters/slideMaster1.xml'] = strToU8(
      masterXml.replace('<p:spPr>', `<p:spPr>${gradient}`),
    );
    const pres = await loadPresentation(zipSync(source));
    const shape = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
    if (!shape) throw new Error('expected title placeholder');

    expect(getShapeGradientFillEffective(pres, shape)).toBeNull();
  });

  it('inherits a literal master gradient when shape and layout have no fill choice', async () => {
    const source = unzipSync(
      await readFile(
        fileURLToPath(new URL('./fixtures/minimal/one-text-slide.pptx', import.meta.url)),
      ),
    );
    const gradient =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill>';
    const masterXml = strFromU8(source['ppt/slideMasters/slideMaster1.xml']!);
    source['ppt/slideMasters/slideMaster1.xml'] = strToU8(
      masterXml.replace('<p:spPr>', `<p:spPr>${gradient}`),
    );
    const pres = await loadPresentation(zipSync(source));
    const shape = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
    if (!shape) throw new Error('expected title placeholder');

    expect(getShapeGradientFillEffective(pres, shape)?.stops.map((stop) => stop.color)).toEqual([
      '#FF0000',
      '#0000FF',
    ]);
  });
});

describe('getParagraphPropertiesEffective bullet cascade', () => {
  it('body placeholders inherit the master bodyStyle bullet; titles do not', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Title and Content');
    if (!layout) throw new Error('Title and Content layout not found');
    const slide = addSlide(pres, { layout });
    const body = findSlidePlaceholder(slide, 'body');
    const title = findSlidePlaceholder(slide, 'title');
    if (!body || !title) throw new Error('placeholders missing');
    setShapeText(body, 'bulleted');
    setShapeText(title, 'no bullet');

    // python-pptx's master authors bodyStyle lvl1 buChar="•" (normalised to
    // the 'bullet' token) and an explicit <a:buNone> in titleStyle.
    const bodyProperties = getParagraphPropertiesEffective(pres, body, 0);
    expect(bodyProperties.bullet).toBe('bullet');
    // This is read from the imported master bodyStyle lvl1pPr, rather than
    // authored on the slide paragraph itself.
    expect(bodyProperties.bulletDetail).toMatchObject({
      font: 'Arial',
      fontFollowText: false,
      color: null,
      sizePct: null,
      sizePts: null,
    });
    expect(getParagraphPropertiesEffective(pres, title, 0).bullet).toBe('none');
  });

  it('round-trips through save/load', async () => {
    const pres = await loadPresentation(await readFile(fixturePath));
    const layout = findSlideLayout(pres, 'Title and Content');
    if (!layout) throw new Error('layout missing');
    const slide = addSlide(pres, { layout });
    const body = findSlidePlaceholder(slide, 'body');
    if (!body) throw new Error('body missing');
    setShapeText(body, 'persisted');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const rSlide = getSlides(reloaded).at(-1)!;
    const rBody = findSlidePlaceholder(rSlide, 'body');
    if (!rBody) throw new Error('reloaded body missing');
    expect(getParagraphPropertiesEffective(reloaded, rBody, 0).bullet).toBe('bullet');
  });
});

describe('table reader sanity on a reloaded deck', () => {
  it('getSlideTables + getTableCellParagraphs survive a round-trip', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      rows: [['x', 'y']],
    });
    const cell = getTableCell(table, 0, 0);
    if (!cell) throw new Error('cell missing');
    setTableCellTextFormat(cell, { size: 20 });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const rSlide = getSlides(reloaded).at(-1)!;
    const rTable = getSlideTables(rSlide)[0]!;
    const rCell = getTableCell(rTable, 0, 0);
    if (!rCell) throw new Error('reloaded cell missing');
    const run = getTableCellParagraphs(rCell)[0]!.elements[0]!;
    if (run.kind !== 'r') throw new Error('expected run');
    expect(run.format?.size).toBe(20);
  });
});
