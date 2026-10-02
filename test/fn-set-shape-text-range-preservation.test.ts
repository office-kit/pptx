import { expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  createPresentation,
  getShapeXmlString,
  getShapeText,
  getShapeRunHyperlink,
  getTableCell,
  getSlideTables,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeRunHyperlink,
  setShapeParagraphs,
  setShapeText,
  setTableCellText,
} from '../src/api/index.ts';

it('preserves run properties and extension children through a saved range replacement', async () => {
  const presentation = createPresentation();
  const slide = addBlankSlide(presentation);
  const textShape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(2),
    text: 'hello',
  });
  setShapeRunHyperlink(textShape, 0, 0, 'https://example.com');
  const parts = unzipSync(await savePresentation(presentation));
  const slidePath = 'ppt/slides/slide1.xml';
  parts[slidePath] = strToU8(
    strFromU8(parts[slidePath]!).replace(
      '<a:r>',
      '<a:r xmlns:x="urn:custom" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" x:custom="keep">',
    ),
  );
  parts[slidePath] = strToU8(
    strFromU8(parts[slidePath]!).replace(
      '</a:rPr>',
      '<a:extLst><a:ext uri="urn:office-kit:test"/></a:extLst></a:rPr>',
    ),
  );

  const loaded = await loadPresentation(zipSync(parts));
  const shape = getSlideShapes(getSlides(loaded)[0]!)[0]!;
  setShapeText(shape, 'a', { range: { start: 1, end: 2 } });
  const reopened = await loadPresentation(await savePresentation(loaded));
  const result = getSlideShapes(getSlides(reopened)[0]!)[0]!;
  const xml = getShapeXmlString(result);
  expect(getShapeText(result)).toBe('hallo');
  expect(xml.match(/x:custom="keep"/g)).toHaveLength(2);
  expect(xml).toContain('<a:ext uri="urn:office-kit:test"/>');
  expect(xml).toMatch(/<a:hlinkClick r:id="rId\d+"\/>/);
});

it('applies Change Case to formatted shape text and survives save/reload', async () => {
  const presentation = createPresentation();
  const slide = addBlankSlide(presentation);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(2),
    text: '',
  });
  setShapeParagraphs(shape, [
    { runs: [{ text: 'Straße', format: { bold: true } }] },
    { runs: [{ text: 'ΟΣ', format: { italic: true } }] },
  ]);
  setShapeRunHyperlink(shape, 0, 0, 'https://example.com');
  const loaded = await loadPresentation(await savePresentation(presentation));
  const loadedShape = getSlideShapes(getSlides(loaded)[0]!)[0]!;
  setShapeText(loadedShape, { case: 'upper' });
  const reopened = await loadPresentation(await savePresentation(loaded));
  const result = getSlideShapes(getSlides(reopened)[0]!)[0]!;
  const xml = getShapeXmlString(result);
  expect(getShapeText(result)).toBe('STRASSE\nΟΣ');
  expect(getShapeRunHyperlink(result, 0, 0)).toBe('https://example.com');
  expect(xml).toMatch(/<a:rPr[^>]*b="1"/);
  expect(xml).toMatch(/<a:rPr[^>]*i="1"/);
});

it('supports Unicode length changes and rejects invalid case/ranges atomically', async () => {
  const presentation = createPresentation();
  const slide = addBlankSlide(presentation);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(2),
    text: 'ß',
  });
  setShapeText(shape, { case: 'upper' });
  expect(getShapeText(shape)).toBe('SS');
  const before = getShapeXmlString(shape);
  expect(() => setShapeText(shape, { case: 'invalid' as 'upper' })).toThrow(RangeError);
  expect(() => setShapeText(shape, { case: 'lower' }, { range: { start: 0, end: 99 } })).toThrow(
    RangeError,
  );
  expect(getShapeXmlString(shape)).toBe(before);
});

it('applies Change Case to table cells through save/reload', async () => {
  const presentation = createPresentation();
  const slide = addBlankSlide(presentation);
  addSlideTable(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(2),
    rows: [['Straße']],
  });
  const loaded = await loadPresentation(await savePresentation(presentation));
  const table = getSlideTables(getSlides(loaded)[0]!)[0]!;
  setTableCellText(getTableCell(table, 0, 0), { case: 'upper' });
  const reopened = await loadPresentation(await savePresentation(loaded));
  expect(getShapeXmlString(getSlideTables(getSlides(reopened)[0]!)[0]!)).toContain('STRASSE');
});

it('preserves run metadata through a saved table-cell range replacement', async () => {
  const presentation = createPresentation();
  const slide = addBlankSlide(presentation);
  addSlideTable(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(2),
    rows: [['hello']],
  });
  const parts = unzipSync(await savePresentation(presentation));
  const slidePath = 'ppt/slides/slide1.xml';
  parts[slidePath] = strToU8(
    strFromU8(parts[slidePath]!)
      .replace(
        '<a:r>',
        '<a:r xmlns:x="urn:custom" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" x:custom="keep">',
      )
      .replace('</a:rPr>', '<a:extLst><a:ext uri="urn:office-kit:table-test"/></a:extLst></a:rPr>'),
  );
  const loaded = await loadPresentation(zipSync(parts));
  const table = getSlideTables(getSlides(loaded)[0]!)[0]!;
  setTableCellText(getTableCell(table, 0, 0), 'a', { range: { start: 1, end: 2 } });
  const reopened = await loadPresentation(await savePresentation(loaded));
  const xml = getShapeXmlString(getSlideTables(getSlides(reopened)[0]!)[0]!);
  expect(xml.match(/x:custom="keep"/g)).toHaveLength(2);
  expect(xml).toContain('<a:ext uri="urn:office-kit:table-test"/>');
});
