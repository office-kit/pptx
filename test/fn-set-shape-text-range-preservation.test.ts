import { expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  createPresentation,
  getShapeXmlString,
  getShapeText,
  getTableCell,
  getSlideTables,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeRunHyperlink,
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
