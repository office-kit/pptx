import { expect, it } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  inches,
  setShapeParagraphs,
  setShapeTextField,
  setParagraphLevel,
  setParagraphTabs,
  setShapeRunHyperlink,
  getShapeText,
  getShapeXmlString,
  getShapeParagraphElements,
  getShapeParagraphCount,
  getParagraphLevel,
  getShapeRunHyperlink,
  getShapeRunHyperlinkTooltip,
  loadPresentation,
  savePresentation,
  getSlides,
  getSlideShapes,
} from '../src/api/index.ts';

const box = (slide: ReturnType<typeof addBlankSlide>, text = '') =>
  addSlideTextBox(slide, { x: inches(1), y: inches(1), w: inches(4), h: inches(2), text });

it('copies a formatted text range across slides with its paragraph properties and links', async () => {
  const pres = createPresentation();
  const source = box(addBlankSlide(pres));
  setShapeParagraphs(source, [
    { runs: [{ text: 'First' }] },
    { runs: [{ text: 'Second', format: { italic: true } }], endFormat: { bold: true } },
    { runs: [{ text: 'Third' }] },
  ]);
  setParagraphLevel(source, 1, 3);
  setParagraphTabs(source, 1, { defaultTabSizeEmu: inches(0.8) });
  setShapeRunHyperlink(source, 1, 0, 'https://example.com/source', 'Retained tooltip');
  const target = box(addBlankSlide(pres), 'Old target');
  setShapeRunHyperlink(target, 0, 0, 'https://example.com/old');
  setShapeParagraphs(target, { source, range: { start: 6, end: 12 } });
  expect(getShapeText(target)).toBe('Second');
  expect(getParagraphLevel(target, 0)).toBe(3);
  expect(getShapeParagraphElements(target, 0)[0]?.format).toMatchObject({ italic: true });
  expect(getShapeRunHyperlink(target, 0, 0)).toBe('https://example.com/source');
  expect(getShapeRunHyperlinkTooltip(target, 0, 0)).toBe('Retained tooltip');
  expect(getShapeXmlString(target)).toContain('defTabSz="731520"');
  expect(getShapeText(source)).toBe('First\nSecond\nThird');
  const restored = await loadPresentation(await savePresentation(pres));
  const saved = getSlideShapes(getSlides(restored)[1]!)[0]!;
  expect(getShapeRunHyperlink(saved, 0, 0)).toBe('https://example.com/source');
  expect(getParagraphLevel(saved, 0)).toBe(3);
});

it('copies paragraph separators and partial runs, and supports copying within the same shape', () => {
  const slide = addBlankSlide(createPresentation());
  const source = box(slide, 'One\nTwo\n');
  setParagraphLevel(source, 1, 2);
  const target = box(slide);
  setShapeParagraphs(target, { source, range: { start: 1, end: 4 } });
  expect(getShapeText(target)).toBe('ne\n');
  expect(getShapeParagraphCount(target)).toBe(2);
  expect(getParagraphLevel(target, 1)).toBe(2);
  const currentSource = getSlideShapes(slide)[0]!;
  setShapeParagraphs(currentSource, { source: currentSource, range: { start: 4, end: 7 } });
  expect(getShapeText(currentSource)).toBe('Two');
  expect(getParagraphLevel(currentSource, 0)).toBe(2);
});

it('validates text range boundaries before modifying the destination', () => {
  const slide = addBlankSlide(createPresentation());
  const source = box(slide, 'A😀B');
  const target = box(slide, 'Unchanged');
  const before = getShapeXmlString(target);
  for (const range of [
    { start: 2, end: 3 },
    { start: 0, end: 9 },
    { start: 3, end: 1 },
  ]) {
    expect(() => setShapeParagraphs(target, { source, range })).toThrow(RangeError);
    expect(getShapeXmlString(target)).toBe(before);
  }
});

it('copies all text to another presentation without changing source relationships', async () => {
  const sourcePres = createPresentation();
  const source = box(addBlankSlide(sourcePres), 'Cross-presentation link');
  setShapeRunHyperlink(source, 0, 0, 'https://example.com/cross');
  const targetPres = createPresentation();
  const target = box(addBlankSlide(targetPres));
  setShapeParagraphs(target, { source });
  const reopened = await loadPresentation(await savePresentation(targetPres));
  expect(getShapeRunHyperlink(getSlideShapes(getSlides(reopened)[0]!)[0]!, 0, 0)).toBe(
    'https://example.com/cross',
  );
  expect(getShapeRunHyperlink(source, 0, 0)).toBe('https://example.com/cross');
});

it('retains complete fields and converts a partial field selection to literal text', async () => {
  const pres = createPresentation();
  const source = box(addBlankSlide(pres));
  setShapeTextField(source, 'datetime1', { text: '2026-10-01' });
  const target = box(addBlankSlide(pres));
  setShapeParagraphs(target, { source });
  expect(getShapeParagraphElements(target, 0)[0]).toMatchObject({
    kind: 'fld',
    text: '2026-10-01',
  });
  const reopened = await loadPresentation(await savePresentation(pres));
  expect(getShapeParagraphElements(getSlideShapes(getSlides(reopened)[1]!)[0]!, 0)[0]?.kind).toBe(
    'fld',
  );
  setShapeParagraphs(target, { source, range: { start: 0, end: 4 } });
  expect(getShapeParagraphElements(target, 0)[0]).toMatchObject({ kind: 'r', text: '2026' });
});

it('retains soft breaks and imported paragraph extensions in a partial selection', async () => {
  const pres = createPresentation();
  box(addBlankSlide(pres), 'Original');
  const parts = unzipSync(await savePresentation(pres));
  const part = 'ppt/slides/slide1.xml';
  parts[part] = strToU8(
    strFromU8(parts[part]!)
      .replace('<a:t>Original</a:t>', '<a:t>AB</a:t>')
      .replace(
        '</a:p>',
        '<a:br/><a:r><a:t>CD</a:t></a:r><a:extLst><a:ext uri="urn:test"><x:data xmlns:x="urn:custom" value="retained"/></a:ext></a:extLst></a:p>',
      ),
  );
  const loaded = await loadPresentation(zipSync(parts));
  const source = getSlideShapes(getSlides(loaded)[0]!)[0]!;
  const target = box(addBlankSlide(loaded));
  setShapeParagraphs(target, { source, range: { start: 1, end: 4 } });
  expect(getShapeText(target)).toBe('B\nC');
  expect(getShapeParagraphElements(target, 0).map((element) => element.kind)).toEqual([
    'r',
    'br',
    'r',
  ]);
  expect(getShapeXmlString(target)).toContain('value="retained"');
  const reopened = await loadPresentation(await savePresentation(loaded));
  expect(getShapeXmlString(getSlideShapes(getSlides(reopened)[1]!)[0]!)).toContain(
    'value="retained"',
  );
});
