// `TextFormat.fontComplexScript` — the `<a:cs>` typeface DrawingML picks for
// complex scripts (Arabic, Hebrew, Thai). Covers the three write sites (run,
// `<a:endParaRPr>`, table cell), the matching readers, and the child order
// CT_TextCharacterProperties mandates: latin → ea → cs → sym → hlinkClick.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  type PresentationData,
  addSlideTable,
  getParagraphEndFormat,
  getShapeParagraphElements,
  getShapeRunFormat,
  getSlideShapes,
  getSlides,
  getTableCell,
  getTableCellParagraphs,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setShapeParagraphs,
  setShapeRunFormat,
  setShapeRunHyperlink,
  setTableCellParagraphs,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decoder = new TextDecoder();
const encoder = new TextEncoder();
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const THAI = 'Leelawadee UI';
const ARABIC = 'Traditional Arabic';

const slideXml = (pres: PresentationData): string =>
  decoder.decode(readPackagePart(pres, '/ppt/slides/slide1.xml')!);

const textSlide = async (): Promise<PresentationData> =>
  loadPresentation(await readFile(fixture('one-text-slide.pptx')));

describe('fn API: TextFormat.fontComplexScript on runs', () => {
  it('round-trips independently of font and fontEastAsian', async () => {
    const pres = await textSlide();
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeRunFormat(shape, 0, 0, {
      font: 'Georgia',
      fontEastAsian: '游明朝',
      fontComplexScript: THAI,
    });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const fmt = getShapeRunFormat(getSlideShapes(getSlides(reloaded)[0]!)[0]!, 0, 0)!;
    expect(fmt.font).toBe('Georgia');
    expect(fmt.fontEastAsian).toBe('游明朝');
    expect(fmt.fontComplexScript).toBe(THAI);
  });

  it('is absent from the read format when the run carries no <a:cs>', async () => {
    const pres = await textSlide();
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeRunFormat(shape, 0, 0, { font: 'Georgia' });
    expect(slideXml(await loadPresentation(await savePresentation(pres)))).not.toContain('<a:cs');
    expect(getShapeRunFormat(shape, 0, 0)!.fontComplexScript).toBeUndefined();
  });

  it('survives setShapeParagraphs → getShapeParagraphElements', async () => {
    const pres = await textSlide();
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeParagraphs(shape, [
      { runs: [{ text: 'สวัสดี', format: { size: 14, fontComplexScript: THAI } }] },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const again = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeParagraphElements(again, 0)).toEqual([
      { kind: 'r', text: 'สวัสดี', format: { size: 14, fontComplexScript: THAI } },
    ]);
  });
});

describe('fn API: TextFormat.fontComplexScript on the paragraph end mark', () => {
  skipIfNoXmllint('writes <a:cs> into <a:endParaRPr> and reads it back', async () => {
    const pres = await textSlide();
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeParagraphs(shape, [
      { runs: [{ text: 'x' }], endFormat: { size: 9, font: 'Arial', fontComplexScript: ARABIC } },
      { runs: [], endFormat: {} },
      { runs: [{ text: 'y' }] },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const again = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getParagraphEndFormat(again, 0)).toEqual({
      size: 9,
      font: 'Arial',
      fontComplexScript: ARABIC,
    });
    // An empty endFormat still writes the end mark, and carries no typeface.
    expect(getParagraphEndFormat(again, 1)).toEqual({});
    expect(getParagraphEndFormat(again, 2)).toBeNull();

    const xml = slideXml(reloaded);
    expect(xml).toContain(
      `<a:endParaRPr sz="900"><a:latin typeface="Arial"/><a:cs typeface="${ARABIC}"/></a:endParaRPr>`,
    );
    expect(xml).toContain('<a:p><a:endParaRPr/></a:p>');
    expectSchemaValid(xml, 'pml');
  });
});

describe('fn API: TextFormat.fontComplexScript in table cells', () => {
  skipIfNoXmllint('round-trips on a cell paragraph and its end mark', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const table = addSlideTable(getSlides(pres)[0]!, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(1),
      rows: [['head'], ['body']],
    });
    setTableCellParagraphs(getTableCell(table, 0, 0), [
      {
        runs: [{ text: 'مرحبا', format: { fontComplexScript: ARABIC } }],
        endFormat: { fontComplexScript: ARABIC },
      },
    ]);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const again = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getTableCellParagraphs(getTableCell(again, 0, 0))).toEqual([
      {
        align: null,
        elements: [{ kind: 'r', text: 'مرحبا', format: { fontComplexScript: ARABIC } }],
        endFormat: { fontComplexScript: ARABIC },
      },
    ]);
    expectSchemaValid(slideXml(reloaded), 'pml');
  });
});

describe('fn API: <a:cs> child order in CT_TextCharacterProperties', () => {
  skipIfNoXmllint('writes latin, then ea, then cs', async () => {
    const pres = await textSlide();
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    // Deliberately set the complex script first: the order in the XML comes
    // from the schema rank, not from the order of the format's keys.
    setShapeRunFormat(shape, 0, 0, { fontComplexScript: THAI });
    setShapeRunFormat(shape, 0, 0, { fontEastAsian: '游明朝' });
    setShapeRunFormat(shape, 0, 0, { font: 'Georgia' });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const xml = slideXml(reloaded);
    expect(xml).toContain(
      `<a:latin typeface="Georgia"/><a:ea typeface="游明朝"/><a:cs typeface="${THAI}"/>`,
    );
    expectSchemaValid(xml, 'pml');
  });

  skipIfNoXmllint('inserts <a:cs> before an existing <a:hlinkClick>', async () => {
    const pres = await textSlide();
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeRunHyperlink(shape, 0, 0, 'https://example.com/');
    setShapeRunFormat(shape, 0, 0, { font: 'Georgia', fontComplexScript: THAI });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const xml = slideXml(reloaded);
    expect(xml).toMatch(
      new RegExp(`<a:latin typeface="Georgia"/><a:cs typeface="${THAI}"/><a:hlinkClick\\b`),
    );
    expectSchemaValid(xml, 'pml');
  });

  skipIfNoXmllint('inserts <a:cs> before an existing <a:sym>', async () => {
    // No public API authors <a:sym>, so it is injected at the zip layer.
    const pres = await textSlide();
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeRunFormat(shape, 0, 0, { font: 'Georgia' });
    const { entries } = readZip(await savePresentation(pres));
    const withSym = entries.map((e) =>
      e.name.endsWith('slides/slide1.xml')
        ? {
            name: e.name,
            data: encoder.encode(
              decoder
                .decode(e.data)
                .replace(
                  '<a:latin typeface="Georgia"/>',
                  '<a:latin typeface="Georgia"/><a:sym typeface="Wingdings"/>',
                ),
            ),
          }
        : e,
    );

    const injected = await loadPresentation(writeZip(withSym));
    const injectedShape = getSlideShapes(getSlides(injected)[0]!)[0]!;
    setShapeRunFormat(injectedShape, 0, 0, { fontComplexScript: THAI });

    const reloaded = await loadPresentation(await savePresentation(injected));
    const xml = slideXml(reloaded);
    expect(xml).toContain(
      `<a:latin typeface="Georgia"/><a:cs typeface="${THAI}"/><a:sym typeface="Wingdings"/>`,
    );
    expectSchemaValid(xml, 'pml');
  });

  it('replaces an existing <a:cs> instead of adding a second', async () => {
    const pres = await textSlide();
    const shape = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setShapeRunFormat(shape, 0, 0, { fontComplexScript: THAI });
    setShapeRunFormat(shape, 0, 0, { fontComplexScript: ARABIC });

    const xml = slideXml(await loadPresentation(await savePresentation(pres)));
    expect(xml.match(/<a:cs /g)).toHaveLength(1);
    expect(xml).toContain(`<a:cs typeface="${ARABIC}"/>`);
  });
});
