import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getPresentationFonts,
  getSlides,
  getSlideShapes,
  getTableCellRunFormatEffective,
  getTableCells,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

describe('fn API: getTableCellRunFormatEffective', () => {
  it('resolves cell paragraph defaults before outline defaults', async () => {
    const original = createPresentation();
    addSlideTable(addBlankSlide(original), {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [['Inherited']],
    });
    const { entries } = readZip(await savePresentation(original));
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const loaded = await loadPresentation(
      writeZip(
        entries.map((entry) =>
          entry.name === 'ppt/slides/slide1.xml'
            ? {
                ...entry,
                data: encoder.encode(
                  decoder
                    .decode(entry.data)
                    .replace(
                      '</a:pPr>',
                      '<a:defRPr sz="2400" b="1" u="sng"><a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill></a:defRPr></a:pPr>',
                    )
                    .replace(
                      /<a:rPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/,
                      '<a:rPr b="0" u="none"><a:uFillTx/></a:rPr>',
                    ),
                ),
              }
            : entry,
        ),
      ),
    );
    const cell = getTableCells(getSlideShapes(getSlides(loaded)[0]!)[0]!)[0]![0]!;
    expect(getTableCellRunFormatEffective(loaded, cell, 0, 0)).toMatchObject({
      size: 24,
      bold: false,
      underline: false,
      underlineColor: null,
    });
  });

  it.each(['+mn-lt', '+mj-lt'])(
    'resolves table run theme colors and font token %s',
    async (token) => {
      const original = createPresentation();
      addSlideTable(addBlankSlide(original), {
        x: inches(0),
        y: inches(0),
        w: inches(4),
        h: inches(2),
        rows: [['Themed']],
      });
      const { entries } = readZip(await savePresentation(original));
      const decoder = new TextDecoder();
      const encoder = new TextEncoder();
      const loaded = await loadPresentation(
        writeZip(
          entries.map((entry) =>
            entry.name === 'ppt/slides/slide1.xml'
              ? {
                  ...entry,
                  data: encoder.encode(
                    decoder
                      .decode(entry.data)
                      .replace(
                        /<a:rPr\b[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/,
                        `<a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="${token}"/></a:rPr>`,
                      ),
                  ),
                }
              : entry,
          ),
        ),
      );
      const cell = getTableCells(getSlideShapes(getSlides(loaded)[0]!)[0]!)[0]![0]!;
      const format = getTableCellRunFormatEffective(loaded, cell, 0, 0);
      expect(format.color).toBe('#4F81BD');
      expect(format.font).toBe(
        token === '+mn-lt'
          ? getPresentationFonts(loaded)!.minorLatin
          : getPresentationFonts(loaded)!.majorLatin,
      );
    },
  );

  it('uses the paragraph outline level when reading cell lstStyle defaults', async () => {
    const original = createPresentation();
    addSlideTable(addBlankSlide(original), {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [['Outline']],
    });
    const { entries } = readZip(await savePresentation(original));
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const loaded = await loadPresentation(
      writeZip(
        entries.map((entry) =>
          entry.name === 'ppt/slides/slide1.xml'
            ? {
                ...entry,
                data: encoder.encode(
                  decoder
                    .decode(entry.data)
                    .replace(
                      /<a:lstStyle\/>/,
                      '<a:lstStyle><a:lvl2pPr><a:defRPr sz="2800" b="1"/></a:lvl2pPr></a:lstStyle>',
                    )
                    .replace(/<a:pPr([^>]*)>/, '<a:pPr$1 lvl="1">'),
                ),
              }
            : entry,
        ),
      ),
    );
    const cell = getTableCells(getSlideShapes(getSlides(loaded)[0]!)[0]!)[0]![0]!;
    expect(getTableCellRunFormatEffective(loaded, cell, 0, 0)).toMatchObject({
      size: 28,
      bold: true,
    });
  });
});
