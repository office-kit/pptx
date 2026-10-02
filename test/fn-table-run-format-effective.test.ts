import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getPresentationFonts,
  getSlides,
  getSlideShapes,
  getTableCellRunFormatEffective,
  getTableCellAppearanceEffective,
  getTableCells,
  mergeTableCells,
  setTableStyleFlags,
  inches,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

describe('fn API: getTableCellRunFormatEffective', () => {
  it('resolves field defaults and direct overrides without changing regular run indices', async () => {
    const original = createPresentation();
    addSlideTable(addBlankSlide(original), {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [['Regular']],
    });
    const { entries } = readZip(await savePresentation(original));
    const loaded = await loadPresentation(
      writeZip(
        entries.map((entry) =>
          entry.name === 'ppt/slides/slide1.xml'
            ? {
                ...entry,
                data: new TextEncoder().encode(
                  new TextDecoder()
                    .decode(entry.data)
                    .replace(
                      '</a:pPr>',
                      '<a:defRPr sz="2800" b="1"><a:latin typeface="Courier New"/></a:defRPr></a:pPr>',
                    )
                    .replace(
                      '</a:p>',
                      '<a:fld id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" type="slidenum"><a:rPr b="0"/><a:t>7</a:t></a:fld></a:p>',
                    ),
                ),
              }
            : entry,
        ),
      ),
    );
    for (const pres of [loaded, await loadPresentation(await savePresentation(loaded))]) {
      const cell = getTableCells(getSlideShapes(getSlides(pres)[0]!)[0]!)[0]![0]!;
      expect(getTableCellRunFormatEffective(pres, cell, 0, { fieldIndex: 0 })).toMatchObject({
        size: 28,
        bold: false,
        font: 'Courier New',
      });
      expect(() => getTableCellRunFormatEffective(pres, cell, 0, 0)).not.toThrow();
      expect(() => getTableCellRunFormatEffective(pres, cell, 0, 1)).toThrow(RangeError);
      expect(() => getTableCellRunFormatEffective(pres, cell, 0, { fieldIndex: 1 })).toThrow(
        RangeError,
      );
    }
  });

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

  it('resolves custom tableStyles.xml tcTxStyle after the cell text cascade', async () => {
    const original = createPresentation();
    addSlideTable(addBlankSlide(original), {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [
        ['Header', 'Body'],
        ['Body', 'Body'],
      ],
    });
    const { entries } = readZip(await savePresentation(original));
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    const styleId = '{11111111-2222-3333-4444-555555555555}';
    const loaded = await loadPresentation(
      writeZip(
        entries.map((entry) => {
          if (entry.name === 'ppt/tableStyles.xml') {
            return {
              ...entry,
              data: encoder.encode(
                `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${styleId}"><a:tblStyle styleId="${styleId}" styleName="Test"><a:wholeTbl><a:tcTxStyle b="on"><a:fontRef idx="major"><a:schemeClr val="accent1"/></a:fontRef><a:srgbClr val="00FF00"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:left><a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></a:left></a:tcBdr><a:fill><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl><a:band1H><a:tcTxStyle i="off"/></a:band1H><a:firstCol><a:tcTxStyle b="off"/></a:firstCol><a:firstRow><a:tcTxStyle b="on" i="on"><a:font><a:latin typeface="Courier New"/><a:ea typeface="MS Gothic"/><a:cs typeface="Arial"/></a:font></a:tcTxStyle></a:firstRow></a:tblStyle></a:tblStyleLst>`,
              ),
            };
          }
          if (entry.name !== 'ppt/slides/slide1.xml') return entry;
          return {
            ...entry,
            data: encoder.encode(
              decoder
                .decode(entry.data)
                .replace(/<a:solidFill><a:srgbClr val="000000"\/><\/a:solidFill>/g, '')
                .replace(
                  /<a:tableStyleId>[^<]*<\/a:tableStyleId>/,
                  `<a:tableStyleId>${styleId}</a:tableStyleId>`,
                ),
            ),
          };
        }),
      ),
    );
    const table = getSlideShapes(getSlides(loaded)[0]!)[0]!;
    setTableStyleFlags(table, { firstRow: true, firstCol: true, bandRow: true });
    mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 2 });
    const cells = getTableCells(table);
    expect(getTableCellAppearanceEffective(loaded, cells[0]![0]!)).toMatchObject({
      fill: { kind: 'solid', color: '#112233' },
      borders: { left: { widthEmu: 12700, color: '#445566' } },
    });
    expect(getTableCellRunFormatEffective(loaded, cells[0]![0]!, 0, 0)).toMatchObject({
      bold: true,
      font: 'Courier New',
      color: '#00FF00',
      italic: true,
    });
    expect(getTableCellRunFormatEffective(loaded, cells[1]![0]!, 0, 0)).toMatchObject({
      bold: false,
      font: getPresentationFonts(loaded)!.majorLatin,
      color: '#00FF00',
      italic: false,
    });
  });
});

it('resolves an empty cell paragraph end mark through paragraph and table defaults', async () => {
  const original = createPresentation();
  addSlideTable(addBlankSlide(original), {
    x: inches(0),
    y: inches(0),
    w: inches(4),
    h: inches(2),
    rows: [['Empty']],
  });
  const { entries } = readZip(await savePresentation(original));
  const loaded = await loadPresentation(
    writeZip(
      entries.map((entry) =>
        entry.name === 'ppt/slides/slide1.xml'
          ? {
              ...entry,
              data: new TextEncoder().encode(
                new TextDecoder()
                  .decode(entry.data)
                  .replace(
                    /<a:p>[\s\S]*?<\/a:p>/,
                    '<a:p><a:pPr><a:defRPr sz="2400" i="1"/></a:pPr><a:endParaRPr sz="3600" b="0"/></a:p>',
                  ),
              ),
            }
          : entry,
      ),
    ),
  );
  const cell = getTableCells(getSlideShapes(getSlides(loaded)[0]!)[0]!)[0]![0]!;
  expect(getTableCellRunFormatEffective(loaded, cell, 0, null)).toMatchObject({
    size: 36,
    bold: false,
    italic: true,
    color: '#FFFFFF',
  });
  expect(() => getTableCellRunFormatEffective(loaded, cell, 0, 0)).toThrow(RangeError);
});

it('separates existing table runs and fields from the paragraph insertion format', async () => {
  const original = createPresentation();
  addSlideTable(addBlankSlide(original), {
    x: inches(0),
    y: inches(0),
    w: inches(4),
    h: inches(2),
    rows: [['Text']],
  });
  const { entries } = readZip(await savePresentation(original));
  const loaded = await loadPresentation(
    writeZip(
      entries.map((entry) =>
        entry.name === 'ppt/slides/slide1.xml'
          ? {
              ...entry,
              data: new TextEncoder().encode(
                new TextDecoder()
                  .decode(entry.data)
                  .replace(
                    /<a:p>[\s\S]*?<\/a:p>/,
                    '<a:p><a:pPr><a:defRPr sz="2400" b="0"/></a:pPr><a:r><a:t>Text</a:t></a:r><a:fld id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" type="slidenum"><a:t>7</a:t></a:fld><a:endParaRPr sz="4800" b="1"/></a:p>',
                  ),
              ),
            }
          : entry,
      ),
    ),
  );
  const roundTrip = await loadPresentation(await savePresentation(loaded));
  const cell = getTableCells(getSlideShapes(getSlides(roundTrip)[0]!)[0]!)[0]![0]!;
  expect(getTableCellRunFormatEffective(roundTrip, cell, 0, 0)).toMatchObject({
    size: 24,
    bold: false,
  });
  expect(getTableCellRunFormatEffective(roundTrip, cell, 0, { fieldIndex: 0 })).toMatchObject({
    size: 24,
    bold: false,
  });
  expect(getTableCellRunFormatEffective(roundTrip, cell, 0, null)).toMatchObject({
    size: 48,
    bold: true,
  });
});
