import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getSlideShapes,
  getSlides,
  getTableCellAppearanceEffective,
  getTableCells,
  inches,
  loadPresentation,
  savePresentation,
  setTableCellFill,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';

const STYLE_ID = '{11111111-2222-3333-4444-555555555555}';
const STYLE_XML =
  `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${STYLE_ID}">` +
  `<a:tblStyle styleId="${STYLE_ID}" styleName="Theme references">` +
  `<a:wholeTbl><a:tcStyle>` +
  `<a:tcBdr><a:left><a:lnRef idx="2"><a:schemeClr val="accent3"/></a:lnRef></a:left></a:tcBdr>` +
  `<a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef>` +
  `</a:tcStyle></a:wholeTbl>` +
  `</a:tblStyle></a:tblStyleLst>`;

const makeDeck = async (styleXml = STYLE_XML, themeMutation?: (xml: string) => string) => {
  const presentation = createPresentation();
  addSlideTable(addBlankSlide(presentation), {
    x: inches(0),
    y: inches(0),
    w: inches(6),
    h: inches(4),
    rows: [['theme ref']],
  });
  const { entries } = readZip(await savePresentation(presentation));
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  return loadPresentation(
    writeZip(
      entries.map((entry) => {
        if (entry.name === 'ppt/tableStyles.xml') {
          return { ...entry, data: encoder.encode(styleXml) };
        }
        if (entry.name === 'ppt/theme/theme1.xml' && themeMutation) {
          return { ...entry, data: encoder.encode(themeMutation(decoder.decode(entry.data))) };
        }
        if (entry.name !== 'ppt/slides/slide1.xml') return entry;
        return {
          ...entry,
          data: encoder.encode(
            decoder
              .decode(entry.data)
              .replace(
                /<a:tableStyleId>[^<]*<\/a:tableStyleId>/,
                `<a:tableStyleId>${STYLE_ID}</a:tableStyleId>`,
              ),
          ),
        };
      }),
    ),
  );
};

describe('table style matrix references', () => {
  it('resolves background fill references starting at index 1001', async () => {
    const pres = await makeDeck(STYLE_XML.replace('<a:fillRef idx="1">', '<a:fillRef idx="1001">'));
    const table = getSlideShapes(getSlides(pres)[0]!)[0]!;
    expect(getTableCellAppearanceEffective(pres, getTableCells(table)[0]![0]!).fill).toEqual({
      kind: 'solid',
      color: '#C0504D',
    });
  });

  it('resolves fillRef and lnRef through the deck theme', async () => {
    const pres = await makeDeck();
    const table = getSlideShapes(getSlides(pres)[0]!)[0]!;
    const cell = getTableCells(table)[0]![0]!;
    const appearance = getTableCellAppearanceEffective(pres, cell);

    expect(appearance.fill).toEqual({ kind: 'solid', color: '#C0504D' });
    expect(appearance.borders.left).toMatchObject({
      color: '#9BBB59',
      widthEmu: 25400,
    });
  });

  it('composes reference and style color transforms, then honors local fill', async () => {
    const transformed = await makeDeck(
      STYLE_XML.replace(
        '<a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef>',
        '<a:fillRef idx="1"><a:schemeClr val="accent2"><a:shade val="50000"/></a:schemeClr></a:fillRef>',
      ),
      (xml) =>
        xml.replace(
          '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>',
          '<a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>',
        ),
    );
    const table = getSlideShapes(getSlides(transformed)[0]!)[0]!;
    const cell = getTableCells(table)[0]![0]!;
    expect(getTableCellAppearanceEffective(transformed, cell).fill).toEqual({
      kind: 'solid',
      color: '#D0BFBF',
    });

    setTableCellFill(cell, '#123456');
    const refreshed = getTableCells(table)[0]![0]!;
    expect(getTableCellAppearanceEffective(transformed, refreshed).fill).toEqual({
      kind: 'solid',
      color: '#123456',
    });
  });

  it('treats fillRef index zero as no fill and ignores an unavailable style', async () => {
    const noFillStyle = STYLE_XML.replace('idx="1"', 'idx="0"');
    const noFill = await makeDeck(noFillStyle);
    const noFillTable = getSlideShapes(getSlides(noFill)[0]!)[0]!;
    expect(
      getTableCellAppearanceEffective(noFill, getTableCells(noFillTable)[0]![0]!).fill,
    ).toEqual({ kind: 'none' });

    const unavailableStyle = STYLE_XML.replace('idx="1"', 'idx="999"');
    const unavailable = await makeDeck(unavailableStyle);
    const unavailableTable = getSlideShapes(getSlides(unavailable)[0]!)[0]!;
    expect(
      getTableCellAppearanceEffective(unavailable, getTableCells(unavailableTable)[0]![0]!).fill,
    ).toEqual({ kind: 'inherit' });
  });
});
