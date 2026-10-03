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
  insertTableRow,
  insertTableColumn,
  removeTableRow,
  removeTableColumn,
  loadPresentation,
  mergeTableCells,
  savePresentation,
  setTableStyleFlags,
} from '../src/api/index.ts';
import { CELL_ELEMENT } from '../src/api/_internal-symbols.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { elem } from '../src/internal/xml/ast.ts';
import { firstChildElement } from '../src/internal/xml/query.ts';
import { NS, qname } from '../src/internal/xml/index.ts';

const STYLE_ID = '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}';
const STYLE_XML = `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${STYLE_ID}"><a:tblStyle styleId="${STYLE_ID}" styleName="Edges"><a:wholeTbl><a:tcStyle><a:tcBdr><a:left><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:left><a:right><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:right><a:top><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:top><a:bottom><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:bottom><a:insideH><a:ln w="6350"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></a:insideH><a:insideV><a:ln w="6350"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></a:insideV></a:tcBdr><a:fill><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl><a:firstRow><a:tcStyle><a:tcBdr><a:bottom><a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></a:bottom></a:tcBdr></a:tcStyle></a:firstRow></a:tblStyle></a:tblStyleLst>`;

const BAND_STYLE_XML = `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${STYLE_ID}"><a:tblStyle styleId="${STYLE_ID}" styleName="Band edges"><a:wholeTbl><a:tcStyle/></a:wholeTbl><a:band1H><a:tcStyle><a:tcBdr><a:left><a:ln w="1016"><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill></a:ln></a:left><a:top><a:ln w="1016"><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:ln></a:top><a:insideH><a:ln w="1016"><a:solidFill><a:srgbClr val="AA00AA"/></a:solidFill></a:ln></a:insideH><a:insideV><a:ln w="1016"><a:solidFill><a:srgbClr val="AA00AA"/></a:solidFill></a:ln></a:insideV></a:tcBdr></a:tcStyle></a:band1H><a:band1V><a:tcStyle><a:tcBdr><a:left><a:ln w="1016"><a:solidFill><a:srgbClr val="0000AA"/></a:solidFill></a:ln></a:left><a:top><a:ln w="1016"><a:solidFill><a:srgbClr val="00AAAA"/></a:solidFill></a:ln></a:top><a:insideH><a:ln w="1016"><a:solidFill><a:srgbClr val="AA5500"/></a:solidFill></a:ln></a:insideH><a:insideV><a:ln w="1016"><a:solidFill><a:srgbClr val="AA5500"/></a:solidFill></a:ln></a:insideV></a:tcBdr></a:tcStyle></a:band1V></a:tblStyle></a:tblStyleLst>`;

const makeDeck = async (styleXml = STYLE_XML) => {
  const original = createPresentation();
  addSlideTable(addBlankSlide(original), {
    x: inches(0),
    y: inches(0),
    w: inches(6),
    h: inches(4),
    rows: [
      ['a', 'b', 'c'],
      ['d', 'e', 'f'],
      ['g', 'h', 'i'],
    ],
  });
  const { entries } = readZip(await savePresentation(original));
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  return loadPresentation(
    writeZip(
      entries.map((entry) => {
        if (entry.name === 'ppt/tableStyles.xml')
          return { ...entry, data: encoder.encode(styleXml) };
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

describe('getTableCellAppearanceEffective table-style edges', () => {
  it('updates outer-edge membership after inserting and removing rows and columns', async () => {
    const pres = await makeDeck();
    const table = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setTableStyleFlags(table, { firstRow: false });
    const appearance = () => getTableCellAppearanceEffective(pres, getTableCells(table)[2]![2]!);
    expect(appearance().borders.bottom?.color).toBe('#FF0000');
    expect(appearance().borders.right?.color).toBe('#FF0000');
    insertTableRow(table);
    expect(appearance().borders.bottom?.color).toBe('#0000FF');
    insertTableColumn(table);
    expect(appearance().borders.right?.color).toBe('#0000FF');
    removeTableRow(table, 3);
    expect(appearance().borders.bottom?.color).toBe('#FF0000');
    removeTableColumn(table, 3);
    expect(appearance().borders.right?.color).toBe('#FF0000');
  });

  it('maps outer and interior edges relative to the table and honors first-row overrides', async () => {
    const pres = await makeDeck();
    const table = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setTableStyleFlags(table, { firstRow: true });
    const cells = getTableCells(table);
    expect(getTableCellAppearanceEffective(pres, cells[0]![0]!).borders).toMatchObject({
      left: { color: '#FF0000' },
      top: { color: '#FF0000' },
      bottom: { color: '#00FF00' },
    });
    expect(getTableCellAppearanceEffective(pres, cells[1]![1]!).borders).toMatchObject({
      left: { color: '#0000FF' },
      right: { color: '#0000FF' },
      top: { color: '#0000FF' },
      bottom: { color: '#0000FF' },
    });
    expect(getTableCellAppearanceEffective(pres, cells[2]![2]!).borders).toMatchObject({
      right: { color: '#FF0000' },
      bottom: { color: '#FF0000' },
    });
  });

  it('uses merged visual boundaries and keeps explicit noFill local', async () => {
    const pres = await makeDeck();
    const table = getSlideShapes(getSlides(pres)[0]!)[0]!;
    const cells = getTableCells(table);
    mergeTableCells(table, { row: 1, col: 1, rowSpan: 1, colSpan: 2 });
    const merged = getTableCells(table)[1]![1]!;
    expect(getTableCellAppearanceEffective(pres, merged).borders.right).toMatchObject({
      color: '#FF0000',
    });
    const tcPr = firstChildElement(cells[1]![1]![CELL_ELEMENT], qname('a', 'tcPr', NS.dml));
    tcPr?.children.push(elem(qname('a', 'noFill', NS.dml)));
    expect(getTableCellAppearanceEffective(pres, cells[1]![1]!).fill).toEqual({ kind: 'none' });
  });

  it('keeps band borders on their cell sides instead of remapping them to whole-table interiors', async () => {
    const pres = await makeDeck(BAND_STYLE_XML);
    const table = getSlideShapes(getSlides(pres)[0]!)[0]!;
    setTableStyleFlags(table, { firstRow: true, firstCol: true, bandRow: true, bandCol: true });
    const cells = getTableCells(table);

    // Row 1 is the first horizontal band, but is not the table boundary.
    expect(getTableCellAppearanceEffective(pres, cells[1]![2]!).borders).toMatchObject({
      left: { color: '#AA0000' },
      top: { color: '#00AA00' },
      bottom: { color: '#AA00AA' },
    });

    // Row 2 / column 1 isolates the first vertical band from the horizontal band.
    expect(getTableCellAppearanceEffective(pres, cells[2]![1]!).borders).toMatchObject({
      left: { color: '#0000AA' },
      top: { color: '#00AAAA' },
      right: { color: '#AA5500' },
    });
  });
});
