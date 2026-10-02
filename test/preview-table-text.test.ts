// Unit tests for table-cell text fidelity in `renderSlideToSvg`.
//
// Cell text is authored through the public API (`setTableCellTextFormat` /
// `setTableCellAlignment`) and rendered in both text-layout modes, asserting
// that per-cell run format, alignment, the unstyled default size, wrapping,
// and the foreignObject path all reach the output.
//
// Import pattern follows test/preview-render-svg.test.ts: import from package
// source directly so vitest resolves TypeScript.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTable,
  findSlideLayout,
  getTableCell,
  getTableCellRunFormatEffective,
  inches,
  emu,
  loadPresentation,
  savePresentation,
  setTableCellAlignment,
  setTableCellMargins,
  setTableStyleId,
  setTableStyleFlags,
  setSlideSize,
  getSlideSize,
  getSlides,
  getSlideShapes,
  isTableShape,
  setTableCellTextFormat,
} from '../src/api/index.ts';
import { readZip, writeZip } from '../src/internal/opc/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { attrsOf, countTags, textContentOf } from './lib/svg-query.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const blankSlide = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  const slide = addSlide(pres, { layout });
  return { pres, slide };
};

describe('table cell text rendering', () => {
  it.each(['svg', 'foreignObject'] as const)(
    'fields inherit the same cell text defaults as regular text (%s)',
    async (textLayout) => {
      const { pres, slide } = await blankSlide();
      addSlideTable(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2),
        rows: [['7']],
      });
      const { entries } = readZip(await savePresentation(pres));
      const load = (field: boolean) =>
        loadPresentation(
          writeZip(
            entries.map((entry) => {
              if (!entry.name.startsWith('ppt/slides/slide') || !entry.name.endsWith('.xml'))
                return entry;
              let xml = new TextDecoder()
                .decode(entry.data)
                .replace(
                  '</a:pPr>',
                  '<a:defRPr sz="2800" b="1"><a:latin typeface="Courier New"/></a:defRPr></a:pPr>',
                );
              if (field)
                xml = xml
                  .replace(
                    '<a:r>',
                    '<a:fld id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" type="slidenum">',
                  )
                  .replace('</a:r>', '</a:fld>');
              return { ...entry, data: new TextEncoder().encode(xml) };
            }),
          ),
        );
      const regular = await load(false);
      const field = await load(true);
      expect(renderSlideToSvg(field, getSlides(field).at(-1)!, { textLayout })).toBe(
        renderSlideToSvg(regular, getSlides(regular).at(-1)!, { textLayout }),
      );
    },
  );

  for (const textLayout of ['svg', 'foreignObject'] as const) {
    it.each([false, true])(
      `${textLayout}: omitted cell margins match OOXML defaults (zero left: %s)`,
      async (zeroLeft) => {
        const { pres, slide } = await blankSlide();
        const table = addSlideTable(slide, {
          x: inches(1),
          y: inches(1),
          w: inches(1.5),
          h: inches(2),
          rows: [['Wrapping cell text spans several lines here']],
        });
        const cell = getTableCell(table, 0, 0);
        setTableCellMargins(cell, zeroLeft ? { left: 0 } : null);
        const implicit = renderSlideToSvg(pres, slide, { textLayout });
        setTableCellMargins(cell, {
          left: zeroLeft ? 0 : 91440,
          right: 91440,
          top: 45720,
          bottom: 45720,
        });
        expect(renderSlideToSvg(pres, slide, { textLayout })).toBe(implicit);
      },
    );
  }

  it.each(['svg', 'foreignObject'] as const)(
    'page fitting scales unsized cell text and preserves wrapping (%s)',
    async (textLayout) => {
      const { pres, slide } = await blankSlide();
      addSlideTable(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(2),
        h: inches(3),
        rows: [['Default table text wraps over several lines']],
      });
      const before = renderSlideToSvg(pres, slide, { textLayout });
      const size = getSlideSize(pres)!;
      setSlideSize(
        pres,
        { width: emu(size.width * 2), height: emu(size.height * 2) },
        { content: 'fit' },
      );
      const after = renderSlideToSvg(pres, slide, { textLayout });
      if (textLayout === 'svg') {
        expect(after).toContain('font-size="48"');
        expect(countTags(after, 'text')).toBe(countTags(before, 'text'));
      } else expect(after).toContain('font-size:48.00px');
    },
  );

  it('page fitting scales the default cell text box proportionally', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(2),
      rows: [['Fit table']],
    });
    setTableCellTextFormat(getTableCell(table, 0, 0), { size: 18 });
    const options = { textLayout: 'foreignObject' as const };
    const before = attrsOf(renderSlideToSvg(pres, slide, options), 'foreignObject')[0]!;
    const size = getSlideSize(pres)!;
    setSlideSize(
      pres,
      { width: emu(size.width * 2), height: emu(size.height * 2) },
      { content: 'fit' },
    );
    const after = attrsOf(renderSlideToSvg(pres, slide, options), 'foreignObject')[0]!;
    for (const attr of ['x', 'y', 'width', 'height']) {
      expect(Number(after[attr])).toBeCloseTo(Number(before[attr]) * 2, 2);
    }
  });

  it('svg mode: an explicitly formatted cell carries its size / weight / color per run', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [['Styled', 'Plain']],
    });
    // 28 pt bold red on the first cell; the sibling stays unformatted.
    setTableCellTextFormat(getTableCell(table, 0, 0), { size: 28, bold: true, color: '#CC0000' });

    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    // 28 pt → 28 * 96/72 = 37.33 px.
    expect(svg).toContain('font-size="37.33"');
    expect(svg).toContain('font-weight="700"');
    expect(svg).toMatch(/fill="#[Cc][Cc]0+0+"/);
    // Both cell texts reach the output.
    const text = textContentOf(svg);
    expect(text).toContain('Styled');
    expect(text).toContain('Plain');
  });

  it('svg mode: an unformatted cell falls back to the 18 pt table-cell default', async () => {
    const { pres, slide } = await blankSlide();
    addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [['Plain']],
    });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    // 18 pt → 18 * 96/72 = 24 px, PowerPoint's default for a freshly
    // inserted table cell (no explicit <a:rPr sz>).
    expect(svg).toContain('font-size="24"');
  });

  it('svg mode: a centered cell renders its text with text-anchor="middle"', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [['Centered']],
    });
    setTableCellAlignment(getTableCell(table, 0, 0), 'center');
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const anchors = attrsOf(svg, 'text').map((a) => a['text-anchor']);
    expect(anchors).toContain('middle');
  });

  it('svg mode: long cell text wraps within the cell width across multiple lines', async () => {
    const { pres, slide } = await blankSlide();
    // A single narrow cell forces the long sentence to wrap.
    addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(1.5),
      h: inches(2),
      rows: [['Wrapping cell text spans several lines here']],
    });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    // One <text> element is emitted per laid-out line; wrapping yields >= 2.
    expect(countTags(svg, 'text')).toBeGreaterThanOrEqual(2);
  });

  it('foreignObject mode: cell text is emitted inside a <foreignObject>', async () => {
    const { pres, slide } = await blankSlide();
    const table = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [['Styled', 'Plain']],
    });
    setTableCellTextFormat(getTableCell(table, 0, 0), { size: 28, bold: true, color: '#CC0000' });
    const svg = renderSlideToSvg(pres, slide, { textLayout: 'foreignObject' });
    expect(countTags(svg, 'foreignObject')).toBeGreaterThan(0);
    const text = textContentOf(svg);
    expect(text).toContain('Styled');
    expect(text).toContain('Plain');
    // The styled cell's run still carries its bold weight and red color.
    expect(svg).toContain('font-weight:700');
    expect(svg).toMatch(/color:#[Cc][Cc]0+0+/);
  });

  it.each(['svg', 'foreignObject'] as const)(
    'table-cell paragraph defaults reach the renderer and explicit false wins (%s)',
    async (textLayout) => {
      const { pres, slide } = await blankSlide();
      addSlideTable(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(2),
        rows: [['Inherited']],
      });
      const { entries } = readZip(await savePresentation(pres));
      const decoder = new TextDecoder();
      const encoder = new TextEncoder();
      const mutated = writeZip(
        entries.map((entry) =>
          entry.name === 'ppt/slides/slide1.xml'
            ? {
                ...entry,
                data: encoder.encode(
                  decoder
                    .decode(entry.data)
                    // The table cell's paragraph is the only paragraph in
                    // this minimal slide. Its defRPr supplies the effective
                    // size/weight/fill; the run's explicit b=0 must override
                    // the inherited bold value without erasing the others.
                    .replaceAll(
                      '<a:pPr marL="0" indent="0">',
                      '<a:pPr marL="0" indent="0"><a:defRPr sz="2400" b="1"><a:solidFill><a:srgbClr val="CC0000"/></a:solidFill></a:defRPr>',
                    )
                    .replaceAll('<a:r><a:rPr lang="en-US">', '<a:r><a:rPr lang="en-US" b="0">')
                    .replace(/(<a:r><a:rPr[^>]*>)<a:solidFill>[\s\S]*?<\/a:solidFill>/, '$1'),
                ),
              }
            : entry,
        ),
      );
      const loaded = await loadPresentation(mutated);
      const effective = getTableCellRunFormatEffective(
        loaded,
        getTableCell(getSlideShapes(getSlides(loaded)[0]!).find(isTableShape)!, 0, 0),
        0,
        0,
      );
      expect(effective).toMatchObject({ size: 24, bold: false, color: '#CC0000' });
      const output = renderSlideToSvg(loaded, getSlides(loaded)[0]!, { textLayout });
      expect(output).toContain(textLayout === 'svg' ? 'font-size="32"' : 'font-size:32.00px');
      expect(output).toMatch(textLayout === 'svg' ? /fill="#CC0000"/ : /color:#CC0000/);
      // Explicit b=0 overrides the paragraph default in both output modes.
      expect(output).not.toMatch(/font-weight(?:=|:)\s*700/);
      expect(textContentOf(output)).toContain('Inherited');
    },
  );

  it.each(['svg', 'foreignObject'] as const)(
    'tableStyles.xml wholeTbl and firstRow text styles are inherited (%s)',
    async (textLayout) => {
      const { pres, slide } = await blankSlide();
      addSlideTable(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(6),
        h: inches(2),
        rows: [['Header'], ['Body']],
      });
      const tableShape = getSlideShapes(slide).find(isTableShape)!;
      setTableStyleFlags(tableShape, { firstRow: true });

      const { entries } = readZip(await savePresentation(pres));
      const decoder = new TextDecoder();
      const encoder = new TextEncoder();
      const styleId = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
      const stylesXml = `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${styleId}"><a:tblStyle styleId="${styleId}" styleName="Regression"><a:wholeTbl><a:tcTxStyle b="on"><a:font><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:font><a:srgbClr val="008800"/></a:tcTxStyle></a:wholeTbl><a:firstRow><a:tcTxStyle b="off"><a:font><a:latin typeface="Courier New"/><a:ea typeface="Courier New"/><a:cs typeface="Courier New"/></a:font><a:srgbClr val="CC0000"/></a:tcTxStyle></a:firstRow></a:tblStyle></a:tblStyleLst>`;
      const mutated = writeZip(
        entries.map((entry) =>
          entry.name === 'ppt/tableStyles.xml'
            ? { ...entry, data: encoder.encode(stylesXml) }
            : entry.name === 'ppt/slides/slide1.xml'
              ? {
                  ...entry,
                  // The generated table run carries an explicit black fill;
                  // remove it so the table-style text color is observable.
                  data: encoder.encode(
                    decoder
                      .decode(entry.data)
                      .replace(/(<a:r><a:rPr[^>]*>)<a:solidFill>[\s\S]*?<\/a:solidFill>/g, '$1'),
                  ),
                }
              : entry,
        ),
      );
      const loaded = await loadPresentation(mutated);
      const loadedSlide = getSlides(loaded)[0]!;
      const loadedTable = getSlideShapes(loadedSlide).find(isTableShape)!;
      expect(
        getTableCellRunFormatEffective(loaded, getTableCell(loadedTable, 0, 0), 0, 0),
      ).toMatchObject({
        font: 'Courier New',
        bold: false,
        color: '#CC0000',
      });
      expect(
        getTableCellRunFormatEffective(loaded, getTableCell(loadedTable, 1, 0), 0, 0),
      ).toMatchObject({
        font: 'Arial',
        bold: true,
        color: '#008800',
      });

      const output = renderSlideToSvg(loaded, loadedSlide, { textLayout });
      expect(textContentOf(output)).toContain('Header');
      expect(textContentOf(output)).toContain('Body');
      expect(output).toMatch(
        textLayout === 'svg'
          ? /font-family="(?:Courier New|Liberation Mono)"/
          : /font-family:Courier New/,
      );
      expect(output).toMatch(
        textLayout === 'svg' ? /font-family="(?:Arial|Liberation Sans)"/ : /font-family:Arial/,
      );
      expect(output).toMatch(textLayout === 'svg' ? /fill="#CC0000"/ : /color:#CC0000/);
      expect(output).toMatch(textLayout === 'svg' ? /fill="#008800"/ : /color:#008800/);
    },
  );

  it('renders embedded table-style cell fills and borders, preserving explicit noFill', async () => {
    const { pres, slide } = await blankSlide();
    const tableShape = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [['Header'], ['Transparent']],
    });
    setTableStyleFlags(tableShape, { firstRow: true, lastRow: true, bandRow: false });

    const { entries } = readZip(await savePresentation(pres));
    const decoder = new TextDecoder();
    const encoder = new TextEncoder();
    let mutatedSlideXml = '';
    const styleId = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
    const stylesXml = `<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${styleId}"><a:tblStyle styleId="${styleId}" styleName="Appearance regression"><a:wholeTbl><a:tcStyle><a:tcBdr><a:left><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:left><a:right><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:right><a:top><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:top><a:bottom><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl><a:firstRow><a:tcStyle><a:tcBdr><a:bottom><a:ln w="25400"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:fill></a:tcStyle></a:firstRow><a:lastRow><a:tcStyle><a:tcBdr><a:top><a:ln w="38100"><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill></a:ln></a:top></a:tcBdr></a:tcStyle></a:lastRow></a:tblStyle></a:tblStyleLst>`;
    const mutated = writeZip(
      entries.map((entry) => {
        if (entry.name === 'ppt/tableStyles.xml') {
          return { ...entry, data: encoder.encode(stylesXml) };
        }
        if (entry.name !== 'ppt/slides/slide1.xml') return entry;
        let tcPrIndex = 0;
        const xml = decoder.decode(entry.data).replace(/<a:tcPr(?:\s[^>]*)?\/>/g, () => {
          tcPrIndex += 1;
          return tcPrIndex === 2 ? '<a:tcPr><a:noFill/></a:tcPr>' : '<a:tcPr/>';
        });
        mutatedSlideXml = xml;
        return { ...entry, data: encoder.encode(xml) };
      }),
    );

    expect(mutatedSlideXml).toContain('<a:tcPr><a:noFill/></a:tcPr>');
    const loaded = await loadPresentation(mutated);
    const loadedSlide = getSlides(loaded)[0]!;
    const output = renderSlideToSvg(loaded, loadedSlide, { textLayout: 'svg' });
    expect(output).toMatch(/<g data-pptx-cell="0,0"><rect[^>]*fill="#00AA00"/);
    expect(output).toMatch(/<g data-pptx-cell="1,0"><rect[^>]*fill="none"/);
    expect(output).toMatch(/x1="96\.00" y1="192\.00" x2="672\.00" y2="192\.00" stroke="#00FFFF"/);
    expect(output).not.toMatch(
      /x1="96\.00" y1="192\.00" x2="672\.00" y2="192\.00" stroke="#FF0000"/,
    );
    expect(output).not.toContain('stroke="#9CA3AF"');
    expect(output).not.toMatch(
      /<rect x="96\.00" y="96\.00" width="576\.00" height="192\.00" fill="#FFFFFF"/,
    );
  });

  it('does not invent table paint when the referenced style is unavailable', async () => {
    const { pres, slide } = await blankSlide();
    const tableShape = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [['No style']],
    });
    setTableStyleId(tableShape, '{00000000-0000-0000-0000-000000000000}');
    setTableStyleFlags(tableShape, { firstRow: true, bandRow: true });
    const output = renderSlideToSvg(pres, getSlides(pres)[0]!, { textLayout: 'svg' });
    expect(output).toMatch(/<g data-pptx-cell="0,0"><rect[^>]*fill="none"/);
    expect(output).not.toContain('stroke="#9CA3AF"');
  });
});

it.each(['svg', 'foreignObject'] as const)(
  'retains the height of a sized empty table paragraph (%s)',
  async (textLayout) => {
    const { pres, slide } = await blankSlide();
    addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(3),
      rows: [['After']],
    });
    const { entries } = readZip(await savePresentation(pres));
    const outputs: string[] = [];
    for (const size of [1800, 3600]) {
      const loaded = await loadPresentation(
        writeZip(
          entries.map((entry) =>
            entry.name === 'ppt/slides/slide1.xml'
              ? {
                  ...entry,
                  data: new TextEncoder().encode(
                    new TextDecoder()
                      .decode(entry.data)
                      .replace('<a:p>', `<a:p><a:endParaRPr sz="${size}"/></a:p><a:p>`),
                  ),
                }
              : entry,
          ),
        ),
      );
      outputs.push(renderSlideToSvg(loaded, getSlides(loaded)[0]!, { textLayout }));
    }
    if (textLayout === 'svg') {
      const beforeY = Number(attrsOf(outputs[0]!, 'text')[0]!.y);
      const afterY = Number(attrsOf(outputs[1]!, 'text')[0]!.y);
      expect(afterY - beforeY).toBeGreaterThan(20);
    } else {
      expect(attrsOf(outputs[0]!, 'p')[0]!.style).toContain('font-size:24.00px');
      expect(attrsOf(outputs[1]!, 'p')[0]!.style).toContain('font-size:48.00px');
    }
    expect(textContentOf(outputs[0]!)).toContain('After');
    expect(textContentOf(outputs[1]!)).toContain('After');
  },
);

it.each(['svg', 'foreignObject'] as const)(
  'retains the height of a sized leading table line break (%s)',
  async (textLayout) => {
    const { pres, slide } = await blankSlide();
    addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(3),
      rows: [['After']],
    });
    const { entries } = readZip(await savePresentation(pres));
    const outputs: string[] = [];
    for (const size of [1800, 3600]) {
      const loaded = await loadPresentation(
        writeZip(
          entries.map((entry) =>
            entry.name === 'ppt/slides/slide1.xml'
              ? {
                  ...entry,
                  data: new TextEncoder().encode(
                    new TextDecoder()
                      .decode(entry.data)
                      .replace('<a:p>', `<a:p><a:br><a:rPr sz="${size}"/></a:br>`),
                  ),
                }
              : entry,
          ),
        ),
      );
      outputs.push(renderSlideToSvg(loaded, getSlides(loaded)[0]!, { textLayout }));
    }
    if (textLayout === 'svg') {
      const beforeY = Number(attrsOf(outputs[0]!, 'text')[0]!.y);
      const afterY = Number(attrsOf(outputs[1]!, 'text')[0]!.y);
      expect(afterY - beforeY).toBeGreaterThan(20);
    } else {
      expect(attrsOf(outputs[0]!, 'span')[0]!.style).toContain('font-size:24.00px');
      expect(attrsOf(outputs[1]!, 'span')[0]!.style).toContain('font-size:48.00px');
    }
    expect(textContentOf(outputs[0]!)).toContain('After');
    expect(textContentOf(outputs[1]!)).toContain('After');
  },
);
