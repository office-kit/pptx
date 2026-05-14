// Table authoring via graphicFrame.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation, inches } from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/presentation.ts';
import { partName } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('L3: Slide.addTable', () => {
  it('emits a graphicFrame with a:tbl matching rows × cols', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const tbl = slide.addTable({
      x: inches(1),
      y: inches(1),
      w: inches(8),
      h: inches(3),
      rows: [
        ['Name', 'Score', 'Notes'],
        ['Alice', '42', 'first'],
        ['Bob', '7', 'second'],
      ],
    });
    expect(tbl.kind).toBe('graphicFrame');
    expect(tbl.position).toEqual({ x: inches(1), y: inches(1) });
    expect(tbl.size).toEqual({ w: inches(8), h: inches(3) });

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('<a:tbl>');
    expect(xml).toContain('<a:tblPr');
    expect(xml).toContain('<a:tblGrid>');
    expect((xml.match(/<a:gridCol/g) ?? []).length).toBe(3);
    expect((xml.match(/<a:tr /g) ?? []).length).toBe(3);
    expect((xml.match(/<a:tc>/g) ?? []).length).toBe(9);
    expect(xml).toContain('Alice');
    expect(xml).toContain('Bob');
    expect(xml).toContain('Score');
  });

  it('respects custom column widths and row heights', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    slide.addTable({
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      colWidths: [inches(1), inches(2), inches(3)],
      rowHeights: [inches(0.5), inches(1.5)],
      rows: [
        ['A', 'B', 'C'],
        ['1', '2', '3'],
      ],
    });
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    // Distinct column widths land on the gridCol elements.
    expect(xml).toMatch(/<a:gridCol w="914400"\/>/); // 1in
    expect(xml).toMatch(/<a:gridCol w="1828800"\/>/); // 2in
    expect(xml).toMatch(/<a:gridCol w="2743200"\/>/); // 3in
  });

  it('rejects ragged rows', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    expect(() =>
      slide.addTable({
        x: inches(0),
        y: inches(0),
        w: inches(4),
        h: inches(2),
        rows: [
          ['a', 'b'],
          ['c'], // wrong length
        ],
      }),
    ).toThrow(/row 1 has 1 cells; expected 2/);
  });

  it('rejects empty input', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    expect(() =>
      slide.addTable({
        x: inches(0),
        y: inches(0),
        w: inches(4),
        h: inches(2),
        rows: [],
      }),
    ).toThrow(/at least one row/);
    expect(() =>
      slide.addTable({
        x: inches(0),
        y: inches(0),
        w: inches(4),
        h: inches(2),
        rows: [[]],
      }),
    ).toThrow(/at least one column/);
  });

  skipIfNoXmllint('table-bearing slide validates against pml.xsd', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    slide.addTable({
      x: inches(1),
      y: inches(1),
      w: inches(8),
      h: inches(3),
      rows: [
        ['H1', 'H2'],
        ['r1c1', 'r1c2'],
        ['r2c1', 'r2c2'],
      ],
    });
    const pkg = _internalPackageOf(pres);
    expectSchemaValid(
      decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array()),
      'pml',
    );
  });

  it('table content round-trips through save/reload', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    slide.addTable({
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [
        ['k', 'v'],
        ['a', '1'],
        ['b', '2'],
      ],
    });
    // Re-load the saved bytes and inspect the raw slide XML: cell text
    // lives inside `a:tc/a:txBody`, which `slide.text` doesn't (yet)
    // surface. Table-aware text extraction is a separate task.
    const bytes = await pres.save();
    const reloaded = await Presentation.load(bytes);
    const repkg = _internalPackageOf(reloaded);
    const xml = decode(repkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    for (const expected of ['k', 'v', 'a', '1', 'b', '2']) {
      expect(xml).toContain(expected);
    }
  });
});
