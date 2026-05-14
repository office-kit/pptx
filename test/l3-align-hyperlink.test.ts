// Paragraph alignment + hyperlinks.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation, inches } from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/presentation.ts';
import { partName, parseRels, relsPartNameFor } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('L3: paragraph alignment', () => {
  it('writes algn="ctr" on every paragraph for setAlignment("center")', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(8),
      h: inches(2),
      text: 'left line\ncenter line',
    });
    box.setText('first\nsecond');
    box.setAlignment('center');

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect((xml.match(/algn="ctr"/g) ?? []).length).toBe(2);
  });

  it('accepts ECMA-376 raw tokens directly', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1),
      text: 'x',
    });
    box.setAlignment('just');
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('algn="just"');
  });

  skipIfNoXmllint('aligned shape validates against pml.xsd', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(8),
      h: inches(2),
      text: 'a\nb',
    });
    box.setAlignment('center');
    const pkg = _internalPackageOf(pres);
    expectSchemaValid(
      decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array()),
      'pml',
    );
  });
});

describe('L3: hyperlinks', () => {
  it('adds a hyperlink rel and a:hlinkClick rPr child', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(1),
      text: 'Click me',
    });
    box.setHyperlink('https://example.com/docs');

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('<a:hlinkClick');
    expect(xml).toMatch(/r:id="rId\d+"/);

    // The slide rels contains the hyperlink target as External.
    const relsBytes =
      pkg.getPart(relsPartNameFor(partName('/ppt/slides/slide1.xml')))?.data ?? new Uint8Array();
    const rels = parseRels(decode(relsBytes));
    const hl = rels.items.find((r) => r.target === 'https://example.com/docs');
    expect(hl).toBeDefined();
    expect(hl?.targetMode).toBe('External');
  });

  it('reuses an existing hyperlink rel for the same URL', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const a = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'A',
    });
    const b = slide.addTextBox({
      x: inches(1),
      y: inches(3),
      w: inches(4),
      h: inches(1),
      text: 'B',
    });
    a.setHyperlink('https://example.com');
    b.setHyperlink('https://example.com');

    const pkg = _internalPackageOf(pres);
    const relsBytes =
      pkg.getPart(relsPartNameFor(partName('/ppt/slides/slide1.xml')))?.data ?? new Uint8Array();
    const rels = parseRels(decode(relsBytes));
    // Only one hyperlink rel for both shapes.
    expect(rels.items.filter((r) => r.target === 'https://example.com').length).toBe(1);
  });

  it('removes the hlinkClick when setHyperlink(null) is called', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'x',
    });
    box.setHyperlink('https://example.com');
    box.setHyperlink(null);
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).not.toContain('<a:hlinkClick');
  });

  skipIfNoXmllint('hyperlink output validates against pml.xsd', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'docs',
    });
    box.setHyperlink('https://example.com');
    const pkg = _internalPackageOf(pres);
    expectSchemaValid(
      decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array()),
      'pml',
    );
  });
});
