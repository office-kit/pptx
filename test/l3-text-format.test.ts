// Text run formatting: font, size, color, bold, italic, underline.

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

describe('L3: SlideShape.setTextFormat', () => {
  it('sets size, bold, italic, underline, font, and color on all runs', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Styled text',
    });
    box.setTextFormat({
      font: 'Calibri',
      size: 24,
      color: '#3366CC',
      bold: true,
      italic: true,
      underline: true,
    });

    const pkg = _internalPackageOf(pres);
    const slidePart = pkg.getPart(partName('/ppt/slides/slide1.xml'));
    const xml = decode(slidePart?.data ?? new Uint8Array());

    // The serialized run-properties carry every attribute we set.
    expect(xml).toContain('sz="2400"');
    expect(xml).toContain('b="1"');
    expect(xml).toContain('i="1"');
    expect(xml).toContain('u="sng"');
    expect(xml).toContain('<a:latin typeface="Calibri"/>');
    expect(xml).toContain('val="3366CC"');

    // Round-trip preserves the formatting.
    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides[0]?.text).toContain('Styled text');
  });

  it('partial updates compose without losing prior formatting', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Compose',
    });
    box.setTextFormat({ bold: true });
    box.setTextFormat({ color: '#FF0000' });

    const pkg = _internalPackageOf(pres);
    const slidePart = pkg.getPart(partName('/ppt/slides/slide1.xml'));
    const xml = decode(slidePart?.data ?? new Uint8Array());
    expect(xml).toContain('b="1"');
    expect(xml).toContain('val="FF0000"');
  });

  it('accepts scheme-color tokens', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Theme color',
    });
    box.setTextFormat({ color: 'accent1' });
    const pkg = _internalPackageOf(pres);
    const slidePart = pkg.getPart(partName('/ppt/slides/slide1.xml'));
    const xml = decode(slidePart?.data ?? new Uint8Array());
    expect(xml).toContain('<a:schemeClr val="accent1"/>');
  });

  it('rejects unrecognized color values', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 't',
    });
    expect(() => box.setTextFormat({ color: 'pumpkin spice' })).toThrow(/color/);
  });

  skipIfNoXmllint('formatted output validates against pml.xsd', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Validates with format',
    });
    box.setTextFormat({
      font: 'Helvetica',
      size: 18.5,
      color: '#222222',
      bold: true,
      italic: false,
      underline: 'dbl',
    });
    const pkg = _internalPackageOf(pres);
    const slidePart = pkg.getPart(partName('/ppt/slides/slide1.xml'));
    expectSchemaValid(decode(slidePart?.data ?? new Uint8Array()), 'pml');
  });

  it('applies formatting before setText and the format survives setText', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Initial',
    });
    box.setTextFormat({ bold: true, color: '#00AA00' });
    box.setText('Replaced text'); // setText clones the first rPr per the
    // existing setText contract, so formatting carries over.

    const pkg = _internalPackageOf(pres);
    const slidePart = pkg.getPart(partName('/ppt/slides/slide1.xml'));
    const xml = decode(slidePart?.data ?? new Uint8Array());
    expect(xml).toContain('b="1"');
    expect(xml).toContain('val="00AA00"');
    expect(xml).toContain('Replaced text');
  });
});
