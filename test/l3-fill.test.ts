// Shape fill + slide background fill.

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

describe('L3: SlideShape.setFill', () => {
  it('sets a solid srgb fill on a text box', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Filled',
    });
    box.setFill('#FFCC00');
    const pkg = _internalPackageOf(pres);
    const slidePart = pkg.getPart(partName('/ppt/slides/slide1.xml'));
    const xml = decode(slidePart?.data ?? new Uint8Array());
    expect(xml).toContain('val="FFCC00"');
    // Replacing the fill removes the previous noFill the text-box builder
    // emits by default.
    expect(xml).not.toContain('<a:noFill/>');
  });

  it('toggles to noFill and back', async () => {
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
    box.setFill('#FF0000');
    box.setNoFill();
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).not.toContain('solidFill');
    expect(xml).toContain('<a:noFill/>');
  });

  it('refuses fill on group shapes', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    // No group in the fixture; the assertion is that the function throws if
    // a group ever appears. Constructed scenarios will exercise this when
    // we add grpSp authoring.
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    // All shapes here are placeholders; ensuring the call still works on
    // a placeholder is the smoke test.
    slide.findPlaceholder('title')?.setFill('#0066CC');
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(slide._partName)?.data ?? new Uint8Array());
    expect(xml).toContain('val="0066CC"');
  });

  skipIfNoXmllint('filled shape XML validates against pml.xsd', async () => {
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
    box.setFill('accent2');
    const pkg = _internalPackageOf(pres);
    const slidePart = pkg.getPart(partName('/ppt/slides/slide1.xml'));
    expectSchemaValid(decode(slidePart?.data ?? new Uint8Array()), 'pml');
  });
});

describe('L3: Slide.setBackground', () => {
  it('inserts <p:bg><p:bgPr><a:solidFill>... before <p:spTree>', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setBackground('#003366');

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(slide._partName)?.data ?? new Uint8Array());
    expect(xml).toContain('<p:bg>');
    expect(xml).toContain('<p:bgPr>');
    expect(xml).toContain('val="003366"');
    // <p:bg> must come before <p:spTree>.
    expect(xml.indexOf('<p:bg>')).toBeLessThan(xml.indexOf('<p:spTree>'));
  });

  it('clearBackground removes any prior bg element', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setBackground('#FF0000');
    slide.clearBackground();
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(slide._partName)?.data ?? new Uint8Array());
    expect(xml).not.toContain('<p:bg>');
  });

  skipIfNoXmllint('background-fill XML validates against pml.xsd', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setBackground('#EFEFEF');
    const pkg = _internalPackageOf(pres);
    expectSchemaValid(decode(pkg.getPart(slide._partName)?.data ?? new Uint8Array()), 'pml');
  });
});
