// Shape outline (a:ln) authoring.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation, inches, pt } from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/presentation.ts';
import { partName } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('L3: SlideShape.setStroke', () => {
  it('sets line color and width', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Bordered',
    });
    box.setStroke({ color: '#FF8800', widthEmu: pt(2) });

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('<a:ln');
    expect(xml).toContain('val="FF8800"');
    expect(xml).toMatch(/<a:ln[^>]*w="25400"/); // 2pt = 25400 EMU
  });

  it('toggles to noStroke', async () => {
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
    box.setStroke({ color: '#000000', widthEmu: pt(1) });
    box.setNoStroke();
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('<a:ln');
    expect(xml).toContain('<a:noFill/>');
  });

  skipIfNoXmllint('outlined shape validates against pml.xsd', async () => {
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
    box.setStroke({ color: 'accent3', widthEmu: pt(1.5) });
    box.setFill('#FFFFFF');
    box.setTextFormat({ color: '#222222', size: 14 });

    const pkg = _internalPackageOf(pres);
    expectSchemaValid(
      decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array()),
      'pml',
    );
  });
});
