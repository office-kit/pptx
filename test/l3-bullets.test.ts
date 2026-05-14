// Bullet and numbered list authoring.

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

describe('L3: bullet / numbered lists', () => {
  it('emits <a:buChar char="•"/> for bullet style', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(3),
      text: 'Apples\nOranges\nBananas',
    });
    box.setText('Apples\nOranges\nBananas', { bullets: 'bullet' });

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    // One <a:buChar> per paragraph (3 paragraphs).
    expect((xml.match(/<a:buChar/g) ?? []).length).toBe(3);
    expect(xml).toContain('char="•"');
  });

  it('emits <a:buAutoNum type="arabicPeriod"/> for numbered style', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      text: 'a\nb',
    });
    box.setText('a\nb', { bullets: 'number' });

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect((xml.match(/<a:buAutoNum/g) ?? []).length).toBe(2);
    expect(xml).toContain('type="arabicPeriod"');
  });

  it('accepts a custom bullet character', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      text: 'one\ntwo',
    });
    box.setText('one\ntwo', { bullets: { char: '◆' } });
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('char="◆"');
  });

  it('setBullets after setText applies bullets without touching text', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      text: 'Already there',
    });
    box.setText('First\nSecond');
    box.setBullets('bullet');
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('First');
    expect(xml).toContain('Second');
    expect((xml.match(/<a:buChar/g) ?? []).length).toBe(2);
  });

  it("setBullets('none') forces explicit bullet-free paragraphs", async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      text: 'plain text',
    });
    box.setText('a\nb');
    box.setBullets('bullet');
    box.setBullets('none');
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).not.toContain('<a:buChar');
    expect((xml.match(/<a:buNone\/>/g) ?? []).length).toBe(2);
  });

  skipIfNoXmllint('bullet output validates against pml.xsd', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(3),
      text: 'a\nb\nc',
    });
    box.setText('Apples\nOranges\nBananas', { bullets: 'bullet' });
    const pkg = _internalPackageOf(pres);
    expectSchemaValid(
      decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array()),
      'pml',
    );
  });
});
