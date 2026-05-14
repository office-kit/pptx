// Shape rotation + flip.

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

describe('L3: Shape rotation', () => {
  it('writes rot="5400000" for setRotation(90)', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      text: '↑',
    });
    box.setRotation(90);

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('rot="5400000"');
    expect(box.rotation).toBe(90);
  });

  it('normalizes out-of-range values', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      text: 'x',
    });
    box.setRotation(-90);
    expect(box.rotation).toBe(270);

    box.setRotation(540); // 540 % 360 = 180
    expect(box.rotation).toBe(180);
  });

  it('clears rotation when set to 0', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      text: 'x',
    });
    box.setRotation(45);
    box.setRotation(0);

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).not.toContain('rot=');
  });
});

describe('L3: Shape flip', () => {
  it('writes flipH="1" / flipV="1"', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      text: 'x',
    });
    box.setFlip({ horizontal: true, vertical: true });

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toMatch(/flipH="1"/);
    expect(xml).toMatch(/flipV="1"/);
    expect(box.flip).toEqual({ horizontal: true, vertical: true });
  });

  it('clears individual flip flags', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      text: 'x',
    });
    box.setFlip({ horizontal: true, vertical: true });
    box.setFlip({ horizontal: false });
    expect(box.flip).toEqual({ horizontal: false, vertical: true });
  });

  skipIfNoXmllint('rotated+flipped shape validates against pml.xsd', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const box = slide.addTextBox({
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(2),
      text: 'x',
    });
    box.setRotation(45);
    box.setFlip({ horizontal: true });

    const pkg = _internalPackageOf(pres);
    expectSchemaValid(
      decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array()),
      'pml',
    );
  });
});
