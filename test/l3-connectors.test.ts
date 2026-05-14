// Connector lines (p:cxnSp with prstGeom prst="line").

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

describe('L3: Slide.addLine', () => {
  it('emits a p:cxnSp with prstGeom prst="line"', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    const line = slide.addLine({
      from: { x: inches(1), y: inches(1) },
      to: { x: inches(5), y: inches(3) },
      color: '#0066CC',
      widthEmu: pt(2),
    });
    expect(line.kind).toBe('connector');

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    expect(xml).toContain('<p:cxnSp>');
    expect(xml).toContain('prst="line"');
    expect(xml).toContain('val="0066CC"');
    expect(xml).toMatch(/<a:ln[^>]*w="25400"/); // 2pt
  });

  it('writes flipH / flipV when from is to the right of / below to', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    slide.addLine({
      from: { x: inches(5), y: inches(3) },
      to: { x: inches(1), y: inches(1) },
    });
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    // Bounding box is the same; orientation captured by flip attrs.
    expect(xml).toContain('flipH="1"');
    expect(xml).toContain('flipV="1"');
  });

  it('omits the ln element when no color or width is given', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    slide.addLine({
      from: { x: inches(0), y: inches(0) },
      to: { x: inches(1), y: inches(0) },
    });
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array());
    // No <a:ln> means the layout's default line color / width applies.
    expect(xml).toContain('<p:cxnSp>');
    expect(xml).toContain('prst="line"');
    // The slide we built has exactly one shape (the connector); count occurrences.
    expect((xml.match(/<a:ln/g) ?? []).length).toBe(0);
  });

  skipIfNoXmllint('connector validates against pml.xsd', async () => {
    const pres = await Presentation.load(await readFile(fixture('blank.pptx')));
    const layout = pres.slideLayouts.find((l) => l.name === 'Blank');
    if (!layout) throw new Error('expected layout');
    const slide = pres.addSlide({ layout });
    slide.addLine({
      from: { x: inches(1), y: inches(1) },
      to: { x: inches(5), y: inches(3) },
      color: 'accent1',
      widthEmu: pt(1.5),
    });
    const pkg = _internalPackageOf(pres);
    expectSchemaValid(
      decode(pkg.getPart(partName('/ppt/slides/slide1.xml'))?.data ?? new Uint8Array()),
      'pml',
    );
  });
});
