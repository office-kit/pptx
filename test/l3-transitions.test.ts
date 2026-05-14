// Slide transitions (p:transition).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/api/index.ts';
import { _internalPackageOf } from '../src/api/presentation.ts';
import { partName } from '../src/internal/opc/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('L3: Slide.setTransition', () => {
  it('emits <p:transition><p:fade/></p:transition>', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setTransition({ effect: 'fade', speed: 'med' });

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(slide._partName)?.data ?? new Uint8Array());
    expect(xml).toContain('<p:transition');
    expect(xml).toContain('spd="med"');
    expect(xml).toContain('<p:fade/>');
  });

  it('writes effect-specific attributes', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setTransition({ effect: 'push', direction: 'l', speed: 'fast' });
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(slide._partName)?.data ?? new Uint8Array());
    expect(xml).toContain('<p:push dir="l"/>');
    expect(xml).toContain('spd="fast"');
  });

  it('emits advClick / advTm only when not default', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setTransition({
      effect: 'cut',
      advanceOnClick: false,
      advanceAfterMs: 3000,
    });
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(slide._partName)?.data ?? new Uint8Array());
    expect(xml).toContain('advClick="0"');
    expect(xml).toContain('advTm="3000"');
  });

  it('replaces any prior transition', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setTransition({ effect: 'fade' });
    slide.setTransition({ effect: 'push', direction: 'r' });

    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(slide._partName)?.data ?? new Uint8Array());
    expect(xml).not.toContain('<p:fade');
    expect(xml).toContain('<p:push dir="r"/>');
    // Exactly one transition element on the slide.
    expect((xml.match(/<p:transition/g) ?? []).length).toBe(1);
  });

  it('clearTransition removes the element', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setTransition({ effect: 'fade' });
    slide.clearTransition();
    const pkg = _internalPackageOf(pres);
    const xml = decode(pkg.getPart(slide._partName)?.data ?? new Uint8Array());
    expect(xml).not.toContain('<p:transition');
  });

  it('preserves transition through save/load', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    pres.slides[0]?.setTransition({ effect: 'wipe', direction: 'd' });
    pres.slides[1]?.setTransition({ effect: 'dissolve', speed: 'slow' });
    const reloaded = await Presentation.load(await pres.save());
    const pkg = _internalPackageOf(reloaded);
    const xml0 = decode(
      pkg.getPart(partName(reloaded.slides[0]?._partName ?? '/x'))?.data ?? new Uint8Array(),
    );
    expect(xml0).toContain('<p:wipe dir="d"/>');
  });

  skipIfNoXmllint('transition XML validates against pml.xsd', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected slide');
    slide.setTransition({ effect: 'fade', speed: 'med' });
    const pkg = _internalPackageOf(pres);
    expectSchemaValid(decode(pkg.getPart(slide._partName)?.data ?? new Uint8Array()), 'pml');
  });
});
