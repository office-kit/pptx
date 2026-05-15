// Slide transitions (p:transition).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  clearSlideTransition,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideTransition,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('L3: setSlideTransition', () => {
  it('emits <p:transition><p:fade/></p:transition>', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideTransition(slide, { effect: 'fade', speed: 'med' });
    const xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).toContain('<p:transition');
    expect(xml).toContain('spd="med"');
    expect(xml).toContain('<p:fade/>');
  });

  it('writes effect-specific attributes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideTransition(slide, { effect: 'push', direction: 'l', speed: 'fast' });
    const xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).toContain('<p:push dir="l"/>');
    expect(xml).toContain('spd="fast"');
  });

  it('emits advClick / advTm only when not default', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideTransition(slide, {
      effect: 'cut', advanceOnClick: false, advanceAfterMs: 3000,
    });
    const xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).toContain('advClick="0"');
    expect(xml).toContain('advTm="3000"');
  });

  it('replaces any prior transition', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideTransition(slide, { effect: 'fade' });
    setSlideTransition(getSlides(pres)[0]!, { effect: 'push', direction: 'r' });
    const xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).not.toContain('<p:fade');
    expect(xml).toContain('<p:push dir="r"/>');
    expect((xml.match(/<p:transition/g) ?? []).length).toBe(1);
  });

  it('clearSlideTransition removes the element', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideTransition(slide, { effect: 'fade' });
    clearSlideTransition(getSlides(pres)[0]!);
    const xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).not.toContain('<p:transition');
  });

  it('preserves transition through save/load', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    setSlideTransition(getSlides(pres)[0]!, { effect: 'wipe', direction: 'd' });
    setSlideTransition(getSlides(pres)[1]!, { effect: 'dissolve', speed: 'slow' });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const xml0 = getSlideXmlString(getSlides(reloaded)[0]!);
    expect(xml0).toContain('<p:wipe dir="d"/>');
  });

  skipIfNoXmllint('transition XML validates against pml.xsd', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideTransition(slide, { effect: 'fade', speed: 'med' });
    expectSchemaValid(getSlideXmlString(getSlides(pres)[0]!), 'pml');
  });
});
