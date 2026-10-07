// Transition duration (PowerPoint 2010's p14:dur in mc:AlternateContent) and
// transition sounds (p:sndAc).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { INTERNAL_PACKAGE, SLIDE_PART_NAME } from '../src/api/_internal-symbols.ts';
import {
  clearSlideTransition,
  getSlideTransition,
  getSlideTransitionSound,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideTransition,
  setSlideTransitionSound,
} from '../src/api/index.ts';
import { REL_TYPES } from '../src/internal/presentationml/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = fileURLToPath(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url));
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

// The smallest RIFF/WAVE header with an empty data chunk.
const wav = (): Uint8Array => {
  const bytes = new Uint8Array(44);
  const view = new DataView(bytes.buffer);
  const ascii = (at: number, text: string) =>
    [...text].forEach((c, i) => view.setUint8(at + i, c.charCodeAt(0)));
  ascii(0, 'RIFF');
  view.setUint32(4, 36, true);
  ascii(8, 'WAVE');
  ascii(12, 'fmt ');
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, 1, true);
  view.setUint32(24, 8000, true);
  view.setUint32(28, 8000, true);
  view.setUint16(32, 1, true);
  view.setUint16(34, 8, true);
  ascii(36, 'data');
  return bytes;
};

// xmllint validates against ECMA-376 alone, which has no markup-compatibility
// processing; validate what a strict consumer reads (the fallback).
const resolveFallback = (xml: string): string =>
  xml.replace(
    /<mc:AlternateContent\b[^>]*>[\s\S]*?<mc:Fallback>([\s\S]*?)<\/mc:Fallback><\/mc:AlternateContent>/g,
    '$1',
  );

const firstSlide = async () => {
  const pres = await loadPresentation(await readFile(fixture));
  return { pres, slide: getSlides(pres)[0]! };
};
const audioRels = (slide: ReturnType<typeof getSlides>[number]) =>
  slide[INTERNAL_PACKAGE]
    .getRels(slide[SLIDE_PART_NAME])!
    .items.filter((rel) => rel.type === REL_TYPES.audio);

describe('transition duration', () => {
  it('writes p14:dur with a fallback speed and reads it back after save', async () => {
    const { pres, slide } = await firstSlide();
    setSlideTransition(slide, { effect: 'fade', durationMs: 2000 });
    const xml = getSlideXmlString(slide);
    expect(xml).toContain('<mc:Choice xmlns:p14=');
    expect(xml).toContain('p14:dur="2000"');
    expect(xml).toMatch(/<mc:Fallback><p:transition spd="slow"><p:fade\/><\/p:transition>/);
    const reloaded = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideTransition(reloaded)).toEqual({
      effect: 'fade',
      speed: 'slow',
      durationMs: 2000,
    });
  });

  it('replaces an AlternateContent transition with a single plain one', async () => {
    const { slide } = await firstSlide();
    // PowerPoint leaves out spd when it is the default, fast (Cut at 0.1 s).
    setSlideTransition(slide, { effect: 'fade', durationMs: 400 });
    expect(getSlideTransition(slide)).toEqual({ effect: 'fade', durationMs: 400 });
    setSlideTransition(slide, { effect: 'push', direction: 'l' });
    const xml = getSlideXmlString(slide);
    expect(xml).not.toContain('AlternateContent');
    expect(xml.match(/<p:transition\b/g)).toHaveLength(1);
    clearSlideTransition(slide);
    expect(getSlideTransition(slide)).toBeNull();
  });

  // Mac PowerPoint saves Push (1 s) as `spd="slow"` alone, Uncover (0.75 s) as
  // `spd="med"` and Flash (1 s) without p14:dur inside its p14 wrapper.
  it('writes a duration equal to its speed as the speed alone', async () => {
    const { slide } = await firstSlide();
    setSlideTransition(slide, { effect: 'push', direction: 'u', durationMs: 1000 });
    expect(getSlideXmlString(slide)).toContain(
      '<p:transition spd="slow"><p:push dir="u"/></p:transition>',
    );
    expect(getSlideXmlString(slide)).not.toContain('AlternateContent');
    expect(getSlideTransition(slide)).toEqual({ effect: 'push', direction: 'u', speed: 'slow' });
    setSlideTransition(slide, { effect: 'pull', durationMs: 750 });
    expect(getSlideTransition(slide)).toEqual({ effect: 'pull', speed: 'med' });
    setSlideTransition(slide, { effect: 'flash', durationMs: 1000 });
    const xml = getSlideXmlString(slide);
    expect(xml).toContain('<p:transition spd="slow"><p14:flash/></p:transition>');
    expect(xml).not.toContain('p14:dur');
  });

  it('takes the fastest speed at least as long as the duration', async () => {
    const { slide } = await firstSlide();
    const speedOf = (durationMs: number) => {
      setSlideTransition(slide, { effect: 'fade', durationMs });
      return getSlideTransition(slide)?.speed;
    };
    // Fade 700 ms is med and Shape 800 ms slow natively: not the nearest speed.
    expect([100, 500, 501, 700, 750, 800, 900, 1500].map(speedOf)).toEqual([
      undefined,
      undefined,
      'med',
      'med',
      'med',
      'slow',
      'slow',
      'slow',
    ]);
    // An explicit speed is kept, and decides whether p14:dur is needed.
    setSlideTransition(slide, { effect: 'fade', speed: 'fast', durationMs: 500 });
    expect(getSlideTransition(slide)).toEqual({ effect: 'fade', speed: 'fast' });
    setSlideTransition(slide, { effect: 'fade', speed: 'fast', durationMs: 700 });
    expect(getSlideTransition(slide)).toEqual({ effect: 'fade', speed: 'fast', durationMs: 700 });
  });

  it('rejects a negative duration without touching the slide', async () => {
    const { slide } = await firstSlide();
    const before = getSlideXmlString(slide);
    expect(() => setSlideTransition(slide, { effect: 'fade', durationMs: -1 })).toThrow(
      /durationMs/,
    );
    expect(getSlideXmlString(slide)).toBe(before);
  });

  skipIfNoXmllint('keeps the ECMA-376 fallback schema-valid', async () => {
    const { slide } = await firstSlide();
    setSlideTransition(slide, { effect: 'wipe', direction: 'u', durationMs: 1500 });
    setSlideTransitionSound(slide, { kind: 'stop' });
    expectSchemaValid(resolveFallback(getSlideXmlString(slide)), 'pml');
  });
});

describe('transition sound', () => {
  it('embeds a WAV, keeps it when the effect changes and survives save', async () => {
    const { pres, slide } = await firstSlide();
    setSlideTransition(slide, { effect: 'fade' });
    setSlideTransitionSound(slide, { kind: 'play', data: wav(), name: 'chime.wav', loop: true });
    expect(audioRels(slide)).toHaveLength(1);
    setSlideTransition(slide, { effect: 'cut', durationMs: 1200 });
    // Both the p14 choice and the fallback carry the sound.
    expect(getSlideXmlString(slide).match(/<p:sndAc>/g)).toHaveLength(2);
    const reloaded = getSlides(await loadPresentation(await savePresentation(pres)))[0]!;
    expect(getSlideTransitionSound(reloaded)).toEqual({
      kind: 'play',
      name: 'chime.wav',
      loop: true,
    });
    expect(getSlideTransition(reloaded)).toMatchObject({ effect: 'cut', durationMs: 1200 });
  });

  it('adds an effect-less transition for a sound and releases the rel when removed', async () => {
    const { slide } = await firstSlide();
    setSlideTransitionSound(slide, { kind: 'play', data: wav(), name: 'a.wav' });
    expect(getSlideTransition(slide)).toEqual({ effect: 'none' });
    setSlideTransitionSound(slide, { kind: 'stop' });
    expect(getSlideTransitionSound(slide)).toEqual({ kind: 'stop' });
    expect(audioRels(slide)).toHaveLength(0);
    setSlideTransitionSound(slide, { kind: 'play', data: wav(), name: 'a.wav' });
    clearSlideTransition(slide);
    expect(getSlideTransitionSound(slide)).toBeNull();
    expect(audioRels(slide)).toHaveLength(0);
  });

  it('rejects audio other than WAV before touching the slide', async () => {
    const { slide } = await firstSlide();
    const before = getSlideXmlString(slide);
    const mp3 = new Uint8Array([0x49, 0x44, 0x33, 3, 0, 0, 0, 0, 0, 0]);
    expect(() =>
      setSlideTransitionSound(slide, { kind: 'play', data: mp3, name: 'a.mp3' }),
    ).toThrow(/WAV/);
    expect(getSlideXmlString(slide)).toBe(before);
  });
});
