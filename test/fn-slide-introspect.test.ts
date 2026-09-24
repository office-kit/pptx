// Slide-level introspection: `getSlideTransition` + `getSlideBackground`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  createPresentation,
  savePresentation,
  clearSlideBackground,
  clearSlideTransition,
  getSlideBackground,
  getSlideTransition,
  getSlides,
  loadPresentation,
  setSlideBackground,
  setSlideTransition,
  setSlideAdvanceTiming,
} from '../src/api/index.ts';

import { SLIDE_DOCUMENT } from '../src/api/_internal-symbols.ts';
import { NS, parseXml, type XmlElement } from '../src/internal/xml/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: slide introspection', () => {
  it('getSlideTransition returns null before any set, and the configured effect after', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideTransition(slide)).toBeNull();

    setSlideTransition(slide, { effect: 'fade', speed: 'fast' });
    const got = getSlideTransition(slide);
    expect(got?.effect).toBe('fade');
    expect(got?.speed).toBe('fast');

    clearSlideTransition(slide);
    expect(getSlideTransition(slide)).toBeNull();
  });

  it('getSlideBackground reports inherit when no <p:bg> is set', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideBackground(slide).kind).toBe('inherit');
  });

  it('getSlideBackground reads back a solid color after setSlideBackground', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideBackground(slide, '#ABCDEF');
    expect(getSlideBackground(slide)).toEqual({ kind: 'solid', color: '#ABCDEF' });

    clearSlideBackground(slide);
    expect(getSlideBackground(slide).kind).toBe('inherit');
  });
});

it('retains automatic timing for transitions without visual effects', async () => {
  const p = createPresentation();
  const slide = addBlankSlide(p);
  const settings = { effect: 'none', advanceOnClick: false, advanceAfterMs: 3000 };
  setSlideTransition(slide, settings);
  expect(getSlideTransition(slide)).toEqual(settings);
  const loaded = await loadPresentation(await savePresentation(p));
  expect(getSlideTransition(getSlides(loaded)[0]!)).toEqual(settings);
});

it('reads XML booleans and does not mistake sound actions for visual effects', () => {
  const p = createPresentation();
  const slide = addBlankSlide(p);
  const children = slide[SLIDE_DOCUMENT].root.children;
  const node = parseXml(
    `<p:transition xmlns:p="${NS.pml}" advClick="false" advTm="2500"><p:sndAc><p:endSnd/></p:sndAc><p:extLst/></p:transition>`,
  ).root;
  children.push(node);
  expect(getSlideTransition(slide)).toEqual({
    effect: 'none',
    advanceOnClick: false,
    advanceAfterMs: 2500,
  });
  node.children.unshift(parseXml(`<p:fade xmlns:p="${NS.pml}" thruBlk="true"/>`).root);
  expect(getSlideTransition(slide)).toEqual({
    effect: 'fade',
    thruBlack: true,
    advanceOnClick: false,
    advanceAfterMs: 2500,
  });
});

it('edits advance timing without replacing transition effects, sounds or extensions', async () => {
  const p = createPresentation();
  const slide = addBlankSlide(p);
  const node = parseXml(
    `<p:transition xmlns:p="${NS.pml}" spd="slow"><p:fade thruBlk="1"/><p:sndAc><p:endSnd/></p:sndAc><p:extLst><p:ext uri="timing-test"/></p:extLst></p:transition>`,
  ).root;
  slide[SLIDE_DOCUMENT].root.children.push(node);
  const children = JSON.stringify(node.children);
  setSlideAdvanceTiming(slide, { advanceOnClick: false, advanceAfterMs: 1250 });
  expect(
    JSON.stringify(
      slide[SLIDE_DOCUMENT].root.children.find(
        (child): child is XmlElement =>
          child.kind === 'element' && child.name.localName === 'transition',
      )?.children,
    ),
  ).toBe(children);
  const loaded = await loadPresentation(await savePresentation(p));
  expect(getSlideTransition(getSlides(loaded)[0]!)).toEqual({
    effect: 'fade',
    speed: 'slow',
    thruBlack: true,
    advanceOnClick: false,
    advanceAfterMs: 1250,
  });
  setSlideAdvanceTiming(slide, { advanceOnClick: true, advanceAfterMs: null });
  expect(getSlideTransition(slide)?.advanceAfterMs).toBeUndefined();
  expect(getSlideTransition(slide)?.effect).toBe('fade');
  const before = await savePresentation(p);
  for (const advanceAfterMs of [-1, 1.5, NaN, 4294967296, undefined]) {
    expect(() =>
      setSlideAdvanceTiming(slide, {
        advanceOnClick: false,
        advanceAfterMs: advanceAfterMs as number,
      }),
    ).toThrow();
    expect(await savePresentation(p)).toEqual(before);
  }
});
