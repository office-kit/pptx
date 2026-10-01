// How a clip plays: the `<p:cMediaNode>` attributes and the start condition
// of its time node. PowerPoint's Playback tab, as far as the core schema
// states it — trimming lives in a 2010 extension and is not covered.

import { describe, expect, it } from 'vitest';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import {
  addBlankSlide,
  addSlideMedia,
  addSlideShape,
  createPresentation,
  getShapeMediaPlayback,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  type PresentationData,
  readPackagePart,
  savePresentation,
  setShapeMediaPlayback,
  type SlideShapeData,
  _internalPackageOf,
} from '../src/api/index.ts';
import { partName } from '../src/internal/opc/index.ts';

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;
const ascii = (s: string): number[] => Array.from(s, (ch) => ch.charCodeAt(0));
const mp4 = (): Uint8Array =>
  new Uint8Array([0, 0, 0, 0x18, ...ascii('ftypmp42'), 0, 0, 0, 0, ...ascii('mp42isom')]);
const mp3 = (): Uint8Array => new Uint8Array([...ascii('ID3'), 3, 0, 0, 0, 0, 0, 0, 0xff, 0xfb]);
const box = { x: inches(1), y: inches(1), w: inches(4), h: inches(2.25) };

const deckWith = (kind: 'video' | 'audio') => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideMedia(
    slide,
    kind === 'video'
      ? { kind: 'video', data: mp4(), ...box }
      : { kind: 'audio', data: mp3(), ...box },
  );
  return { pres, slide, shape };
};

const slideXml = (pres: PresentationData): string =>
  new TextDecoder().decode(readPackagePart(pres, '/ppt/slides/slide1.xml')!);

describe('media playback', () => {
  it('reports what a freshly added clip does: click to play, full volume kept', () => {
    const { shape } = deckWith('video');

    expect(getShapeMediaPlayback(shape)).toEqual({
      autoplay: false,
      loop: false,
      volume: 0.8,
      muted: false,
      fullScreen: false,
      hideWhenStopped: false,
    });
  });

  it.each(['onClick', 'onNext', 'onBegin'])(
    'does not report %s-triggered media as slide autoplay',
    async (event) => {
      const { pres } = deckWith('video');
      const slidePart = _internalPackageOf(pres).getPart(partName('/ppt/slides/slide1.xml'))!;
      slidePart.data = new TextEncoder().encode(
        new TextDecoder()
          .decode(slidePart.data)
          .replace('<p:cond delay="indefinite"/>', `<p:cond evt="${event}" delay="0"/>`),
      );
      const reloaded = await loadPresentation(await savePresentation(pres));
      const shape = getSlideShapes(getSlides(reloaded)[0]!)[0]!;

      expect(getShapeMediaPlayback(shape)?.autoplay).toBe(false);
    },
  );

  it('reports nothing for a shape that is not a clip', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape: SlideShapeData = addSlideShape(slide, { preset: 'rect', ...box });

    expect(getShapeMediaPlayback(shape)).toBeNull();
  });

  it('starts with the slide, repeats, and dims down', () => {
    const { pres, shape } = deckWith('video');

    setShapeMediaPlayback(shape, { autoplay: true, loop: true, volume: 0.25 });

    expect(getShapeMediaPlayback(shape)).toMatchObject({
      autoplay: true,
      loop: true,
      volume: 0.25,
    });
    const xml = slideXml(pres);
    expect(xml).toContain('<p:cond delay="0"/>');
    expect(xml).toContain('repeatCount="indefinite"');
    expect(xml).toContain('vol="25000"');
  });

  it('reads and writes an automatic start delay in milliseconds', async () => {
    const { pres, shape } = deckWith('video');

    setShapeMediaPlayback(shape, { autoplay: true, delayMs: 1250 });
    expect(getShapeMediaPlayback(shape)).toMatchObject({ autoplay: true, delayMs: 1250 });
    expect(slideXml(pres)).toContain('<p:cond delay="1250"/>');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const saved = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapeMediaPlayback(saved)).toMatchObject({ autoplay: true, delayMs: 1250 });
  });

  it('preserves an automatic delay when autoplay is reaffirmed', () => {
    const { shape } = deckWith('video');

    setShapeMediaPlayback(shape, { autoplay: true, delayMs: 1250 });
    setShapeMediaPlayback(shape, { autoplay: true });

    expect(getShapeMediaPlayback(shape)).toMatchObject({ autoplay: true, delayMs: 1250 });
  });

  it('starts automatic playback immediately when switching from click playback', () => {
    const { pres, shape } = deckWith('video');

    setShapeMediaPlayback(shape, { autoplay: true });

    expect(getShapeMediaPlayback(shape)).toMatchObject({ autoplay: true });
    expect(getShapeMediaPlayback(shape)).not.toHaveProperty('delayMs');
    expect(slideXml(pres)).toContain('<p:cond delay="0"/>');
  });

  it('reads and writes playback across slides through cMediaNode numSld', async () => {
    const { pres, shape } = deckWith('audio');
    expect(getShapeMediaPlayback(shape)).not.toHaveProperty('slideCount');

    setShapeMediaPlayback(shape, { slideCount: 3 });
    expect(getShapeMediaPlayback(shape)).toMatchObject({ slideCount: 3 });
    expect(slideXml(pres)).toContain('numSld="3"');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const loadedShape = getSlideShapes(getSlides(reloaded)[0]!).at(0)!;
    expect(getShapeMediaPlayback(loadedShape)).toMatchObject({ slideCount: 3 });

    setShapeMediaPlayback(loadedShape, { slideCount: 1 });
    expect(getShapeMediaPlayback(loadedShape)).not.toHaveProperty('slideCount');
    expect(slideXml(reloaded)).not.toContain('numSld=');
  });

  it.each([
    -1,
    1.5,
    0x100000000,
    Number.MAX_SAFE_INTEGER + 1,
    Number.POSITIVE_INFINITY,
    Number.NaN,
  ])('rejects invalid slide counts before changing the clip (%s)', (slideCount) => {
    const { pres, shape } = deckWith('video');
    const before = slideXml(pres);

    expect(() => setShapeMediaPlayback(shape, { volume: 0.2, slideCount })).toThrow(/slideCount/);
    expect(slideXml(pres)).toBe(before);
  });

  skipIfNoXmllint('writes schema-valid unsigned 32-bit slide-count boundaries', () => {
    const { pres, shape } = deckWith('audio');
    for (const slideCount of [0, 0xffffffff]) {
      setShapeMediaPlayback(shape, { slideCount });
      expect(getShapeMediaPlayback(shape)).toMatchObject({ slideCount });
      expectSchemaValid(slideXml(pres), 'pml');
    }
  });

  it('drops an event trigger when switching event playback to automatic playback', async () => {
    const { pres } = deckWith('video');
    const slidePart = _internalPackageOf(pres).getPart(partName('/ppt/slides/slide1.xml'))!;
    slidePart.data = new TextEncoder().encode(
      new TextDecoder()
        .decode(slidePart.data)
        .replace('<p:cond delay="indefinite"/>', '<p:cond evt="onClick" delay="0"/>'),
    );
    const reloaded = await loadPresentation(await savePresentation(pres));
    const loadedShape = getSlideShapes(getSlides(reloaded)[0]!)[0]!;

    setShapeMediaPlayback(loadedShape, { autoplay: true });

    expect(getShapeMediaPlayback(loadedShape)).toMatchObject({ autoplay: true });
    expect(getShapeMediaPlayback(loadedShape)).not.toHaveProperty('delayMs');
    expect(
      new TextDecoder().decode(
        _internalPackageOf(reloaded).getPart(partName('/ppt/slides/slide1.xml'))!.data,
      ),
    ).toContain('<p:cond delay="0"/>');
  });

  it('preserves additional start conditions when changing an automatic delay', async () => {
    const { pres } = deckWith('video');
    const part = _internalPackageOf(pres).getPart(partName('/ppt/slides/slide1.xml'))!;
    part.data = new TextEncoder().encode(
      new TextDecoder()
        .decode(part.data)
        .replace(
          '<p:cond delay="indefinite"/>',
          '<p:cond delay="100"/><p:cond evt="onClick" delay="0"/>',
        ),
    );
    const reloaded = await loadPresentation(await savePresentation(pres));
    const shape = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    setShapeMediaPlayback(shape, { delayMs: 250 });
    expect(slideXml(reloaded)).toContain('<p:cond delay="250"/><p:cond evt="onClick" delay="0"/>');
  });

  it('rejects delay changes unless the clip already starts automatically', () => {
    const { pres, shape } = deckWith('video');
    const before = slideXml(pres);

    expect(() => setShapeMediaPlayback(shape, { delayMs: 250 })).toThrow(/autoplay/);
    expect(() => setShapeMediaPlayback(shape, { autoplay: false, delayMs: 250 })).toThrow(
      /delayMs/,
    );
    expect(slideXml(pres)).toBe(before);
  });

  it.each([-1, 1.5, Number.POSITIVE_INFINITY, Number.NaN])(
    'rejects invalid automatic delays before changing the clip (%s)',
    (delayMs) => {
      const { pres, shape } = deckWith('video');
      setShapeMediaPlayback(shape, { autoplay: true });
      const before = slideXml(pres);

      expect(() => setShapeMediaPlayback(shape, { delayMs })).toThrow(/delayMs/);
      expect(slideXml(pres)).toBe(before);
    },
  );

  it('goes back to waiting for a click, and stops repeating', () => {
    const { pres, shape } = deckWith('audio');
    setShapeMediaPlayback(shape, { autoplay: true, loop: true });

    setShapeMediaPlayback(shape, { autoplay: false, loop: false });

    expect(getShapeMediaPlayback(shape)).toMatchObject({ autoplay: false, loop: false });
    expect(slideXml(pres)).toContain('<p:cond delay="indefinite"/>');
    // The attribute is dropped rather than set to 1 — "no repeat" is the
    // schema's own default.
    expect(slideXml(pres)).not.toContain('repeatCount');
  });

  it('keeps the properties it was not asked about', () => {
    const { shape } = deckWith('video');
    setShapeMediaPlayback(shape, { volume: 0.3, muted: true, hideWhenStopped: true });

    setShapeMediaPlayback(shape, { autoplay: true });

    expect(getShapeMediaPlayback(shape)).toEqual({
      autoplay: true,
      loop: false,
      volume: 0.3,
      muted: true,
      fullScreen: false,
      hideWhenStopped: true,
    });
  });

  it.each([1.5, -0.1, NaN, Infinity, -Infinity])(
    'refuses invalid volume %s before changing anything',
    (volume) => {
      const { pres, shape } = deckWith('video');
      const before = slideXml(pres);
      expect(() => setShapeMediaPlayback(shape, { volume, muted: true })).toThrow(
        /between 0 and 1/,
      );
      expect(slideXml(pres)).toBe(before);
      expect(getShapeMediaPlayback(shape)?.volume).toBe(0.8);
    },
  );

  it('refuses full screen on an audio clip, which has no such attribute', () => {
    const { shape } = deckWith('audio');

    expect(() => setShapeMediaPlayback(shape, { fullScreen: true })).toThrow(/video only/);
  });

  it('plays a video full screen when asked', () => {
    const { pres, shape } = deckWith('video');

    setShapeMediaPlayback(shape, { fullScreen: true });

    expect(getShapeMediaPlayback(shape)?.fullScreen).toBe(true);
    expect(slideXml(pres)).toContain('fullScrn="1"');
  });

  it('refuses a shape with no media time node', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const shape = addSlideShape(slide, { preset: 'rect', ...box });

    expect(() => setShapeMediaPlayback(shape, { loop: true })).toThrow(/no media time node/);
  });

  it('survives the save/load round trip', async () => {
    const { pres, shape } = deckWith('video');
    setShapeMediaPlayback(shape, {
      autoplay: true,
      delayMs: 400,
      loop: true,
      volume: 0.4,
      muted: true,
      fullScreen: true,
      hideWhenStopped: true,
    });
    const before = getShapeMediaPlayback(shape);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const saved = getSlideShapes(getSlides(reloaded)[0]!)[0]!;

    expect(getShapeMediaPlayback(saved)).toEqual(before);
  });

  skipIfNoXmllint('the written timing tree validates', async () => {
    const { pres, shape } = deckWith('video');
    setShapeMediaPlayback(shape, {
      autoplay: true,
      delayMs: 400,
      loop: true,
      volume: 0.4,
      muted: true,
      fullScreen: true,
      hideWhenStopped: true,
    });
    const saved = await loadPresentation(await savePresentation(pres));

    expectSchemaValid(slideXml(saved), 'pml');
  });
});
