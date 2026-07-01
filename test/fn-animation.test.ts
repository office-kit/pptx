// Free-function animation API (v1 — single click-effect per slide).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  clearSlideAnimations,
  createPresentation,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeAnimation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: animations (v1 — single click-effect)', () => {
  it('setShapeAnimation fadeIn adds a click-triggered entrance effect', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeAnimation(shape, { effect: 'fadeIn' });

    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<p:timing>');
    expect(xml).toContain('presetID="10"');
    expect(xml).toContain('presetClass="entr"');
    expect(xml).toContain('nodeType="clickEffect"');
    expect(xml).toContain('<p:bldLst>');
    expect(xml).toContain('style.opacity');
  });

  it('setShapeAnimation appear emits the entrance preset without opacity anim', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeAnimation(shape, { effect: 'appear' });

    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('presetID="1"');
    expect(xml).toContain('presetClass="entr"');
    // `appear` is instantaneous — no opacity tween.
    expect(xml).not.toContain('style.opacity');
    expect(xml).toContain('style.visibility');
  });

  it('setShapeAnimation fadeOut emits an exit effect', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeAnimation(shape, { effect: 'fadeOut' });

    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('presetClass="exit"');
    expect(xml).toContain('val="hidden"');
  });

  it('setShapeAnimation merges a second effect into the existing timing', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeAnimation(shape, { effect: 'fadeIn' });
    setShapeAnimation(shape, { effect: 'fadeOut', durationMs: 1000 });

    const xml = await slideXml(await savePresentation(pres), 0);
    // A second call merges rather than replaces: still a single <p:timing>, but
    // both the entrance and the exit effect survive (a replace implementation
    // would drop the fadeIn). Two build entries, both preset classes present.
    expect(xml.match(/<p:timing>/g)?.length).toBe(1);
    expect(xml).toContain('presetClass="entr"');
    expect(xml).toContain('presetClass="exit"');
    expect((xml.match(/<p:bldP/g) ?? []).length).toBe(2);
    expect(xml).toContain('dur="1000"');
  });

  it('keeps grpId unique across several animated shapes', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    for (let i = 0; i < 3; i++) {
      const shape = addSlideShape(slide, {
        preset: 'rect',
        x: inches(1 + i),
        y: inches(1),
        w: inches(1),
        h: inches(1),
        text: `s${i}`,
      });
      setShapeAnimation(shape, { effect: 'fadeIn' });
    }
    const xml = getSlideXmlString(getSlides(pres)[0]!);
    const grpIds = [...xml.matchAll(/grpId="(\d+)"/g)].map((m) => m[1]);
    // Each shape's build group must carry a distinct grpId — a duplicate would
    // make PowerPoint fold two shapes' effects into one group.
    const bldGrpIds = [...xml.matchAll(/<p:bldP[^>]*grpId="(\d+)"/g)].map((m) => m[1]);
    expect(bldGrpIds.length).toBe(3);
    expect(new Set(bldGrpIds).size).toBe(3);
    expect(grpIds.length).toBeGreaterThan(0);
  });

  it('clearSlideAnimations removes the timing block', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeAnimation(shape, { effect: 'fadeIn' });
    expect(await slideXml(await savePresentation(pres), 0)).toContain('<p:timing>');

    clearSlideAnimations(slide);
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('<p:timing>');
  });
});
