import { expect, it } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import {
  addBlankSlide,
  applySlideBackgroundToAll,
  copySlideBackground,
  createPresentation,
  getMediaParts,
  getSlides,
  getSlideXmlString,
  getSlideBackgroundImageBytes,
  getSlideBackgroundImageCrop,
  getSlideBackgroundImageOpacity,
  getSlideBackgroundImageFillLayout,
  loadPresentation,
  savePresentation,
  setSlideBackground,
  setSlideBackgroundImage,
  setSlideBackgroundImageFillLayout,
  setSlideBackgroundImageOpacity,
} from '../src/api/index.ts';
import { buildPng } from './lib/build-png.ts';

async function sourceDeck(broken = false) {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  setSlideBackgroundImage(
    slide,
    buildPng(4, 4, () => [255, 0, 0]),
  );
  setSlideBackgroundImageOpacity(slide, 0.4);
  setSlideBackgroundImageFillLayout(slide, { mode: 'tile', scaleX: 0.5, flip: 'xy' });
  const zip = unzipSync(await savePresentation(pres));
  for (const [name, bytes] of Object.entries(zip)) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue;
    let xml = strFromU8(bytes).replace(
      '</a:blip>',
      '<a:grayscl/></a:blip><a:srcRect l="-10000" r="25000"/>',
    );
    if (broken) xml = xml.replace(/r:embed="[^"]+"/, 'r:embed="missing"');
    zip[name] = strToU8(xml);
  }
  return loadPresentation(zipSync(zip));
}

it('copies inherited picture XML across decks and restores it independently of later edits', async () => {
  const source = await sourceDeck();
  const slide = getSlides(source)[0]!;
  const bytes = getSlideBackgroundImageBytes(slide);
  applySlideBackgroundToAll(source, slide);
  const snapshotDeck = createPresentation();
  const snapshot = addBlankSlide(snapshotDeck);
  copySlideBackground(snapshot, slide);
  expect(getSlideBackgroundImageBytes(snapshot)).toEqual(bytes);
  expect(getSlideXmlString(snapshot)).toContain('<a:grayscl');
  expect(getSlideBackgroundImageCrop(snapshot)).toEqual({
    left: -0.1,
    right: 0.25,
    top: 0,
    bottom: 0,
  });
  setSlideBackground(slide, '#0000FF');
  copySlideBackground(slide, snapshot);
  const restored = await loadPresentation(await savePresentation(source));
  const result = getSlides(restored)[0]!;
  expect(getSlideBackgroundImageBytes(result)).toEqual(bytes);
  expect(getSlideBackgroundImageOpacity(result)).toBe(0.4);
  expect(getSlideBackgroundImageFillLayout(result)).toEqual(
    getSlideBackgroundImageFillLayout(snapshot),
  );
  expect(getSlideBackgroundImageCrop(result)).toEqual(getSlideBackgroundImageCrop(snapshot));
  expect(getSlideXmlString(result)).toContain('<a:grayscl');
});

it('shares media within a deck and does not duplicate relationships on repeated copies', async () => {
  const pres = await sourceDeck();
  const source = getSlides(pres)[0]!;
  const target = addBlankSlide(pres);
  const mediaCount = getMediaParts(pres).length;
  copySlideBackground(target, source);
  const xml = getSlideXmlString(target);
  copySlideBackground(target, source);
  expect(getSlideXmlString(target)).toBe(xml);
  expect(getMediaParts(pres)).toHaveLength(mediaCount);
  expect(getSlideBackgroundImageBytes(target)).toEqual(getSlideBackgroundImageBytes(source));
});

it('rejects missing relationships without modifying the target deck', async () => {
  const source = await sourceDeck(true);
  const target = createPresentation();
  const slide = addBlankSlide(target);
  setSlideBackground(slide, '#00FF00');
  const before = unzipSync(await savePresentation(target));
  expect(() => copySlideBackground(slide, getSlides(source)[0]!)).toThrow(/missing relationship/);
  // ZIP headers contain save-time timestamps; compare every package part instead.
  expect(unzipSync(await savePresentation(target))).toEqual(before);
});
