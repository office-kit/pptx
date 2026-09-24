import { describe, it, expect } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import {
  createPresentation,
  addBlankSlide,
  getSlides,
  getSlideLayout,
  setSlideLayoutBackground,
  getSlideBackground,
  getSlideLayoutBackground,
  getSlideMasterBackground,
  getSlideMasterBackgroundImageBytes,
  getSlideMasterBackgroundPatternFill,
  setSlideBackground,
  setSlideBackgroundImage,
  setSlideBackgroundPatternFill,
  applySlideBackgroundToAll,
  savePresentation,
  loadPresentation,
} from '../src/api/index.ts';

const deck = () => {
  const pres = createPresentation();
  const source = addBlankSlide(pres);
  const second = addBlankSlide(pres);
  setSlideBackground(source, '#FF0000');
  setSlideBackground(second, '#0000FF');
  setSlideLayoutBackground(getSlideLayout(source)!, '#00FF00');
  return { pres, source };
};

describe('apply background to all', () => {
  it('matches PowerPoint by clearing slide/layout overrides and storing the source on the master', async () => {
    const { pres, source } = deck();
    applySlideBackgroundToAll(pres, source);
    const loaded = await loadPresentation(await savePresentation(pres));
    for (const presentation of [pres, loaded]) {
      for (const slide of getSlides(presentation)) {
        expect(getSlideBackground(slide)).toEqual({ kind: 'inherit' });
        const layout = getSlideLayout(slide)!;
        expect(getSlideLayoutBackground(layout)).toEqual({ kind: 'inherit' });
        expect(getSlideMasterBackground(presentation, layout)).toEqual({
          kind: 'solid',
          color: '#FF0000',
        });
      }
    }
  });
  it('copies an inherited pattern without resolving theme tokens', async () => {
    const { pres, source } = deck();
    setSlideBackgroundPatternFill(source, {
      preset: 'wave',
      foreground: 'accent2',
      background: 'bg2',
    });
    applySlideBackgroundToAll(pres, source);
    applySlideBackgroundToAll(pres, getSlides(pres)[1]!);
    const loaded = await loadPresentation(await savePresentation(pres));
    expect(
      getSlideMasterBackgroundPatternFill(loaded, getSlideLayout(getSlides(loaded)[0]!)!, {
        preserveTheme: true,
      }),
    ).toEqual({ preset: 'wave', foreground: 'accent2', background: 'bg2' });
  });
  it('rewires background images to the master and reuses relationships on repeated application', async () => {
    const { pres, source } = deck();
    const image = Uint8Array.from(
      Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jv1sAAAAASUVORK5CYII=',
        'base64',
      ),
    );
    setSlideBackgroundImage(source, image);
    applySlideBackgroundToAll(pres, source);
    const first = unzipSync(await savePresentation(pres));
    applySlideBackgroundToAll(pres, source);
    const bytes = await savePresentation(pres);
    const second = unzipSync(bytes);
    for (const name of Object.keys(first).filter((name) =>
      name.startsWith('ppt/slideMasters/_rels/'),
    ))
      expect(strFromU8(second[name]!)).toBe(strFromU8(first[name]!));
    const loaded = await loadPresentation(bytes);
    for (const slide of getSlides(loaded))
      expect(getSlideMasterBackgroundImageBytes(loaded, getSlideLayout(slide)!)).toEqual(image);
  });
  it('rejects a foreign source without changing the presentation', async () => {
    const { pres } = deck();
    const other = deck();
    const before = await savePresentation(pres);
    expect(() => applySlideBackgroundToAll(pres, other.source)).toThrow('source must belong');
    expect(await savePresentation(pres)).toEqual(before);
  });
  it('leaves backgrounds unchanged if an imported image relationship is missing', async () => {
    const { pres } = deck();
    const parts = unzipSync(await savePresentation(pres));
    parts['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(parts['ppt/slides/slide1.xml']!).replace(
        /<p:bg>[\s\S]*?<\/p:bg>/,
        '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="missing"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>',
      ),
    );
    const loaded = await loadPresentation(zipSync(parts));
    const before = await savePresentation(loaded);
    expect(() => applySlideBackgroundToAll(loaded, getSlides(loaded)[0]!)).toThrow(
      'missing relationship',
    );
    expect(await savePresentation(loaded)).toEqual(before);
  });
});
