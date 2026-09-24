import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import {
  asColor,
  getSlides,
  loadPresentation,
  savePresentation,
  getSlideBackgroundGradientFill,
  setSlideBackgroundGradientFill,
} from '../src/api/index.ts';

async function importedGradient() {
  const zip = unzipSync(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  const path = 'ppt/slides/slide1.xml';
  zip[path] = strToU8(
    strFromU8(zip[path]!)
      .replace(/<p:bg\b[^>]*>[\s\S]*?<\/p:bg>/g, '')
      .replace(
        /(<p:cSld\b[^>]*>)/,
        '$1<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="accent1"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill></p:bgPr></p:bg>',
      ),
  );
  const pres = await loadPresentation(zipSync(zip));
  const slide = getSlides(pres)[0]!;
  return { pres, slide, path };
}

it('keeps imported theme color transforms when moving a gradient stop and changing direction', async () => {
  const { pres, slide, path } = await importedGradient();
  const before = getSlideBackgroundGradientFill(slide)!;
  expect(before.stops[0]).toMatchObject({
    colorTransforms: [
      { kind: 'tint', value: 0.8 },
      { kind: 'satMod', value: 3 },
    ],
  });
  setSlideBackgroundGradientFill(slide, {
    ...before,
    angleDeg: 90,
    stops: before.stops.map((stop, index) => ({
      ...stop,
      color: asColor(stop.color)!,
      offset: index === 0 ? 0.1 : stop.offset,
    })),
  });
  const after = getSlideBackgroundGradientFill(slide)!;
  expect(after.angleDeg).toBe(90);
  expect(after.stops[0]!.offset).toBe(0.1);
  expect(after.stops.map((stop) => stop.resolvedColor)).toEqual(
    before.stops.map((stop) => stop.resolvedColor),
  );
  const saved = unzipSync(await savePresentation(pres));
  expect(strFromU8(saved[path]!)).toContain('<a:tint val="80000"/><a:satMod val="300000"/>');
  expect(strFromU8(saved[path]!)).toContain('<a:shade val="30000"/><a:satMod val="200000"/>');
});

it('edits brightness and opacity without discarding other imported adjustments', async () => {
  const { pres, slide, path } = await importedGradient();
  const before = getSlideBackgroundGradientFill(slide)!;
  setSlideBackgroundGradientFill(slide, {
    ...before,
    stops: before.stops.map((stop) => ({
      ...stop,
      color: asColor(stop.color)!,
      brightness: 0.25,
      opacity: 0.6,
    })),
  });
  const after = getSlideBackgroundGradientFill(slide)!;
  expect(after.stops[0]).toMatchObject({ brightness: 0.25, opacity: 0.6 });
  expect(after.stops[0]!.colorTransforms).toEqual([
    { kind: 'tint', value: 0.8 },
    { kind: 'satMod', value: 3 },
    { kind: 'alpha', value: 0.6 },
    { kind: 'lumMod', value: 0.75 },
    { kind: 'lumOff', value: 0.25 },
  ]);
  const bytes = await savePresentation(pres);
  const reloaded = await loadPresentation(bytes);
  expect(getSlideBackgroundGradientFill(getSlides(reloaded)[0]!)).toEqual(after);
  setSlideBackgroundGradientFill(slide, {
    ...after,
    stops: after.stops.map((stop) => ({ ...stop, color: asColor(stop.color)! })),
  });
  expect(strFromU8(unzipSync(await savePresentation(pres))[path]!)).toEqual(
    strFromU8(unzipSync(bytes)[path]!),
  );
});

it('rejects invalid color adjustments before changing the background', async () => {
  const { slide } = await importedGradient();
  const before = getSlideBackgroundGradientFill(slide)!;
  for (const colorTransforms of [
    [{ kind: 'tint' as const, value: 1.1 }],
    [{ kind: 'hue' as const, value: 359.999999 }],
    [{ kind: 'satMod' as const, value: Number.NaN }],
  ]) {
    expect(() =>
      setSlideBackgroundGradientFill(slide, {
        ...before,
        stops: before.stops.map((stop) => ({
          ...stop,
          color: asColor(stop.color)!,
          colorTransforms,
        })),
      }),
    ).toThrow(RangeError);
    expect(getSlideBackgroundGradientFill(slide)).toEqual(before);
  }
});
