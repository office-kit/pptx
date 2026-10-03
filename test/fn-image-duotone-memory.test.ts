import { readFile } from 'node:fs/promises';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { expect, it } from 'vitest';
import {
  getShapeImageDuotone,
  getShapeKind,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeImageRecolor,
} from '../src/api/index.ts';

it('preserves a theme duotone when a read value is applied again', async () => {
  const parts = unzipSync(
    await readFile(new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url)),
  );
  const slidePart = 'ppt/slides/slide1.xml';
  parts[slidePart] = strToU8(
    strFromU8(parts[slidePart]!).replace(
      '<a:blip r:embed="rId2"/>',
      '<a:blip r:embed="rId2"><a:duotone><a:schemeClr val="accent1"><a:tint val="45000"/></a:schemeClr><a:srgbClr val="D9C3A5"><a:satMod val="180000"/></a:srgbClr></a:duotone></a:blip>',
    ),
  );
  const pres = await loadPresentation(zipSync(parts));
  const picture = getSlideShapes(getSlides(pres)[0]!).find(
    (shape) => getShapeKind(shape) === 'picture',
  )!;
  const current = getShapeImageDuotone(pres, picture, { resolveColors: false })!;
  expect(current).toEqual({
    kind: 'duotone',
    colors: [
      { color: 'scheme:accent1', colorTransforms: [{ kind: 'tint', value: 0.45 }] },
      { color: '#D9C3A5', colorTransforms: [{ kind: 'satMod', value: 1.8 }] },
    ],
  });
  setShapeImageRecolor(picture, current);
  const xml = strFromU8(unzipSync(await savePresentation(pres))[slidePart]!);
  expect(xml).toContain('<a:schemeClr val="accent1"><a:tint val="45000"/></a:schemeClr>');
});

it('keeps alternate color transforms single-applied in memory', async () => {
  const parts = unzipSync(
    await readFile(new URL('./fixtures/minimal/one-image-slide.pptx', import.meta.url)),
  );
  const slidePart = 'ppt/slides/slide1.xml';
  parts[slidePart] = strToU8(
    strFromU8(parts[slidePart]!).replace(
      '<a:blip r:embed="rId2"/>',
      '<a:blip r:embed="rId2"><a:duotone><a:hslClr hue="0" sat="100%" lum="50%"><a:shade val="50%"/></a:hslClr><a:srgbClr val="D9C3A5"><a:satMod val="180000"/></a:srgbClr></a:duotone></a:blip>',
    ),
  );
  const pres = await loadPresentation(zipSync(parts));
  const picture = getSlideShapes(getSlides(pres)[0]!).find(
    (shape) => getShapeKind(shape) === 'picture',
  )!;
  const spec = getShapeImageDuotone(pres, picture, { resolveColors: false })!;
  const rendered = getShapeImageDuotone(pres, picture);
  expect(spec).toEqual({
    kind: 'duotone',
    colors: [
      { color: '#FF0000', colorTransforms: [{ kind: 'shade', value: 0.5 }] },
      { color: '#D9C3A5', colorTransforms: [{ kind: 'satMod', value: 1.8 }] },
    ],
  });
  setShapeImageRecolor(picture, spec);
  const xml = strFromU8(unzipSync(await savePresentation(pres))[slidePart]!);
  expect(xml).toContain('<a:srgbClr val="FF0000"><a:shade val="50000"/></a:srgbClr>');
  const reloaded = await loadPresentation(await savePresentation(pres));
  const restored = getSlideShapes(getSlides(reloaded)[0]!).find(
    (shape) => getShapeKind(shape) === 'picture',
  )!;
  expect(getShapeImageDuotone(reloaded, restored)).toEqual(rendered);
  expect(getShapeImageDuotone(reloaded, restored, { resolveColors: false })).toEqual(spec);
});
