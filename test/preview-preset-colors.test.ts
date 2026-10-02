import { expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  getSlideSize,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFill,
} from '@office-kit/pptx';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';

it('renders imported preset fill colors and their transforms after save/reload', async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  for (const x of [1, 4]) {
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(x),
      y: inches(1),
      w: inches(2),
      h: inches(2),
    });
    setShapeFill(shape, x === 1 ? '#112233' : '#445566');
  }
  const parts = unzipSync(await savePresentation(pres));
  const name = 'ppt/slides/slide1.xml';
  const xml = strFromU8(parts[name]!);
  expect(xml).toContain('<a:srgbClr val="112233"/>');
  expect(xml).toContain('<a:srgbClr val="445566"/>');
  parts[name] = strToU8(
    xml
      .replace('<a:srgbClr val="112233"/>', '<a:prstClr val="red"/>')
      .replace(
        '<a:srgbClr val="445566"/>',
        '<a:prstClr val="red"><a:shade val="50000"/></a:prstClr>',
      ),
  );
  const imported = await loadPresentation(zipSync(parts));
  const restored = await loadPresentation(await savePresentation(imported));
  const size = getSlideSize(restored)!;
  const { image } = renderSlideToRgba(restored, getSlides(restored)[0]!, { width: 960 });
  for (const [center, expected] of [
    [2, [255, 0, 0]],
    [5, [188, 0, 0]],
  ] as const) {
    const x = Math.round((inches(center) / size.width) * image.width);
    const y = Math.round((inches(2) / size.height) * image.height);
    const offset = (y * image.width + x) * 4;
    expect(Array.from(image.data.slice(offset, offset + 3))).toEqual(expected);
  }
});
