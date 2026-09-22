import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  getShapeBounds,
  getShapeFlip,
  getShapeRotation,
  getSlideShapes,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeFlip,
  setShapeRotation,
  ungroupShapes,
} from '../src/api/index.ts';
import { renderSlideSvg } from '../packages/preview/src/render-slide.ts';

describe('imported DrawingML boolean flip attributes', () => {
  for (const horizontal of ['1', 'true', '0', 'false', ' true ', ' false ']) {
    for (const vertical of ['1', 'true', '0', 'false', ' true ', ' false ']) {
      it(`renders and ungroups flipH=${horizontal}, flipV=${vertical} like numeric flags`, async () => {
        const pres = createPresentation();
        const slide = addBlankSlide(pres);
        const children = [1, 4].map((x) =>
          addSlideShape(slide, {
            preset: 'triangle',
            x: inches(x),
            y: inches(1),
            w: inches(2),
            h: inches(1),
          }),
        );
        setShapeRotation(children[0]!, 30);
        children.forEach((child) => setShapeFlip(child, { horizontal: true, vertical: true }));
        const group = groupShapes(children);
        setShapeRotation(group, 45);
        setShapeFlip(group, { horizontal: true, vertical: true });
        const parts = unzipSync(await savePresentation(pres));
        const path = 'ppt/slides/slide1.xml';
        const xml = strFromU8(parts[path]!);
        const h = horizontal.trim() === 'true' || horizontal.trim() === '1';
        const v = vertical.trim() === 'true' || vertical.trim() === '1';
        const load = (x: string, y: string) =>
          loadPresentation(
            zipSync({
              ...parts,
              [path]: strToU8(
                xml.replaceAll('flipH="1"', `flipH="${x}"`).replaceAll('flipV="1"', `flipV="${y}"`),
              ),
            }),
          );
        const numeric = await load(h ? '1' : '0', v ? '1' : '0');
        let imported = await load(horizontal, vertical);
        const referenceSlide = getSlides(numeric)[0]!;
        const expectedSvg = renderSlideSvg(numeric, referenceSlide);
        for (let pass = 0; pass < 2; pass++) {
          const importedSlide = getSlides(imported)[0]!;
          for (const shape of getSlideShapes(importedSlide)) {
            expect(getShapeFlip(shape)).toEqual({ horizontal: h, vertical: v });
          }
          expect(renderSlideSvg(imported, importedSlide)).toBe(expectedSvg);
          imported = await loadPresentation(await savePresentation(imported));
        }
        const geometry = (shape: ReturnType<typeof addSlideShape>) => ({
          bounds: getShapeBounds(shape),
          rotation: getShapeRotation(shape),
          flip: getShapeFlip(shape),
        });
        expect(ungroupShapes(getSlideShapes(getSlides(imported)[0]!)[0]!).map(geometry)).toEqual(
          ungroupShapes(getSlideShapes(referenceSlide)[0]!).map(geometry),
        );
      });
    }
  }
});
