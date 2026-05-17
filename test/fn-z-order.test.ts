// Free-function z-order helpers: bring forward / send backward,
// bring to front / send to back.
//
// Verifies that the order of `<p:sp>` / `<p:pic>` / `<p:cxnSp>` /
// `<p:graphicFrame>` children of `<p:spTree>` is rearranged correctly
// while the `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface stays at the top.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  bringShapeForward,
  bringShapeToFront,
  getShapeName,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  sendShapeBackward,
  sendShapeToBack,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const orderedNames = (slide: ReturnType<typeof getSlides>[number]): string[] =>
  getSlideShapes(slide).map((s) => getShapeName(s));

describe('fn API: z-order', () => {
  it('bringShapeToFront moves a shape to the end of spTree', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'A',
    });
    const b = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'B',
    });
    const c = addSlideShape(slide, {
      preset: 'triangle',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'C',
    });
    expect(orderedNames(slide).slice(-3)).toEqual(['A', 'B', 'C']);

    bringShapeToFront(a);
    expect(orderedNames(slide).slice(-3)).toEqual(['B', 'C', 'A']);
    // No-op when already at front.
    bringShapeToFront(a);
    expect(orderedNames(slide).slice(-3)).toEqual(['B', 'C', 'A']);

    // Reference unused locals.
    void b;
    void c;
  });

  it('sendShapeToBack moves a shape behind every other shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'A',
    });
    const b = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'B',
    });
    const c = addSlideShape(slide, {
      preset: 'triangle',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'C',
    });

    sendShapeToBack(c);
    // C is now the first shape child of spTree.
    expect(orderedNames(slide)[0]).toBe('C');

    // Reference unused locals.
    void a;
    void b;
  });

  it('bringShapeForward / sendShapeBackward swap with the adjacent shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const a = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'A',
    });
    const b = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'B',
    });
    const c = addSlideShape(slide, {
      preset: 'triangle',
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      name: 'C',
    });

    // Initial: ..., A, B, C
    bringShapeForward(a);
    // After: ..., B, A, C
    expect(orderedNames(slide).slice(-3)).toEqual(['B', 'A', 'C']);

    sendShapeBackward(c);
    // After: ..., B, C, A
    expect(orderedNames(slide).slice(-3)).toEqual(['B', 'C', 'A']);

    // No-op when at the boundary in each direction.
    bringShapeForward(a);
    expect(orderedNames(slide).slice(-3)).toEqual(['B', 'C', 'A']);

    void b;
  });
});
