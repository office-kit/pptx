// A shape wrapped in shape-tree `mc:AlternateContent` (how the desktop app
// saves a shape holding an equation: an a14 choice, a picture fallback) is
// read from the branch a reader takes, and structural edits move the wrapper.

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  bringShapeToFront,
  copyShape,
  createPresentation,
  getGroupChildren,
  getShapeParagraphElements,
  getShapePosition,
  getShapeZIndex,
  getSlideShapes,
  getSlides,
  groupShapes,
  inches,
  loadPresentation,
  removeShape,
  savePresentation,
  setShapePosition,
  ungroupShapes,
  type PresentationData,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { mathRun, mathZone } from './lib/equation-deck.ts';

const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MARKER = 'WRAPPED';
const MARKER_SHAPE = /<p:sp>(?:(?!<p:sp>).)*?WRAPPED(?:(?!<\/p:sp>).)*<\/p:sp>/;
const MARKER_RUN = /<a:r>(?:(?!<\/a:r>).)*?WRAPPED<\/a:t><\/a:r>/;

interface Wrap {
  /** Attributes on `mc:Choice`; the default requires a14, declared there. */
  readonly choice?: string;
  /** Prefix declarations on `mc:AlternateContent`. */
  readonly outer?: string;
}

const slideXml = async (pres: PresentationData): Promise<string> =>
  strFromU8(unzipSync(await savePresentation(pres))['ppt/slides/slide1.xml']!);

// Slide 1 holds "Before", the wrapped shape, and "After", in that order. The
// choice's copy of the shape holds an equation; the fallback's holds the text
// "Fallback" (the desktop app puts a picture there).
const deckWithWrappedShape = async ({
  choice = `xmlns:a14="${A14}" Requires="a14"`,
  outer = '',
}: Wrap = {}): Promise<PresentationData> => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const box = (y: number, text: string) =>
    addSlideTextBox(slide, { x: inches(1), y: inches(y), w: inches(4), h: inches(1), text });
  box(0.5, 'Before');
  box(2, MARKER);
  box(3.5, 'After');
  const zip = unzipSync(await savePresentation(pres));
  const part = 'ppt/slides/slide1.xml';
  const source = strFromU8(zip[part]!);
  const shape = MARKER_SHAPE.exec(source)?.[0];
  if (shape === undefined) throw new Error('marker shape not found');
  const withEquation = shape.replace(MARKER_RUN, mathZone(mathRun('x=1')));
  const withFallback = shape.replace(MARKER, 'Fallback');
  const wrapped =
    `<mc:AlternateContent xmlns:mc="${MC}" ${outer}>` +
    `<mc:Choice ${choice}>${withEquation}</mc:Choice>` +
    `<mc:Fallback>${withFallback}</mc:Fallback></mc:AlternateContent>`;
  zip[part] = strToU8(source.replace(shape, wrapped));
  return loadPresentation(zipSync(zip));
};

const firstSlide = (pres: PresentationData) => getSlides(pres)[0]!;
const hasEquation = (pres: PresentationData, index: number): boolean =>
  getShapeParagraphElements(getSlideShapes(firstSlide(pres))[index]!)[0]!.some(
    (element) => element.kind === 'math',
  );
const texts = (pres: PresentationData) =>
  getSlideShapes(firstSlide(pres)).map((shape) =>
    getShapeParagraphElements(shape)
      .flat()
      .map((element) => element.text)
      .join(''),
  );

describe('shapes wrapped in mc:AlternateContent', () => {
  it('reads the a14 choice, in document order', async () => {
    const pres = await deckWithWrappedShape();
    expect(texts(pres)).toEqual(['Before', 'x=1', 'After']);
    expect(hasEquation(pres, 1)).toBe(true);
  });

  it('resolves the required prefix against declarations around the wrapper', async () => {
    const pres = await deckWithWrappedShape({
      choice: 'Requires="a14"',
      outer: `xmlns:a14="${A14}"`,
    });
    expect(texts(pres)).toEqual(['Before', 'x=1', 'After']);
  });

  it('reads the fallback when the choice requires an unknown namespace', async () => {
    const pres = await deckWithWrappedShape({ choice: 'xmlns:z="urn:unknown" Requires="z"' });
    expect(texts(pres)).toEqual(['Before', 'Fallback', 'After']);
  });

  it('reads the fallback when the choice also requires an unknown namespace', async () => {
    const pres = await deckWithWrappedShape({
      choice: `xmlns:a14="${A14}" xmlns:z="urn:unknown" Requires="a14 z"`,
    });
    expect(texts(pres)).toEqual(['Before', 'Fallback', 'After']);
  });

  it('draws the equation in the preview', async () => {
    const pres = await deckWithWrappedShape();
    expect(renderSlideToSvg(pres, firstSlide(pres))).toContain('data-pptx-math');
  });

  it('keeps the wrapper and both branches on save', async () => {
    const pres = await deckWithWrappedShape();
    const xml = await slideXml(pres);
    expect(xml).toContain('<mc:Choice');
    expect(xml).toContain('Fallback</a:t>');
  });

  it('removes the wrapper with the shape', async () => {
    const pres = await deckWithWrappedShape();
    removeShape(getSlideShapes(firstSlide(pres))[1]!);
    expect(texts(pres)).toEqual(['Before', 'After']);
    expect(await slideXml(pres)).not.toContain('AlternateContent');
  });

  it('reorders the wrapper as one shape', async () => {
    const pres = await deckWithWrappedShape();
    const wrapped = getSlideShapes(firstSlide(pres))[1]!;
    expect(getShapeZIndex(wrapped)).toBe(1);
    bringShapeToFront(wrapped);
    expect(texts(pres)).toEqual(['Before', 'After', 'x=1']);
    expect(getShapeZIndex(getSlideShapes(firstSlide(pres))[2]!)).toBe(2);
    const xml = await slideXml(pres);
    expect(xml.indexOf('After')).toBeLessThan(xml.indexOf('<mc:AlternateContent'));
  });

  it('copies the wrapper, giving both branches the same new id', async () => {
    const pres = await deckWithWrappedShape();
    const slide = firstSlide(pres);
    copyShape(slide, getSlideShapes(slide)[1]!);
    expect(texts(pres)).toEqual(['Before', 'x=1', 'After', 'x=1']);
    const xml = await slideXml(pres);
    const copy = xml.slice(xml.lastIndexOf('<mc:AlternateContent'));
    const ids = [...copy.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1]);
    expect(ids).toHaveLength(2);
    expect(ids[0]).toBe(ids[1]);
    const original = xml.slice(xml.indexOf('<mc:AlternateContent'));
    expect(ids[0]).not.toBe(/<p:cNvPr id="(\d+)"/.exec(original)?.[1]);
  });

  it('groups and ungroups the wrapper, moving every branch', async () => {
    const pres = await deckWithWrappedShape();
    const slide = firstSlide(pres);
    const [before, wrapped] = getSlideShapes(slide);
    const group = groupShapes([before!, wrapped!]);
    expect(getGroupChildren(group)).toHaveLength(2);
    expect(await slideXml(pres)).toMatch(/<p:grpSp>.*<mc:AlternateContent.*<\/p:grpSp>/);

    setShapePosition(group, inches(2), inches(1));
    const [first, second] = ungroupShapes(group);
    expect(getShapeParagraphElements(first!)[0]![0]!.text).toBe('Before');
    expect(hasEquation(pres, getSlideShapes(slide).indexOf(second!))).toBe(true);
    const moved = getShapePosition(second!);
    expect(moved.x).toBe(inches(2));

    const xml = await slideXml(pres);
    expect(xml).not.toContain('<p:grpSp>');
    const offsets = [
      ...xml.slice(xml.indexOf('<mc:AlternateContent')).matchAll(/<a:off x="(\d+)" y="(\d+)"/g),
    ]
      .slice(0, 2)
      .map((m) => [Number(m[1]), Number(m[2])]);
    expect(offsets[0]).toEqual([moved.x, moved.y]);
    expect(offsets[1]).toEqual(offsets[0]);
  });
});
