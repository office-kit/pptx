import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideImage,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  getShapeId,
  getShapeImageBytes,
  getShapeKind,
  getShapeText,
  getSlides,
  getSlideShapes,
  getSlideText,
  groupShapes,
  inches,
  listPackageParts,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import {
  buildClipboardDeck,
  clipboardMarker,
  clipboardPlainText,
  clipboardShapeBounds,
  clipboardShapes,
  parseClipboardMarker,
  pasteSource,
  selectedSlidesFileName,
} from '../packages/editor/src/core/clipboard-deck.ts';
import { buildPng } from './lib/build-png.ts';
import { expectPackageValid } from './lib/expect-package-valid.ts';

const box = (x: number, y: number, w = 1, h = 1) => ({
  x: inches(x),
  y: inches(y),
  w: inches(w),
  h: inches(h),
});

async function deck() {
  const pres = createPresentation();
  const first = addBlankSlide(pres);
  addSlideTextBox(first, { ...box(1, 1, 4), text: 'Title' });
  addSlideImage(first, buildPng(2, 2, [255, 0, 0]), box(1, 2));
  const second = addBlankSlide(pres);
  addSlideTextBox(second, { ...box(1, 1, 4), text: 'Second' });
  addSlideTable(second, {
    ...box(1, 3, 4, 2),
    rows: [
      ['A', 'B'],
      ['C', 'D'],
    ],
  });
  const third = addBlankSlide(pres);
  addSlideImage(third, buildPng(3, 3, [0, 0, 255]), box(2, 2));
  addSlideTextBox(third, { ...box(1, 1, 4), text: 'Third' });
  return pres;
}

const mediaParts = (pres: Parameters<typeof listPackageParts>[0]) =>
  listPackageParts(pres).filter((part) => part.name.startsWith('/ppt/media/'));

describe('clipboard payload', () => {
  it('keeps only the copied slides, with their pictures, as a valid package', async () => {
    const source = await savePresentation(await deck());
    const copied = await buildClipboardDeck(source, { kind: 'slides', slideIndices: [1, 2] });
    expect(getSlides(copied.pres).map(getSlideText)).toEqual(['Second', 'Third']);
    // The first slide's picture is gone; the third's stays.
    expect(mediaParts(copied.pres)).toHaveLength(1);
    const reloaded = await expectPackageValid(await loadPresentation(copied.bytes));
    expect(getSlides(reloaded)).toHaveLength(2);
    expect(clipboardPlainText(copied)).toBe('Second\nA\tB\nC\tD\n\nThird');
  });

  it('leaves exactly the copied objects at their positions, in stacking order', async () => {
    const pres = await deck();
    const slide = getSlides(pres)[0]!;
    const [title, picture] = getSlideShapes(slide);
    const copied = await buildClipboardDeck(await savePresentation(pres), {
      kind: 'shapes',
      slideIndex: 0,
      shapeIds: [getShapeId(picture!), getShapeId(title!)],
    });
    expect(getSlides(copied.pres)).toHaveLength(1);
    const shapes = clipboardShapes(copied.pres);
    expect(shapes.map(getShapeKind)).toEqual(['shape', 'picture']);
    expect(getShapeText(shapes[0]!)).toBe('Title');
    expect(getShapeImageBytes(shapes[1]!)).not.toBeNull();
    expect(clipboardShapeBounds(copied)).toEqual({
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
    });
    expect(clipboardPlainText(copied)).toBe('Title');
    await expectPackageValid(await loadPresentation(copied.bytes));
  });

  it('lifts an object out of its group without the rest of the group', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const a = addSlideShape(slide, { preset: 'rect', ...box(1, 1), text: 'A' });
    const b = addSlideShape(slide, { preset: 'rect', ...box(4, 1), text: 'B' });
    groupShapes([a, b]);
    const copied = await buildClipboardDeck(await savePresentation(pres), {
      kind: 'shapes',
      slideIndex: 0,
      shapeIds: [getShapeId(b)],
    });
    expect(clipboardPlainText(copied)).toBe('B');
    expect(clipboardShapeBounds(copied)).toEqual(box(4, 1));
  });

  it('marks clipboard HTML with the copy it belongs to', () => {
    const html = `<div data-pptx-editor-clipboard="${clipboardMarker('shapes', 'ab12')}">x</div>`;
    expect(parseClipboardMarker(html)).toEqual({ kind: 'shapes', id: 'ab12' });
    expect(parseClipboardMarker('<p>other app</p>')).toBeNull();
  });

  it('pastes the newest copy', () => {
    const own = { id: 'a', written: 'done' as const };
    // Another app copied after this editor wrote the clipboard.
    expect(pasteSource(own, { id: null, empty: false })).toBe('system');
    // Another editor's copy.
    expect(pasteSource(own, { id: 'b', empty: false })).toBe('system');
    // This editor's own copy: paste it without re-reading the payload.
    expect(pasteSource(own, { id: 'a', empty: false })).toBe('own');
    // The write is still running or failed: the clipboard holds something older.
    expect(pasteSource({ ...own, written: 'pending' }, { id: null, empty: false })).toBe('own');
    expect(pasteSource({ ...own, written: 'failed' }, { id: 'b', empty: false })).toBe('own');
    // Clipboard unreadable or empty.
    expect(pasteSource(own, null)).toBe('own');
    expect(pasteSource(null, { id: null, empty: true })).toBe('none');
    expect(pasteSource(null, { id: null, empty: false })).toBe('system');
  });

  it('names a download after the slides it holds', () => {
    expect(selectedSlidesFileName('Deck.pptx', [2])).toBe('Deck-slide-3.pptx');
    expect(selectedSlidesFileName('Deck.pptx', [0, 3])).toBe('Deck-slides-1-4.pptx');
  });
});

describe('renderSlideToSvg background option', () => {
  it('leaves out the background and master graphics when false', async () => {
    const pres = await deck();
    const slide = getSlides(pres)[0]!;
    const full = renderSlideToSvg(pres, slide, { textLayout: 'svg' });
    const bare = renderSlideToSvg(pres, slide, { textLayout: 'svg', background: false });
    const backgroundRect = /^<svg[^>]*>(?:<defs>.*?<\/defs>)?<rect width="1280.00" height="720.00"/;
    expect(full).toMatch(backgroundRect);
    expect(bare).not.toMatch(backgroundRect);
    expect(bare).toContain('Title');
  });
});
