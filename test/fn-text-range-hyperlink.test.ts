import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  setShapeClickAction,
  getShapeClickAction,
  getShapeRunClickAction,
  getSlideIndex,
  getSlides,
  getSlideShapes,
  getShapeParagraphElements,
  getShapeRunHyperlink,
  getShapeRunHyperlinkTooltip,
  getShapeText,
  getSlideXmlString,
  inches,
  loadPresentation,
  savePresentation,
  setShapeHyperlink,
  setShapeTextFormat,
} from '../src/api/index.ts';

async function fixture() {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/two-slides.pptx', import.meta.url)),
  );
  const slide = getSlides(pres)[0]!;
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text: 'A日本😀Z\nSecond',
  });
  return { pres, slide, shape };
}

describe('text range hyperlinks', () => {
  it('preserves surrounding links and formatting across paragraphs, removal and reload', async () => {
    const { pres, shape } = await fixture();
    setShapeHyperlink(shape, 'https://original.example');
    setShapeTextFormat(shape, { bold: true }, { range: { start: 1, end: 5 } });
    setShapeHyperlink(shape, 'https://selected.example', '選択部分', {
      range: { start: 2, end: 10 },
    });
    expect(getShapeText(shape)).toBe('A日本😀Z\nSecond');
    expect(
      getShapeParagraphElements(shape, 0).map((e) => (e.kind === 'br' ? '\n' : e.text)),
    ).toEqual(['A', '日', '本😀', 'Z']);
    expect([0, 1, 2, 3].map((r) => getShapeRunHyperlink(shape, 0, r))).toEqual([
      'https://original.example',
      'https://original.example',
      'https://selected.example',
      'https://selected.example',
    ]);
    expect(getShapeParagraphElements(shape, 0)[2]!.format?.bold).toBe(true);
    expect(getShapeRunHyperlink(shape, 1, 0)).toBe('https://selected.example');
    expect(getShapeRunHyperlink(shape, 1, 1)).toBe('https://original.example');
    setShapeHyperlink(shape, null, undefined, { range: { start: 3, end: 5 } });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const result = getSlideShapes(getSlides(reloaded)[0]!).at(-1)!;
    expect(getShapeText(result)).toBe('A日本😀Z\nSecond');
    expect(getShapeRunHyperlink(result, 0, 2)).toBe('https://selected.example');
    expect(getShapeRunHyperlinkTooltip(result, 0, 2)).toBe('選択部分');
    expect(getShapeRunHyperlink(result, 0, 3)).toBeNull();
    expect(getShapeParagraphElements(result, 0)[3]!.format?.bold).toBe(true);
  });
  it('round-trips selected slide targets and every navigation preset independently of shape actions', async () => {
    const { pres, shape } = await fixture();
    setShapeHyperlink(shape, 'https://surrounding.example');
    setShapeClickAction(shape, { kind: 'firstSlide' });
    setShapeTextFormat(shape, { bold: true });
    const range = { start: 1, end: 5 };
    setShapeClickAction(shape, { kind: 'slide', slide: getSlides(pres)[1]! }, { range });
    const action = getShapeRunClickAction(shape, 0, 1);
    expect(action?.kind).toBe('slide');
    if (action?.kind === 'slide') expect(getSlideIndex(pres, action.slide)).toBe(1);
    const internalReloaded = await loadPresentation(await savePresentation(pres));
    const internalShape = getSlideShapes(getSlides(internalReloaded)[0]!).at(-1)!;
    const internalAction = getShapeRunClickAction(internalShape, 0, 1);
    expect(internalAction?.kind).toBe('slide');
    if (internalAction?.kind === 'slide')
      expect(getSlideIndex(internalReloaded, internalAction.slide)).toBe(1);
    for (const kind of ['nextSlide', 'prevSlide', 'firstSlide', 'lastSlide'] as const) {
      setShapeClickAction(shape, { kind }, { range });
      const loaded = await loadPresentation(await savePresentation(pres));
      const result = getSlideShapes(getSlides(loaded)[0]!).at(-1)!;
      expect(getShapeRunClickAction(result, 0, 1)).toEqual({ kind });
      expect(getShapeClickAction(result)).toEqual({ kind: 'firstSlide' });
      expect(getShapeRunHyperlink(result, 0, 0)).toBe('https://surrounding.example');
      expect(getShapeRunHyperlink(result, 0, 2)).toBe('https://surrounding.example');
      expect(getShapeParagraphElements(result, 0)[1]!.format?.bold).toBe(true);
    }
    setShapeClickAction(shape, null, { range });
    expect(getShapeRunClickAction(shape, 0, 1)).toBeNull();
    expect(getShapeClickAction(shape)).toEqual({ kind: 'firstSlide' });
  });

  it('rejects foreign targets and invalid ranges without losing an existing action', async () => {
    const { pres, shape } = await fixture();
    const foreign = await fixture();
    setShapeClickAction(shape, { kind: 'lastSlide' });
    const before = await savePresentation(pres);
    expect(() => setShapeClickAction(shape, { kind: 'slide', slide: foreign.slide })).toThrow();
    expect(() =>
      setShapeClickAction(shape, { kind: 'nextSlide' }, { range: { start: 4, end: 5 } }),
    ).toThrow();
    expect(await savePresentation(pres)).toEqual(before);
  });

  it('rejects invalid or split-surrogate ranges and ignores empty ranges without mutation', async () => {
    const { pres, slide, shape } = await fixture();
    const before = getSlideXmlString(slide);
    const bytes = await savePresentation(pres);
    for (const range of [
      { start: -1, end: 2 },
      { start: 4, end: 5 },
      { start: 0, end: 100 },
      { start: 2, end: 1 },
    ]) {
      expect(() =>
        setShapeHyperlink(shape, 'https://unused.example', undefined, { range }),
      ).toThrow();
      expect(getSlideXmlString(slide)).toBe(before);
    }
    setShapeHyperlink(shape, 'https://unused.example', undefined, { range: { start: 1, end: 1 } });
    expect(await savePresentation(pres)).toEqual(bytes);
  });
});
