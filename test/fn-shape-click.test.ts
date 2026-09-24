// Free-function shape click-action API.
//
// Verifies `<a:hlinkClick>` lands on the shape's cNvPr for each of the
// supported action kinds (URL, slide jump, preset show navigation).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideShape,
  createPresentation,
  getGroupChildren,
  getShapeClickAction,
  getShapeHyperlinkTooltip,
  getShapeKind,
  groupShapes,
  inches,
  removeSlide,
  readPackagePart,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeClickAction,
} from '../src/api/index.ts';

import { findCNvPr } from '../src/api/fn/shape-click-action.ts';
import { NS, elem, qname } from '../src/internal/xml/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

const slideRels = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  const bytesPart = readPackagePart(pres, `/ppt/slides/_rels/slide${slideIndex + 1}.xml.rels`);
  return bytesPart ? new TextDecoder().decode(bytesPart) : '';
};

describe('fn API: setShapeClickAction', () => {
  it('attaches a URL click action and removes it on null', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;

    setShapeClickAction(shape, { kind: 'url', url: 'https://example.com/' });
    const wired = await savePresentation(pres);
    expect(await slideRels(wired, 0)).toContain('https://example.com/');
    expect(await slideXml(wired, 0)).toContain('hlinkClick');

    setShapeClickAction(shape, null);
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('hlinkClick');
  });

  it('jumps to a slide via slide-rel + ppaction://hlinksldjump', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    const target = slides[1]!;
    const shape = getSlideShapes(slides[0]!)[0]!;

    setShapeClickAction(shape, { kind: 'slide', slide: target });
    const bytes = await savePresentation(pres);
    expect(await slideXml(bytes, 0)).toContain('hlinksldjump');
    expect(await slideRels(bytes, 0)).toContain('slide2.xml');
  });

  it('preset navigation kinds emit the ppaction without allocating a rel', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeClickAction(shape, { kind: 'nextSlide' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('hlinkshowjump?jump=nextslide');
  });
});

describe('object link persistence', () => {
  it('round-trips a group link and ScreenTip without changing child links', async () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const bounds = { x: inches(1), y: inches(1), w: inches(2), h: inches(1) };
    const first = addSlideShape(slide, { ...bounds, preset: 'rect' });
    const second = addSlideShape(slide, { ...bounds, x: inches(4), preset: 'ellipse' });
    setShapeClickAction(first, { kind: 'nextSlide' });
    const group = groupShapes([first, second]);
    const tooltip = '資料 <詳細> & "リンク"';
    setShapeClickAction(group, { kind: 'url', url: 'https://example.com/?a=1&b=2' }, tooltip);
    const loaded = await loadPresentation(await savePresentation(pres));
    const restored = getSlideShapes(getSlides(loaded)[0]!).find(
      (s) => getShapeKind(s) === 'group',
    )!;
    expect(getShapeClickAction(restored)).toEqual({
      kind: 'url',
      url: 'https://example.com/?a=1&b=2',
    });
    expect(getShapeHyperlinkTooltip(restored)).toBe(tooltip);
    expect(getGroupChildren(restored).map(getShapeClickAction)).toEqual([
      { kind: 'nextSlide' },
      null,
    ]);
    setShapeClickAction(restored, { kind: 'lastSlide' }, '末尾');
    expect(getShapeClickAction(restored)).toEqual({ kind: 'lastSlide' });
    setShapeClickAction(restored, null);
    const cleared = await loadPresentation(await savePresentation(loaded));
    const clearedGroup = getSlideShapes(getSlides(cleared)[0]!).find(
      (s) => getShapeKind(s) === 'group',
    )!;
    expect(getShapeClickAction(clearedGroup)).toBeNull();
    expect(getShapeHyperlinkTooltip(clearedGroup)).toBeNull();
    expect(getGroupChildren(clearedGroup).map(getShapeClickAction)).toEqual([
      { kind: 'nextSlide' },
      null,
    ]);
  });

  it('rejects foreign and deleted destinations without mutating an existing link', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slide, target] = getSlides(pres);
    const shape = getSlideShapes(slide!)[0]!;
    setShapeClickAction(shape, { kind: 'url', url: 'https://example.com/' }, 'Existing');
    const foreign = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    removeSlide(pres, target!);
    const before = getSlideXmlString(slide!);
    const beforeRels = await slideRels(await savePresentation(pres), 0);
    for (const destination of [getSlides(foreign)[1]!, target!]) {
      expect(() => setShapeClickAction(shape, { kind: 'slide', slide: destination })).toThrow(
        'target slide must belong',
      );
      expect(getSlideXmlString(slide!)).toBe(before);
      expect(await slideRels(await savePresentation(pres), 0)).toBe(beforeRels);
      expect(getShapeHyperlinkTooltip(shape)).toBe('Existing');
    }
  });

  it('places the click link before hover and extensions and preserves them on removal', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    const metadata = findCNvPr(shape)!;
    metadata.children.push(
      elem(qname('a', 'hlinkHover', NS.dml)),
      elem(qname('a', 'extLst', NS.dml)),
    );
    setShapeClickAction(shape, { kind: 'firstSlide' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml.indexOf('hlinkClick')).toBeLessThan(xml.indexOf('hlinkHover'));
    expect(xml.indexOf('hlinkHover')).toBeLessThan(xml.indexOf('a:extLst'));
    setShapeClickAction(shape, null);
    const cleared = await slideXml(await savePresentation(pres), 0);
    expect(cleared).not.toContain('hlinkClick');
    expect(cleared).toContain('hlinkHover');
    expect(cleared).toContain('a:extLst');
  });
});
