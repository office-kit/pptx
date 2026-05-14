// Two small additions: slide visibility (`show="0"`) and stroke dash
// patterns (`<a:prstDash>`).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  addSlideShape,
  getSlides,
  inches,
  isSlideHidden,
  loadPresentation,
  savePresentation,
  setShapeStroke,
  setShapeStrokeDash,
  setSlideHidden,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await Presentation.load(bytes);
  const pkg = _internalPackageOf(pres);
  const part = pkg.parts.find((p) => p.name === `/ppt/slides/slide${slideIndex + 1}.xml`);
  if (!part) throw new Error(`slide${slideIndex + 1}.xml not found`);
  return new TextDecoder().decode(part.data);
};

describe('fn API: slide visibility', () => {
  it('toggles <p:sld show="0"/> correctly', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(isSlideHidden(slide)).toBe(false);

    setSlideHidden(slide, true);
    expect(isSlideHidden(slide)).toBe(true);
    expect(await slideXml(await savePresentation(pres), 0)).toContain('show="0"');

    setSlideHidden(slide, false);
    expect(isSlideHidden(slide)).toBe(false);
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('show="0"');
  });
});

describe('fn API: setShapeStrokeDash', () => {
  it('writes <a:prstDash val="..."/>', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeStroke(shape, { color: '#000000', widthEmu: 12700 });
    setShapeStrokeDash(shape, 'dash');
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:prstDash val="dash"');
  });

  it('reapplying replaces the prior dash choice', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'ellipse', x: inches(0), y: inches(0), w: inches(2), h: inches(2),
    });
    setShapeStroke(shape, { color: '#FF0000' });
    setShapeStrokeDash(shape, 'dot');
    setShapeStrokeDash(shape, 'dashDot');
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('val="dashDot"');
    expect(xml).not.toContain('val="dot"');
  });
});
