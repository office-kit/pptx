// `getSlideLayoutShapes` / `getSlideMasterShapes` — non-placeholder decoration
// on a layout / master (logos, bars, watermark text) as render-ready
// SlideShapeData. The key behaviour vs the older flat
// `getSlideLayoutBackgroundShapes` is that PICTURES are included and their
// bytes resolve, because each shape is bound to the layout/master part's rels.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getShapeImageBytes,
  getShapeKind,
  getShapeName,
  getShapeText,
  getSlideLayout,
  getSlideLayoutShapes,
  getSlideMasterShapes,
  getSlides,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const load = async () => {
  const pres = await loadPresentation(await readFile(fixture('layout-decoration.pptx')));
  const layout = getSlideLayout(getSlides(pres)[0]!);
  if (!layout) throw new Error('no layout');
  return { pres, layout };
};

describe('getSlideLayoutShapes', () => {
  it('returns the layout decoration including a picture logo with resolvable bytes', async () => {
    const { pres, layout } = await load();
    const shapes = getSlideLayoutShapes(pres, layout);
    const byName = new Map(shapes.map((s) => [getShapeName(s), s]));

    const logo = byName.get('Logo');
    expect(logo).toBeDefined();
    expect(getShapeKind(logo!)).toBe('picture');
    // The blip rel lives in the LAYOUT's relationship table, not the slide's.
    const bytes = getShapeImageBytes(logo!);
    expect(bytes).not.toBeNull();
    expect(bytes!.length).toBeGreaterThan(0);

    const bar = byName.get('Template Bar');
    expect(bar).toBeDefined();
    expect(getShapeText(bar!)).toBe('TEMPLATE');
  });

  it('excludes placeholders (those render through the slide)', async () => {
    const { pres, layout } = await load();
    const names = getSlideLayoutShapes(pres, layout).map(getShapeName);
    // Only the injected decoration, no Title/Subtitle/etc. placeholders.
    expect(names).toContain('Logo');
    expect(names).toContain('Template Bar');
    expect(names.some((n) => /Title|Placeholder|Footer Placeholder/.test(n))).toBe(false);
  });
});

describe('getSlideMasterShapes', () => {
  it('returns the master decoration', async () => {
    const { pres, layout } = await load();
    const shapes = getSlideMasterShapes(pres, layout);
    expect(shapes.map(getShapeName)).toContain('Master Footer Bar');
  });
});
