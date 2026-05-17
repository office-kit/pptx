// getShapeRunHyperlink — per-run hyperlink lookup. Per-run setting
// goes through the existing setShapeHyperlink (applies to every run)
// for now; this reader makes it observable per run.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeRunHyperlink,
  getSlides,
  inches,
  loadPresentation,
  setShapeHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: getShapeRunHyperlink', () => {
  it('returns the URL applied via setShapeHyperlink', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'click me',
    });
    setShapeHyperlink(tb, 'https://example.com/');
    expect(getShapeRunHyperlink(tb, 0, 0)).toBe('https://example.com/');
  });

  it('returns null on a run with no hyperlink', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'plain',
    });
    expect(getShapeRunHyperlink(tb, 0, 0)).toBeNull();
  });

  it('returns null after the hyperlink is cleared', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'toggle',
    });
    setShapeHyperlink(tb, 'https://example.com/');
    setShapeHyperlink(tb, null);
    expect(getShapeRunHyperlink(tb, 0, 0)).toBeNull();
  });
});
