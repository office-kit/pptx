// `setShapeHyperlink` and `setShapeRunHyperlink` — optional `tooltip`
// arg that drops a `tooltip="…"` attribute on the emitted
// `<a:hlinkClick>`. Pairs the writers with the existing
// `getShape{,Run}HyperlinkTooltip` readers.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getShapeHyperlinkTooltip,
  getShapeRunHyperlinkTooltip,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeHyperlink,
  setShapeRunHyperlink,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: setShapeHyperlink tooltip', () => {
  it('round-trips a shape-level tooltip through save/reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'click me',
    });
    setShapeHyperlink(tb, 'https://example.com', 'Visit example.com');
    expect(getShapeHyperlinkTooltip(tb)).toBe('Visit example.com');

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    expect(getShapeHyperlinkTooltip(reShapes[reShapes.length - 1]!)).toBe('Visit example.com');
  });

  it('omits the tooltip attribute when no tooltip arg is passed', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'no tip',
    });
    setShapeHyperlink(tb, 'https://example.com');
    expect(getShapeHyperlinkTooltip(tb)).toBeNull();
  });
});

describe('fn API: setShapeRunHyperlink tooltip', () => {
  it('round-trips a per-run tooltip through save/reload', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(3),
      h: inches(1),
      text: 'run',
    });
    setShapeRunHyperlink(tb, 0, 0, 'https://example.com', 'per-run tip');
    expect(getShapeRunHyperlinkTooltip(tb, 0, 0)).toBe('per-run tip');

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    const reShapes = getSlideShapes(getSlides(reloaded)[0]!);
    expect(getShapeRunHyperlinkTooltip(reShapes[reShapes.length - 1]!, 0, 0)).toBe('per-run tip');
  });
});
