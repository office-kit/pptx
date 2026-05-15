// Level-2 token-replacement smoke test.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findSlidePlaceholder,
  getShapeText,
  getSlides,
  loadPresentation,
  replaceTokensInPresentation,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const titleText = (slide: ReturnType<typeof getSlides>[number]): string | undefined => {
  const ph = findSlidePlaceholder(slide, 'title');
  return ph ? getShapeText(ph) : undefined;
};

const setTitle = (slide: ReturnType<typeof getSlides>[number], value: string): void => {
  const ph = findSlidePlaceholder(slide, 'title');
  if (!ph) throw new Error('expected title');
  setShapeText(ph, value);
};

describe('L2: token-based template fill', () => {
  it('substitutes {{tokens}} in a placeholder and survives round-trip', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    setTitle(getSlides(pres)[0]!, 'Hello, {{name}}! Welcome to {{event}}.');
    const n = replaceTokensInPresentation(pres, { name: 'Alice', event: 'Re:Invent' });
    expect(n).toBeGreaterThanOrEqual(1);

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(titleText(getSlides(reloaded)[0]!)).toBe('Hello, Alice! Welcome to Re:Invent.');
  });

  it('leaves unknown tokens untouched', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    setTitle(getSlides(pres)[0]!, 'Hello, {{name}}! Status: {{status}}.');
    replaceTokensInPresentation(pres, { name: 'Alice' });
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(titleText(getSlides(reloaded)[0]!)).toBe('Hello, Alice! Status: {{status}}.');
  });

  it('returns 0 when no tokens match', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    expect(replaceTokensInPresentation(pres, { irrelevant: 'value' })).toBe(0);
  });

  it('handles multiple slides in one call', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [s1, s2] = getSlides(pres);
    if (!s1 || !s2) throw new Error('expected two slides');
    setTitle(s1, 'Slide 1: {{team}} review');
    setTitle(s2, 'Slide 2: {{team}} action items');

    const n = replaceTokensInPresentation(pres, { team: 'Platform' });
    expect(n).toBe(2);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const reSlides = getSlides(reloaded);
    expect(titleText(reSlides[0]!)).toBe('Slide 1: Platform review');
    expect(titleText(reSlides[1]!)).toBe('Slide 2: Platform action items');
  });

  it('respects Object.hasOwn (rejects inherited prototype keys)', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    setTitle(getSlides(pres)[0]!, 'Hello, {{constructor}}!');
    replaceTokensInPresentation(pres, {});
    expect(titleText(getSlides(pres)[0]!)).toBe('Hello, {{constructor}}!');
  });
});
