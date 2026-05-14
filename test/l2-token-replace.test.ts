// Level-2 token-replacement smoke test.
//
// Scenario:
//   1. Load `one-text-slide.pptx`.
//   2. Use `setText` to put `Hello, {{name}}! Welcome to {{event}}.` on
//      the title placeholder.
//   3. Call `pres.replaceTokens({ name: 'Alice', event: 'Re:Invent' })`.
//   4. Save → reload → confirm the rendered text is `Hello, Alice! Welcome
//      to Re:Invent.`
//
// Confirms the "fill the template" workflow: load template → swap tokens →
// save, with all formatting preserved (since replaceTokens only touches the
// text content of `<a:t>` elements).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('L2: token-based template fill', () => {
  it('substitutes {{tokens}} in a placeholder and survives round-trip', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-text-slide.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected one slide');
    const title = slide.findPlaceholder('title');
    if (!title) throw new Error('expected title placeholder');

    title.setText('Hello, {{name}}! Welcome to {{event}}.');
    const n = pres.replaceTokens({ name: 'Alice', event: 'Re:Invent' });
    expect(n).toBeGreaterThanOrEqual(1);

    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides[0]?.findPlaceholder('title')?.text).toBe(
      'Hello, Alice! Welcome to Re:Invent.',
    );
  });

  it('leaves unknown tokens untouched', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-text-slide.pptx')));
    const title = pres.slides[0]?.findPlaceholder('title');
    if (!title) throw new Error('expected title');
    title.setText('Hello, {{name}}! Status: {{status}}.');

    pres.replaceTokens({ name: 'Alice' }); // no `status`

    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides[0]?.findPlaceholder('title')?.text).toBe(
      'Hello, Alice! Status: {{status}}.',
    );
  });

  it('returns 0 when no tokens match', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-text-slide.pptx')));
    expect(pres.replaceTokens({ irrelevant: 'value' })).toBe(0);
  });

  it('handles multiple slides in one call', async () => {
    const pres = await Presentation.load(await readFile(fixture('two-slides.pptx')));
    const [s1, s2] = pres.slides;
    if (!s1 || !s2) throw new Error('expected two slides');
    s1.findPlaceholder('title')?.setText('Slide 1: {{team}} review');
    s2.findPlaceholder('title')?.setText('Slide 2: {{team}} action items');

    const n = pres.replaceTokens({ team: 'Platform' });
    expect(n).toBe(2);

    const reloaded = await Presentation.load(await pres.save());
    expect(reloaded.slides[0]?.findPlaceholder('title')?.text).toBe('Slide 1: Platform review');
    expect(reloaded.slides[1]?.findPlaceholder('title')?.text).toBe(
      'Slide 2: Platform action items',
    );
  });

  it('respects Object.hasOwn (rejects inherited prototype keys)', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-text-slide.pptx')));
    const title = pres.slides[0]?.findPlaceholder('title');
    if (!title) throw new Error('expected title');
    title.setText('Hello, {{constructor}}!');
    pres.replaceTokens({}); // No matching own keys; built-ins must not leak.
    expect(pres.slides[0]?.findPlaceholder('title')?.text).toBe('Hello, {{constructor}}!');
  });
});
