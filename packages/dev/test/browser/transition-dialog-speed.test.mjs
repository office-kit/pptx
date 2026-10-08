import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { compile, Presentation, Slide } from '@office-kit/pptx-dsl';
import { getSlides, savePresentation, setSlideTransition } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// A transition without spd runs at the schema default, fast (as Cut is saved
// by the reference desktop app); the Slide transition dialog shows that speed, as the
// Transitions tab's Duration does, instead of assuming Medium.
test('the Slide transition dialog reads a missing spd as fast', { timeout: 90000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-transition-dialog-'));
  let preview, browser;
  try {
    const deck = await compile(Presentation({ children: [Slide({}), Slide({})] }));
    setSlideTransition(getSlides(deck)[0], { effect: 'fade' });
    setSlideTransition(getSlides(deck)[1], { effect: 'pull', durationMs: 750 });
    await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
    );
    preview = await startPreview(file);
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
    await page.goto(preview.url);
    const editor = page.frameLocator('#editor-frame');
    await editor.getByText('Saved to this project', { exact: true }).waitFor();
    const speedShown = async (index) => {
      await editor.locator('.thumb-row').nth(index).click({ button: 'right' });
      await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();
      await editor.getByRole('button', { name: 'Slide transition', exact: true }).click();
      const dialog = editor.getByRole('dialog', { name: 'Slide transition', exact: true });
      const speed = await dialog.getByLabel('Transition speed', { exact: true }).inputValue();
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      return speed;
    };
    assert.equal(await speedShown(0), 'fast');
    assert.equal(await speedShown(1), 'med');
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});
