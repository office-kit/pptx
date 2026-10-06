import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeDescription, getSlides, getSlideShapes, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// A 1×1 PNG.
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const DECK = `import { Presentation, Slide, Image, Text } from '@office-kit/pptx-dsl';
const png = Uint8Array.from(atob('${PNG}'), (c) => c.charCodeAt(0));
export default (
  <Presentation>
    <Slide><Text x={1} y={1} width={4} height={1}>One</Text></Slide>
    <Slide><Image data={png} x={1} y={1} width={1} height={1} /></Slide>
  </Presentation>
);
`;

test(
  'status bar shows language, accessibility issues and presentation views',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-status-bar-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const context = await browser.newContext({
        locale: 'en-JP',
        viewport: { width: 1400, height: 900 },
      });
      const page = await context.newPage();
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const bar = editor.locator('.statusbar');
      await bar.getByText('English (Japan)', { exact: true }).waitFor();

      // Both slides lack a title and the picture lacks alternative text.
      const a11y = bar.getByRole('button', { name: 'Accessibility: Investigate', exact: true });
      await a11y.click();
      const issues = editor.getByRole('dialog', { name: 'Accessibility', exact: true });
      assert.equal(await issues.getByRole('button').count(), 3);
      await issues.getByRole('button', { name: /Missing alternative text/ }).click();
      await issues.waitFor({ state: 'detached' });
      await editor.getByText('Slide 2 of 2', { exact: true }).waitFor();
      assert.equal(await editor.locator('.hit.selected').count(), 1);

      // Reading View presents the current slide in the window, without full screen.
      await bar.getByRole('button', { name: 'Reading View', exact: true }).click();
      await page.waitForFunction(() => document.body.classList.contains('presenting'));
      assert.equal(await page.evaluate(() => document.fullscreenElement), null);
      await page.waitForFunction(
        () => document.querySelector('#present-count')?.textContent === 'Slide 2 of 2',
      );
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => !document.body.classList.contains('presenting'));

      const bytes = new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer());
      const picture = getSlideShapes(getSlides(await loadPresentation(bytes))[1])[0];
      assert.equal(getShapeDescription(picture) ?? '', '');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
