import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  duplicateSlide,
  getSlides,
  isSlideBackgroundGraphicsHidden,
  loadPresentation,
  savePresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'background graphics hide on selected slides, undo, show mixed state and survive reload',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-background-graphics-'));
    let preview, browser;
    try {
      const pres = await loadPresentation(
        await readFile(
          new URL('../../../../test/fixtures/minimal/layout-decoration.pptx', import.meta.url),
        ),
      );
      const slide = getSlides(pres)[0];
      duplicateSlide(pres, slide);
      duplicateSlide(pres, slide);
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(pres));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlides(deck).map(isSlideBackgroundGraphicsHidden);
      };
      const open = async () => {
        await editor.getByRole('tab', { name: 'Design', exact: true }).click();
        await editor
          .getByRole('tabpanel', { name: 'Design', exact: true })
          .getByRole('button', { name: 'Format Background', exact: true })
          .click();
      };
      await saved();
      await open();
      const thumbs = editor.locator('.thumb-row');
      const checkbox = editor.getByRole('checkbox', {
        name: 'Hide Background Graphics',
        exact: true,
      });
      const canvas = editor.locator('.stage .paint');
      await thumbs.nth(0).click();
      assert.match(await canvas.innerHTML(), /TEMPLATE/);
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      await checkbox.check();
      await saved();
      assert.deepEqual(await read(), [true, true, false]);
      assert.doesNotMatch(await canvas.innerHTML(), /TEMPLATE/);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), [false, false, false]);
      assert.match(await canvas.innerHTML(), /TEMPLATE/);
      await checkbox.check();
      await saved();
      await thumbs.nth(2).click({ modifiers: ['Shift'] });
      assert.equal(await checkbox.evaluate((element) => element.indeterminate), true);
      await checkbox.check();
      await saved();
      assert.deepEqual(await read(), [true, true, true]);
      await page.reload();
      await saved();
      await open();
      assert.equal(await checkbox.isChecked(), true);
      assert.doesNotMatch(await canvas.innerHTML(), /TEMPLATE/);
      await thumbs.nth(0).click();
      await checkbox.uncheck();
      await saved();
      assert.deepEqual(await read(), [false, true, true]);
      assert.match(await canvas.innerHTML(), /TEMPLATE/);
      await checkbox.check();
      await saved();
      await editor.getByRole('button', { name: 'Apply to All', exact: true }).click();
      await saved();
      await checkbox.uncheck();
      await saved();
      assert.match(await canvas.innerHTML(), /TEMPLATE/);
      assert.doesNotMatch(await canvas.innerHTML(), /#2E75B6/);
      await editor.getByRole('button', { name: 'Apply to All', exact: true }).click();
      await saved();
      assert.deepEqual(await read(), [false, false, false]);
      assert.match(await canvas.innerHTML(), /#2E75B6/);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
