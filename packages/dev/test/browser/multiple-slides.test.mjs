import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlides, getSlideText, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'multiple slides support range/toggle selection, clipboard, duplicate, delete, reorder and history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-multiple-slides-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['A 日本語','B','C','D','E'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const thumbs = editor.locator('.thumb-row');
      let locale = 'en';
      const saved = () =>
        editor
          .getByText(locale === 'en' ? 'Saved to this project' : 'このプロジェクトに保存済み', {
            exact: true,
          })
          .waitFor();
      const order = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        ).map(getSlideText);
      const selected = () =>
        editor
          .locator('.thumb-row[aria-pressed="true"]')
          .evaluateAll((nodes) => nodes.map((node) => Number(node.dataset.slideIndex)));
      const undo = async () => {
        await page.keyboard.press('Control+z');
        await saved();
      };
      await saved();
      await thumbs.nth(1).click();
      await thumbs.nth(3).click({ modifiers: ['Shift'] });
      assert.deepEqual(await selected(), [1, 2, 3]);
      await editor.getByText('Selected slides: 3', { exact: true }).waitFor();
      await thumbs.nth(2).click({ modifiers: ['Meta'] });
      assert.deepEqual(await selected(), [1, 3]);
      // A context menu on a selected thumbnail must retain the whole selection.
      await thumbs.nth(1).click({ button: 'right' });
      assert.deepEqual(await selected(), [1, 3]);
      await editor.getByRole('menuitem', { name: /^Copy/ }).click();
      await thumbs.nth(4).click();
      await thumbs.nth(4).press('Control+v');
      await thumbs.nth(6).waitFor();
      await saved();
      assert.deepEqual(await order(), ['A 日本語', 'B', 'C', 'D', 'E', 'B', 'D']);
      assert.deepEqual(await selected(), [5, 6]);
      await undo();
      assert.deepEqual(await order(), ['A 日本語', 'B', 'C', 'D', 'E']);
      await page.keyboard.press('Control+Shift+z');
      await thumbs.nth(6).waitFor();
      await saved();
      assert.deepEqual(await selected(), [5, 6]);
      await thumbs.nth(5).press('Delete');
      await saved();
      assert.deepEqual(await order(), ['A 日本語', 'B', 'C', 'D', 'E']);
      await undo();
      assert.deepEqual(await selected(), [5, 6]);
      await thumbs.nth(5).press('Delete');
      await saved();
      await editor.locator('.lang select').selectOption('ja');
      locale = 'ja';
      await thumbs.nth(1).click();
      await thumbs.nth(1).press('Shift+ArrowDown');
      assert.deepEqual(await selected(), [1, 2]);
      await editor.getByText('選択中のスライド: 2', { exact: true }).waitFor();
      await thumbs.nth(2).press('Control+d');
      await thumbs.nth(6).waitFor();
      await saved();
      assert.deepEqual(await order(), ['A 日本語', 'B', 'C', 'B', 'C', 'D', 'E']);
      assert.deepEqual(await selected(), [3, 4]);
      await thumbs.nth(3).press('Alt+ArrowUp');
      await saved();
      assert.deepEqual(await order(), ['A 日本語', 'B', 'B', 'C', 'C', 'D', 'E']);
      assert.deepEqual(await selected(), [2, 3]);
      await thumbs.nth(2).press('Alt+ArrowDown');
      await saved();
      assert.deepEqual(await order(), ['A 日本語', 'B', 'C', 'B', 'C', 'D', 'E']);
      await thumbs.nth(3).dragTo(thumbs.nth(6));
      await saved();
      assert.deepEqual(await order(), ['A 日本語', 'B', 'C', 'D', 'E', 'B', 'C']);
      assert.deepEqual(await selected(), [5, 6]);
      await undo();
      assert.deepEqual(await selected(), [3, 4]);
      assert.equal(await thumbs.nth(3).evaluate((node) => node === document.activeElement), true);
      await page.screenshot({ path: '/tmp/pptx-pr287-multiple-slides-ja.png', fullPage: true });
      await thumbs.nth(3).press('Control+x');
      await saved();
      assert.deepEqual(await order(), ['A 日本語', 'B', 'C', 'D', 'E']);
      await thumbs.nth(4).click();
      await thumbs.nth(4).press('Control+v');
      await thumbs.nth(6).waitFor();
      await saved();
      assert.deepEqual(await order(), ['A 日本語', 'B', 'C', 'D', 'E', 'B', 'C']);
      await page.reload();
      await saved();
      assert.deepEqual(await order(), ['A 日本語', 'B', 'C', 'D', 'E', 'B', 'C']);
      await thumbs.nth(0).click();
      await thumbs.nth(0).press('Control+a');
      assert.equal((await selected()).length, 7);
      await thumbs.nth(0).press('Delete');
      await saved();
      assert.equal(await thumbs.count(), 0);
      await undo();
      assert.equal((await selected()).length, 7);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({
        path: '/tmp/pptx-pr287-multiple-slides-failure.png',
        fullPage: true,
      });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
