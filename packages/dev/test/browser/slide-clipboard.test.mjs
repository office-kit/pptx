import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideText,
  getSlideShapes,
  getSlideNotes,
  getShapeChartSpec,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'slide clipboard preserves content, insertion order and history in English and Japanese',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-slide-clipboard-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text,Chart} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes={'発表メモ\\nSpeaker notes'}><Text x={1} y={1} width={7} height={1}>Original 日本語</Text><Chart x={1} y={2} width={5} height={3} spec={{kind:'column',categories:['日本語','English'],series:[{name:'Sales',values:[3,7]}]}} /></Slide><Slide><Text x={1} y={1} width={7} height={1}>Second</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let locale = 'en';
      const saved = () =>
        editor
          .getByText(locale === 'en' ? 'Saved to this project' : 'このプロジェクトに保存済み', {
            exact: true,
          })
          .waitFor();
      const slides = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        );
      const order = async () => (await slides()).map(getSlideText);
      const count = async (n) => {
        await editor
          .locator('.thumb-row')
          .nth(n - 1)
          .waitFor();
        await saved();
        assert.equal(await editor.locator('.thumb-row').count(), n);
      };
      await saved();
      await editor.locator('.thumb-row').first().click({ button: 'right' });
      await editor.getByRole('menuitem', { name: /^Copy/ }).click();
      await editor.locator('.thumb-row').nth(1).click({ button: 'right' });
      await editor.getByRole('menuitem', { name: /^Paste/ }).click();
      await count(3);
      assert.deepEqual(await order(), ['Original 日本語', 'Second', 'Original 日本語']);
      let result = await slides();
      assert.equal(getSlideNotes(result[2]), '発表メモ\nSpeaker notes');
      assert.deepEqual(getShapeChartSpec(getSlideShapes(result[2])[1]).series[0].values, [3, 7]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await order(), ['Original 日本語', 'Second']);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await count(3);
      await editor.locator('.lang select').selectOption('ja');
      locale = 'ja';
      await editor.locator('.thumb-row').first().click({ button: 'right' });
      await editor.getByRole('menuitem', { name: /^切り取り/ }).click();
      await saved();
      assert.deepEqual(await order(), ['Second', 'Original 日本語']);
      await editor.locator('.thumb-row').first().click();
      await editor.locator('.thumb-row').first().press('Control+v');
      await count(3);
      assert.deepEqual(await order(), ['Second', 'Original 日本語', 'Original 日本語']);
      await editor.locator('.thumb-row').nth(1).click();
      await editor.locator('.thumb-row').nth(1).press('Control+c');
      await editor.locator('.thumb-row').nth(1).press('Delete');
      await saved();
      await editor.locator('.thumb-row').first().click();
      await editor.locator('.thumb-row').first().press('Control+v');
      await count(3);
      assert.deepEqual(await order(), ['Second', 'Original 日本語', 'Original 日本語']);
      await page.screenshot({ path: '/tmp/pptx-pr287-slide-clipboard-ja.png', fullPage: true });
      await page.reload();
      await count(3);
      result = await slides();
      assert.equal(getSlideNotes(result[1]), '発表メモ\nSpeaker notes');
      assert.deepEqual(getShapeChartSpec(getSlideShapes(result[1])[1]).categories, [
        '日本語',
        'English',
      ]);
      // The captured slide also survives replacing the entire document.
      await editor.locator('.thumb-row').nth(1).click();
      await editor.locator('.thumb-row').nth(1).press('Control+c');
      await editor.getByTitle('新規', { exact: true }).click();
      await saved();
      await editor.locator('.thumb-row').first().click();
      await editor.locator('.thumb-row').first().press('Control+v');
      await count(2);
      result = await slides();
      assert.equal(getSlideText(result[1]), 'Original 日本語');
      assert.equal(getSlideNotes(result[1]), '発表メモ\nSpeaker notes');
      assert.deepEqual(getShapeChartSpec(getSlideShapes(result[1])[1]).series[0].values, [3, 7]);
      // Cutting the final slide must still leave a usable clipboard.
      await editor.locator('.thumb-row').first().click();
      await editor.locator('.thumb-row').first().press('Delete');
      await saved();
      await editor.locator('.thumb-row').first().click();
      await editor.locator('.thumb-row').first().press('Control+x');
      await saved();
      assert.equal(await editor.locator('.thumb-row').count(), 0);
      await page.keyboard.press('Control+v');
      await count(1);
      assert.deepEqual(await order(), ['Original 日本語']);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({
        path: '/tmp/pptx-pr287-slide-clipboard-failure.png',
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
