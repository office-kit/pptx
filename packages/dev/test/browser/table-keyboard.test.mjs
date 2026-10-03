import { installRichTextSelection } from '../helpers/rich-text.mjs';
import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getTableCells,
  getTableCellText,
  isTableShape,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'table Tab navigation and spreadsheet paste expand, undo and persist in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-keyboard-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Table x={1} y={1} width={6} height={2} rows={[["A","B"],["C","D"]]} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const values = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getTableCells(getSlideShapes(getSlides(pres)[0]).find(isTableShape)).map((row) =>
          row.map(getTableCellText),
        );
      };
      const inline = editor.locator('.inline-edit');
      await saved();
      const hit = editor.locator('.hit').first();
      const bounds = await hit.boundingBox();
      await hit.dblclick({ position: { x: bounds.width / 4, y: bounds.height / 4 } });
      await inline.fill('日本語');
      await inline.press('Tab');
      assert.equal(await inline.textContent(), 'B');
      await inline.press('Shift+Tab');
      assert.equal(await inline.textContent(), '日本語');
      await inline.press('Shift+Tab');
      assert.equal(await inline.textContent(), '日本語');
      await inline.press('Tab');
      await inline.press('Tab');
      assert.equal(await inline.textContent(), 'C');
      await inline.press('Tab');
      await inline.fill('Last');
      await inline.press('Tab');
      assert.equal(await inline.textContent(), '');
      await inline.press('Escape');
      await saved();
      assert.deepEqual(await values(), [
        ['日本語', 'B'],
        ['C', 'Last'],
        ['', ''],
      ]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await values(), [
        ['日本語', 'B'],
        ['C', 'D'],
      ]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await hit.dblclick({ position: { x: bounds.width * 0.75, y: bounds.height / 2 } });
      // Select the second row, second column explicitly before entering edit mode.
      await inline.press('Escape');
      await editor.getByRole('button', { name: 'Cell 2, 2', exact: true }).click();
      await editor.getByRole('button', { name: 'Cell 2, 2', exact: true }).press('Enter');
      assert.equal(await inline.textContent(), 'Last');
      await inline.evaluate((node) => window.selectEditorText(node, 0, node.textContent.length));
      await inline.evaluate((node) => {
        const data = new DataTransfer();
        data.setData('text/plain', '"one\ttwo"');
        node.dispatchEvent(
          new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
        );
      });
      assert.equal(await inline.textContent(), 'one\ttwo');
      await inline.evaluate((node) => {
        const data = new DataTransfer();
        data.setData('text/plain', '"改行\nEnglish"\t"say ""Hi"""\r\n終わり\t42\r\n');
        node.dispatchEvent(
          new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
        );
      });
      assert.equal(await inline.textContent(), '改行\nEnglish');
      await inline.press('Escape');
      await saved();
      const expected = [
        ['日本語', 'B', ''],
        ['C', '改行\nEnglish', 'say "Hi"'],
        ['', '終わり', '42'],
      ];
      assert.deepEqual(await values(), expected);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await values(), [
        ['日本語', 'B'],
        ['C', 'Last'],
        ['', ''],
      ]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await saved();
      await page.screenshot({ path: '/tmp/pptx-pr287-table-keyboard-ja.png', fullPage: true });
      await page.reload();
      await saved();
      assert.deepEqual(await values(), expected);
      await editor.locator('.hit').first().click();
      await editor.getByRole('button', { name: 'セル 1, 1', exact: true }).click();
      await editor
        .getByRole('button', { name: 'セル 1, 2', exact: true })
        .click({ modifiers: ['Shift'] });
      await editor.getByRole('button', { name: 'セルを結合', exact: true }).click();
      await saved();
      await editor.getByRole('button', { name: 'セル 1, 1', exact: true }).press('Enter');
      assert.equal(await inline.textContent(), '日本語\nB');
      await inline.press('Tab');
      assert.equal(await inline.textContent(), '');
      await inline.press('Tab');
      assert.equal(await inline.textContent(), 'C');
      await inline.press('Shift+Tab');
      await inline.press('Shift+Tab');
      assert.equal(await inline.textContent(), '日本語\nB');
      await inline.evaluate((node) => {
        const data = new DataTransfer();
        data.setData('text/plain', 'X\tY');
        node.dispatchEvent(
          new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
        );
      });
      await editor
        .getByText('複数セルを貼り付ける前に結合セルを分割してください', { exact: true })
        .waitFor();
      assert.equal(await inline.textContent(), '日本語\nB');
      await inline.press('Escape');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await values(), expected);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({
        path: '/tmp/pptx-pr287-table-keyboard-failure.png',
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
