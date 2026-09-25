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
  'context menus support keyboard navigation without editing the underlying selection',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-menu-keyboard-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Table x={1} y={1} width={6} height={3} rows={[["A","B","C"],["D","E","F"],["G","H","I"]]} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const table = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        ).find(isTableShape);
      const values = async () =>
        getTableCells(await table()).map((row) => row.map(getTableCellText));
      const cell = (r, c) =>
        editor.getByRole('button', { name: `${ja ? 'セル' : 'Cell'} ${r}, ${c}`, exact: true });
      const focused = async (r, c) =>
        assert.equal(await cell(r, c).evaluate((node) => node === document.activeElement), true);
      await saved();
      await editor.locator('.hit').first().click();
      await cell(1, 1).click();
      const menuFocus = async (label) =>
        assert.equal(
          await editor
            .getByRole('menuitem', { name: label, exact: false })
            .evaluate((node) => node === document.activeElement),
          true,
        );
      await cell(1, 1).press('Shift+F10');
      await menuFocus('Cut');
      await page.keyboard.press('ArrowUp');
      await menuFocus('Select table');
      await page.keyboard.press('Home');
      await menuFocus('Cut');
      await page.keyboard.press('ArrowDown');
      await menuFocus('Copy');
      await page.keyboard.press('ArrowDown');
      await menuFocus('Clear cell text');
      // Disabled Paste is skipped; menu keys must not edit the underlying cell.
      await page.keyboard.press('Delete');
      assert.equal((await values())[0][0], 'A');
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 1);
      await page.keyboard.press('Enter');
      await saved();
      assert.equal((await values())[0][0], '');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await cell(1, 1).click();
      await cell(1, 1).press('Shift+F10');
      await page.keyboard.press('End');
      await menuFocus('Select table');
      await page.keyboard.press('ArrowUp');
      await menuFocus('Select all cells');
      await page.keyboard.press('Enter');
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 9);
      await cell(1, 1).press('Shift+F10');
      await page.keyboard.press('Escape');
      assert.equal(await editor.getByRole('menu').count(), 0);
      await focused(1, 1);
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 9);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await cell(2, 2).click();
      await cell(2, 2).press('Shift+F10');
      await menuFocus('切り取り');
      await page.keyboard.press('ArrowDown');
      await page.keyboard.press('ArrowDown');
      await menuFocus('セルの文字列を消去');
      await page.keyboard.press('Enter');
      await saved();
      assert.equal((await values())[1][1], '');
      await page.reload();
      await saved();
      assert.deepEqual(await values(), [
        ['A', 'B', 'C'],
        ['D', '', 'F'],
        ['G', 'H', 'I'],
      ]);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
