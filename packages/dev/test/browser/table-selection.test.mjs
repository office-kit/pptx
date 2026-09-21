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
  getShapeBounds,
  isTableShape,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'cell arrow selection and Delete preserve the table, geometry and undoable ranges',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-selection-'));
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
      const bounds = getShapeBounds(await table());
      await editor.locator('.hit').first().click();
      await cell(1, 1).click();
      await cell(1, 1).press('ArrowRight');
      await focused(1, 2);
      await cell(1, 2).press('ArrowDown');
      await focused(2, 2);
      await cell(2, 2).press('Delete');
      await saved();
      assert.deepEqual(await values(), [
        ['A', 'B', 'C'],
        ['D', '', 'F'],
        ['G', 'H', 'I'],
      ]);
      assert.deepEqual(getShapeBounds(await table()), bounds);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await values())[1][1], 'E');
      await cell(1, 1).click();
      await cell(1, 1).press('Shift+ArrowRight');
      await cell(1, 2).press('Shift+ArrowDown');
      await focused(2, 2);
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 4);
      await cell(2, 2).press('Backspace');
      await saved();
      assert.deepEqual(await values(), [
        ['', '', 'C'],
        ['', '', 'F'],
        ['G', 'H', 'I'],
      ]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      // A toolbar focus must keep the same range for global cell shortcuts.
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).evaluate((node) => node.focus());
      await page.keyboard.press('Shift+ArrowRight');
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 6);
      await page.keyboard.press('Shift+ArrowLeft');
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 4);
      await page.keyboard.press('Delete');
      await saved();
      assert.deepEqual(await values(), [
        ['', '', 'C'],
        ['', '', 'F'],
        ['G', 'H', 'I'],
      ]);
      assert.deepEqual(getShapeBounds(await table()), bounds);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 4);
      // Right-click inside a range keeps it; Delete in the menu clears only text.
      await cell(2, 2).click({ button: 'right' });
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 4);
      await editor.getByRole('menuitem', { name: 'Clear cell text', exact: false }).click();
      await saved();
      assert.deepEqual(await values(), [
        ['', '', 'C'],
        ['', '', 'F'],
        ['G', 'H', 'I'],
      ]);
      assert.deepEqual(getShapeBounds(await table()), bounds);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      // Right-click outside the range targets that cell. Select all stays inside the table.
      await cell(3, 3).click({ button: 'right' });
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 1);
      await editor.getByRole('menuitem', { name: 'Select all cells', exact: false }).click();
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 9);
      await cell(1, 1).click();
      await cell(1, 1).press('ControlOrMeta+a');
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 9);
      await cell(1, 1).click();
      await cell(2, 2).click({ modifiers: ['Shift'] });
      await editor.getByRole('button', { name: 'Merge cells', exact: true }).click();
      await saved();
      await cell(1, 1).click();
      await cell(1, 1).press('ArrowRight');
      await focused(1, 3);
      await cell(1, 3).press('ArrowDown');
      await cell(2, 3).press('ArrowLeft');
      await focused(1, 1);
      await cell(1, 1).press('ArrowDown');
      await focused(3, 1);
      await cell(3, 1).press('ArrowRight');
      await cell(3, 2).press('ArrowUp');
      await focused(1, 1);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      // Leave inline editing, then exercise the global canvas cell shortcuts.
      await cell(1, 1).press('Enter');
      await editor.locator('.inline-edit').press('Escape');
      await page.keyboard.press('ArrowDown');
      assert.equal(await cell(3, 1).getAttribute('aria-pressed'), 'true');
      await page.keyboard.press('Delete');
      await saved();
      assert.equal((await values())[2][0], '');
      assert.deepEqual(getShapeBounds(await table()), bounds);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await values())[2][0], 'G');
      const mergedValues = await values();
      await cell(1, 1).click({ button: 'right' });
      await editor.getByRole('menuitem', { name: 'すべてのセルを選択', exact: false }).click();
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 6);
      await cell(3, 3).click({ button: 'right' });
      assert.equal(
        await editor.getByRole('menu').evaluate((node) => {
          const bounds = node.getBoundingClientRect();
          return (
            bounds.left >= 0 &&
            bounds.top >= 0 &&
            bounds.right <= innerWidth &&
            bounds.bottom <= innerHeight
          );
        }),
        true,
      );
      await page.screenshot({ path: '/tmp/pptx-pr287-cell-context-ja.png' });
      await editor.getByRole('menuitem', { name: 'セルの文字列を消去', exact: false }).click();
      await saved();
      assert.ok((await values()).flat().every((value) => value === ''));
      assert.deepEqual(getShapeBounds(await table()), bounds);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await values(), mergedValues);
      await page.reload();
      await saved();
      assert.equal((await values())[2][0], 'G');
      assert.deepEqual(getShapeBounds(await table()), bounds);
      await editor.locator('.hit').first().click();
      await cell(1, 1).click();
      await cell(1, 1).click({ button: 'right' });
      await editor.getByRole('menuitem', { name: '表全体を選択', exact: true }).click();
      await page.keyboard.press('Delete');
      await saved();
      assert.equal(await table(), undefined);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(getShapeBounds(await table()), bounds);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
