import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getShapeBounds,
  getShapeText,
  getTableCells,
  getTableCellText,
  isTableShape,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'canvas context targets shapes, ranges, cells and empty space without right-button gestures',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-canvas-context-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={2} height={1}>A</Text><Text x={4} y={1} width={2} height={1}>B</Text><Table x={1} y={3} width={6} height={2} rows={[["C","D"],["E","F"]]} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const shapes = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        );
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const hits = editor.locator('.hit');
      const selected = editor.locator('.hit.selected');
      const menu = editor.getByRole('menu');
      await saved();
      const original = (await shapes()).map(getShapeBounds);
      await hits.nth(0).click();
      await editor.locator('.stage').click({ button: 'right', position: { x: 3, y: 3 } });
      assert.equal(await selected.count(), 0);
      assert.equal(await menu.getByRole('menuitem', { name: 'Delete', exact: false }).count(), 0);
      await menu.getByRole('menuitem', { name: 'Select all', exact: false }).click();
      assert.equal(await selected.count(), 3);
      await hits.nth(0).click({ button: 'right', modifiers: ['Shift'] });
      assert.equal(await selected.count(), 3);
      await page.keyboard.press('Escape');
      // A right-button drag must never modify shape geometry.
      const box = await hits.nth(0).boundingBox();
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.mouse.down({ button: 'right' });
      await page.mouse.move(box.x + box.width / 2 + 40, box.y + box.height / 2 + 20, { steps: 5 });
      await page.mouse.up({ button: 'right' });
      await page.keyboard.press('Escape');
      assert.deepEqual((await shapes()).map(getShapeBounds), original);
      await editor.locator('.stage').click({ position: { x: 3, y: 3 } });
      await hits.nth(0).click();
      await hits.nth(1).click({ button: 'right' });
      assert.equal(await selected.count(), 1);
      await menu.getByRole('menuitem', { name: 'Delete', exact: false }).click();
      await saved();
      assert.deepEqual((await shapes()).filter((s) => !isTableShape(s)).map(getShapeText), ['A']);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      const tableBox = await hits.nth(2).boundingBox();
      const cellContext = async (row, col) => {
        await page.mouse.click(
          tableBox.x + (tableBox.width * (col + 0.5)) / 2,
          tableBox.y + (tableBox.height * (row + 0.5)) / 2,
          { button: 'right' },
        );
      };
      await cellContext(1, 1);
      assert.equal(
        await editor
          .getByRole('button', { name: 'セル 2, 2', exact: true })
          .getAttribute('aria-pressed'),
        'true',
      );
      await menu.getByRole('menuitem', { name: 'セルの文字列を消去', exact: false }).click();
      await saved();
      assert.deepEqual(
        getTableCells((await shapes()).find(isTableShape)).map((row) => row.map(getTableCellText)),
        [
          ['C', 'D'],
          ['E', ''],
        ],
      );
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      await editor.getByRole('button', { name: 'セル 1, 1', exact: true }).click();
      await editor
        .getByRole('button', { name: 'セル 2, 2', exact: true })
        .click({ modifiers: ['Shift'] });
      await cellContext(0, 1);
      assert.equal(await editor.locator('.cell-grid button[aria-pressed="true"]').count(), 4);
      await page.screenshot({ path: '/tmp/pptx-pr287-canvas-context-ja.png', fullPage: true });
      await menu.getByRole('menuitem', { name: 'セルの文字列を消去', exact: false }).click();
      await saved();
      await page.reload();
      await saved();
      assert.deepEqual(
        getTableCells((await shapes()).find(isTableShape)).map((row) => row.map(getTableCellText)),
        [
          ['', ''],
          ['', ''],
        ],
      );
      assert.deepEqual((await shapes()).map(getShapeBounds), original);
      await hits.nth(0).dblclick();
      const nativeMenuAllowed = await editor
        .locator('.inline-edit')
        .evaluate((node) =>
          node.dispatchEvent(
            new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2 }),
          ),
        );
      assert.equal(nativeMenuAllowed, true);
      assert.equal(await menu.count(), 0);
      await page.keyboard.press('Escape');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
