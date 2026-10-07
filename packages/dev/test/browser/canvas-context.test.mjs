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
  isShapeLocked,
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
      // Mac PowerPoint's object menu, in its order and with its separators
      // (native capture 2026-10-07). It has no Delete; the key deletes the
      // selection. Commands the editor cannot perform are disabled.
      const entries = await menu
        .locator(':scope > .ctx-item, :scope > .branch > .ctx-item, :scope > .ctx-sep')
        .evaluateAll((nodes) =>
          nodes.map((node) =>
            node.classList.contains('ctx-sep')
              ? '----'
              : `${node.textContent.replace(/[›⌘⌥⇧].*$/, '').trim()}${node.disabled ? ' (disabled)' : ''}`,
          ),
        );
      assert.deepEqual(entries, [
        'Cut',
        'Copy',
        'Paste (disabled)',
        '----',
        'Edit Text',
        'Edit Points (disabled)',
        '----',
        'Reorder Objects (disabled)',
        'Reorder Overlapping Objects (disabled)',
        '----',
        'Group',
        'Bring to Front',
        'Send to Back',
        'Lock',
        '----',
        'Hyperlink...',
        '----',
        'Save as Picture... (disabled)',
        '----',
        'Translate... (disabled)',
        '----',
        'View Alt Text...',
        'Set as Default Shape Style (disabled)',
        'Size and Position...',
        'Format Shape...',
        '----',
        'Action Settings...',
        '----',
        'New Comment',
      ]);
      // Mac menus use 24 pt items and 11 pt separators.
      assert.deepEqual(
        await menu.evaluate((node) => [
          node.querySelector(':scope > .ctx-item').getBoundingClientRect().height,
          (() => {
            const sep = node.querySelector(':scope > .ctx-sep');
            const style = getComputedStyle(sep);
            return (
              sep.getBoundingClientRect().height +
              parseFloat(style.marginTop) +
              parseFloat(style.marginBottom)
            );
          })(),
        ]),
        [24, 11],
      );
      await menu.getByRole('menuitem', { name: 'Group', exact: true }).hover();
      assert.deepEqual(
        await editor
          .getByRole('menu', { name: 'Group', exact: true })
          .locator(':scope > .ctx-item, :scope > .ctx-sep')
          .evaluateAll((nodes) =>
            nodes.map((node) => (node.classList.contains('ctx-sep') ? '----' : node.textContent)),
          ),
        ['Group', 'Regroup', '----', 'Ungroup'],
      );
      // Lock toggles PowerPoint's object locks; the menu then offers Unlock.
      const locked = async (expected) => {
        const deadline = Date.now() + 10000;
        while (isShapeLocked((await shapes())[1]) !== expected) {
          assert.ok(Date.now() < deadline, `shape lock did not become ${expected}`);
          await new Promise((resolve) => setTimeout(resolve, 100));
        }
      };
      await menu.getByRole('menuitem', { name: 'Lock', exact: true }).click();
      await locked(true);
      await hits.nth(1).click({ button: 'right' });
      await menu.getByRole('menuitem', { name: 'Unlock', exact: true }).click();
      await locked(false);
      await hits.nth(1).click({ button: 'right' });
      await page.keyboard.press('Escape');
      await page.keyboard.press('Delete');
      await saved();
      assert.deepEqual((await shapes()).filter((s) => !isTableShape(s)).map(getShapeText), ['A']);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      // The cell menu has no Format item, so open the pane on the slide (before
      // measuring the table, as the pane narrows the canvas); it then follows
      // the selection to the cell.
      await editor.locator('.stage').click({ button: 'right', position: { x: 3, y: 3 } });
      await menu.getByRole('menuitem', { name: '背景の書式設定...', exact: true }).click();
      // Wait for the canvas to refit to the narrower area before measuring.
      let tableBox = await hits.nth(2).boundingBox();
      for (;;) {
        await page.waitForTimeout(100);
        const next = await hits.nth(2).boundingBox();
        if (next.x === tableBox.x && next.width === tableBox.width) break;
        tableBox = next;
      }
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
        .locator('.canvas-shell .inline-edit')
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
