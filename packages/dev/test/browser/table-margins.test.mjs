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
  getTableCellMargins,
  isTableShape,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'cell margins preserve other sides and support mixed ranges, resets and saved Japanese edits',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-margins-'));
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
      await saved();
      const bounds = getShapeBounds(await table());
      await editor.locator('.hit').first().click();
      const cell = (r, c) =>
        editor.getByRole('button', { name: `${ja ? 'セル' : 'Cell'} ${r}, ${c}`, exact: true });
      const margin = (label) => editor.getByLabel(label, { exact: true });
      const setMargin = async (label, value) => {
        await margin(label).fill(value);
        await margin(label).press('Tab');
        await saved();
      };
      const margins = async () =>
        getTableCells(await table()).map((row) => row.map(getTableCellMargins));
      const initial = await margins();
      await cell(1, 1).click();
      await editor.getByText('Cell margins', { exact: true }).click();
      await setMargin('Left margin (points)', '18');
      await setMargin('Right margin (points)', '9.5');
      assert.equal((await margins())[0][0].left, 18 * 12700);
      assert.equal((await margins())[0][0].right, 9.5 * 12700);
      assert.equal((await margins())[0][0].top, initial[0][0].top);
      await cell(2, 2).click({ modifiers: ['Shift'] });
      assert.equal(await margin('Left margin (points)').getAttribute('placeholder'), 'Mixed');
      assert.equal(await margin('Left margin (points)').inputValue(), '');
      await setMargin('Top margin (points)', '12');
      let current = await margins();
      for (const row of current.slice(0, 2))
        for (const cell of row.slice(0, 2)) assert.equal(cell.top, 12 * 12700);
      assert.equal(current[0][0].left, 18 * 12700);
      assert.equal(current[1][1].left, initial[1][1].left);
      assert.deepEqual(current[2], initial[2]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await margins())[0][0].top, initial[0][0].top);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await margin('Top margin (points)').fill('-1');
      await margin('Top margin (points)').press('Tab');
      assert.equal(
        await margin('Top margin (points)').evaluate((node) => node.validity.valid),
        false,
      );
      assert.equal((await margins())[0][0].top, 12 * 12700);
      await setMargin('Top margin (points)', '12');
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await setMargin('下余白（pt）', '0');
      await editor.getByRole('button', { name: '左余白を標準に戻す', exact: true }).click();
      await saved();
      current = await margins();
      assert.equal(current[0][0].left, null);
      assert.equal(current[0][0].right, 9.5 * 12700);
      assert.equal(current[0][0].bottom, 0);
      const expected = current;
      await page.screenshot({ path: '/tmp/pptx-pr287-cell-margins-ja.png', fullPage: true });
      await editor.getByRole('button', { name: 'セル内の余白をリセット', exact: true }).click();
      await saved();
      current = await margins();
      for (const row of current.slice(0, 2))
        for (const cell of row.slice(0, 2))
          assert.deepEqual(cell, { left: null, right: null, top: null, bottom: null });
      assert.deepEqual(current[2], initial[2]);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await margins(), expected);
      assert.deepEqual(getShapeBounds(await table()), bounds);
      assert.deepEqual(await values(), [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
        ['G', 'H', 'I'],
      ]);
      await page.reload();
      await saved();
      assert.deepEqual(await margins(), expected);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
