import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getParagraphPropertiesEffective,
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
  'table paragraph formatting targets individual paragraphs or cell ranges in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-paragraphs-'));
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
      await saved();
      const bounds = getShapeBounds(await table());
      await editor.locator('.hit').first().click();
      await cell(1, 1).click();
      let panel = editor.getByRole('region', { name: 'Paragraph formatting', exact: true });
      const properties = async (row, col, index = 0) => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getParagraphPropertiesEffective(
          pres,
          getTableCells(getSlideShapes(getSlides(pres)[0]).find(isTableShape))[row][col],
          index,
        );
      };
      await cell(2, 2).click({ modifiers: ['Shift'] });
      await panel.getByLabel('List style', { exact: true }).selectOption('bullet');
      await panel.getByLabel('Line spacing mode', { exact: true }).selectOption('pct');
      await panel.getByLabel('Line spacing value', { exact: true }).fill('1.5');
      await panel.getByLabel('Line spacing value', { exact: true }).press('Tab');
      await saved();
      for (const [r, c] of [
        [0, 0],
        [0, 1],
        [1, 0],
        [1, 1],
      ]) {
        assert.equal((await properties(r, c)).bullet, 'bullet');
        assert.deepEqual((await properties(r, c)).lineSpacing, { kind: 'pct', value: 1.5 });
      }
      assert.equal((await properties(2, 2)).bullet, 'none');
      assert.equal(
        await editor.locator('.paint foreignObject p').filter({ hasText: '•' }).count(),
        4,
      );
      await panel.getByLabel('Apply to paragraphs', { exact: true }).selectOption('1');
      await panel.getByLabel('Paragraph alignment', { exact: true }).selectOption('right');
      await saved();
      assert.equal((await properties(0, 1)).align, 'right');
      assert.equal((await properties(0, 0)).align, null);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await properties(0, 1)).align, null);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      panel = editor.getByRole('region', { name: '段落の書式', exact: true });
      await panel.getByLabel('段落前（pt）', { exact: true }).fill('8');
      await panel.getByLabel('段落前（pt）', { exact: true }).press('Tab');
      await saved();
      assert.equal((await properties(0, 1)).spcBefPts, 8);
      assert.equal((await properties(0, 0)).spcBefPts, null);
      await page.screenshot({ path: '/tmp/pptx-pr287-cell-paragraphs-ja.png' });
      await page.reload();
      await saved();
      assert.equal((await properties(0, 1)).spcBefPts, 8);
      assert.deepEqual(await values(), [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
        ['G', 'H', 'I'],
      ]);
      assert.deepEqual(getShapeBounds(await table()), bounds);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
