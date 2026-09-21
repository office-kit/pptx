import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeChartSpec, getSlideShapes, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'chart blank values support bilingual editing, rendering and saved history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-chart-blanks-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide /></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = true;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () =>
        getShapeChartSpec(
          getSlideShapes(
            getSlides(
              await loadPresentation(
                new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
              ),
            )[0],
          )[0],
        );
      await editor.locator('select').first().selectOption('ja');
      await saved();
      await editor.getByRole('button', { name: '挿入', exact: true }).click();
      await editor.locator('button[title$="— addSlideChart"]').click();
      const dialog = editor.getByRole('dialog');
      await dialog.getByLabel('グラフの種類', { exact: true }).selectOption('line');
      await dialog.getByRole('button', { name: '項目を追加', exact: true }).click();
      await dialog.getByLabel('値 4, 1', { exact: true }).fill('25');
      await dialog.getByLabel('値 2, 1', { exact: true }).fill('');
      await dialog.getByLabel('空欄の値', { exact: true }).selectOption('span');
      await page.screenshot({ path: '/tmp/pptx-pr287-chart-blanks-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: 'グラフを挿入', exact: true }).click();
      await saved();
      assert.equal((await read()).dispBlanksAs, 'span');
      await editor.locator('select').first().selectOption('en');
      ja = false;
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      assert.equal(await dialog.getByLabel('Blank values', { exact: true }).inputValue(), 'span');
      await dialog.getByLabel('Blank values', { exact: true }).selectOption('zero');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal((await read()).dispBlanksAs, 'span');
      for (const kind of ['line', 'area']) {
        for (const blanks of ['gap', 'zero', 'span']) {
          await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
          await dialog.getByLabel('Chart type', { exact: true }).selectOption(kind);
          await dialog.getByLabel('Blank values', { exact: true }).selectOption(blanks);
          await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
          await saved();
          const spec = await read();
          assert.equal(spec.dispBlanksAs, blanks);
          assert.deepEqual(spec.series[0].values, [10, null, 15, 25]);
          const path = await editor
            .locator('.paint path[stroke="#4472C4"]')
            .first()
            .getAttribute('d');
          assert.equal(path.match(/M/g)?.length, blanks === 'gap' ? 2 : 1);
          assert.equal(path.match(/[ML]/g)?.length, blanks === 'zero' ? 4 : 3);
          if (kind === 'area') {
            const area = await editor
              .locator('.paint path[fill="#4472C4"]')
              .first()
              .getAttribute('d');
            assert.equal(area.match(/Z/g)?.length, blanks === 'gap' ? 2 : 1);
          }
        }
      }
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await read()).dispBlanksAs, 'zero');
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.equal((await read()).dispBlanksAs, 'span');
      await editor.locator('.hit').first().click();
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      assert.equal(await dialog.getByLabel('Blank values', { exact: true }).inputValue(), 'span');
      await dialog.getByLabel('Chart type', { exact: true }).selectOption('pie');
      assert.equal(await dialog.getByLabel('Blank values', { exact: true }).count(), 0);
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-chart-blanks-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
