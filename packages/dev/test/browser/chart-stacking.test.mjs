import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeChartSpec, getSlideShapes, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'chart stacking supports bilingual editing, chart types and saved history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-chart-stacking-'));
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
      await dialog.getByRole('button', { name: '系列を追加', exact: true }).click();
      for (let i = 1; i <= 3; i++)
        await dialog.getByLabel(`値 ${i}, 2`, { exact: true }).fill('20');
      await dialog.getByLabel('系列の積み上げ', { exact: true }).selectOption('stacked');
      await page.screenshot({ path: '/tmp/pptx-pr287-chart-stacking-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: 'グラフを挿入', exact: true }).click();
      await saved();
      assert.equal((await read()).grouping, 'stacked');
      assert.equal((await read()).overlapPct, 100);
      await editor.locator('select').first().selectOption('en');
      ja = false;
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      assert.equal(
        await dialog.getByLabel('Series stacking', { exact: true }).inputValue(),
        'stacked',
      );
      await dialog.getByLabel('Series stacking', { exact: true }).selectOption('none');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal((await read()).grouping, 'stacked');
      for (const kind of ['column', 'bar', 'line', 'area']) {
        for (const stacking of ['none', 'stacked', 'percentStacked']) {
          await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
          await dialog.getByLabel('Chart type', { exact: true }).selectOption(kind);
          await dialog.getByLabel('Series stacking', { exact: true }).selectOption(stacking);
          await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
          await saved();
          const spec = await read();
          assert.equal(spec.kind, kind);
          assert.equal(
            spec.grouping,
            stacking === 'none'
              ? kind === 'column' || kind === 'bar'
                ? 'clustered'
                : 'standard'
              : stacking,
          );
          assert.deepEqual(
            spec.series.map((s) => s.values),
            [
              [10, 20, 15],
              [20, 20, 20],
            ],
          );
          if (kind === 'column' || kind === 'bar')
            assert.equal(spec.overlapPct ?? 0, stacking === 'none' ? 0 : 100);
          if (stacking === 'percentStacked')
            assert.match(await editor.locator('.paint').textContent(), /100%/);
        }
      }
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await read()).grouping, 'stacked');
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.equal((await read()).grouping, 'percentStacked');
      await editor.locator('.hit').first().click();
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      assert.equal(
        await dialog.getByLabel('Series stacking', { exact: true }).inputValue(),
        'percentStacked',
      );
      await dialog.getByLabel('Chart type', { exact: true }).selectOption('radar');
      assert.equal(await dialog.getByLabel('Series stacking', { exact: true }).count(), 0);
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({
        path: '/tmp/pptx-pr287-chart-stacking-failure.png',
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
