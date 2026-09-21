import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeChartSpec, getSlideShapes, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'data label formats support bilingual editing, rendering and saved history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-chart-label-format-'));
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
      await dialog.getByLabel('グラフの種類', { exact: true }).selectOption('pie');
      await dialog.getByLabel('値を表示', { exact: true }).check();
      await dialog.getByText('データラベルの書式', { exact: true }).click();
      await dialog.getByLabel('ラベルの位置', { exact: true }).selectOption('outEnd');
      await dialog.getByLabel('ラベルの表示形式', { exact: true }).fill('0.00');
      await page.screenshot({ path: '/tmp/pptx-pr287-chart-label-format-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: 'グラフを挿入', exact: true }).click();
      await saved();
      assert.equal((await read()).dataLabels.position, 'outEnd');
      assert.equal((await read()).dataLabels.numberFormat, '0.00');
      const label = editor.locator('.paint text').filter({ hasText: /^10\.00$/ });
      assert.equal(await label.count(), 1);
      const outsideX = await label.getAttribute('x');
      await editor.locator('select').first().selectOption('en');
      ja = false;
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByText('Data label format', { exact: true }).click();
      assert.equal(
        await dialog.getByLabel('Label position', { exact: true }).inputValue(),
        'outEnd',
      );
      assert.equal(
        await dialog.getByLabel('Label number format', { exact: true }).inputValue(),
        '0.00',
      );
      await dialog.getByLabel('Label position', { exact: true }).selectOption('ctr');
      await dialog.getByLabel('Label number format', { exact: true }).fill('0.0');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal((await read()).dataLabels.position, 'outEnd');
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByText('Data label format', { exact: true }).click();
      await dialog.getByLabel('Label position', { exact: true }).selectOption('ctr');
      await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
      await saved();
      assert.notEqual(await label.getAttribute('x'), outsideX);
      assert.equal((await read()).dataLabels.numberFormat, '0.00');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await read()).dataLabels.position, 'outEnd');
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.equal((await read()).dataLabels.position, 'ctr');
      await editor.locator('.hit').first().click();
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByText('Data label format', { exact: true }).click();
      await dialog.getByLabel('Label position', { exact: true }).selectOption('');
      await dialog.getByLabel('Label number format', { exact: true }).fill('');
      await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
      await saved();
      assert.equal((await read()).dataLabels.position, undefined);
      assert.equal((await read()).dataLabels.numberFormat, undefined);
      assert.equal(await editor.locator('.paint text').filter({ hasText: /^10$/ }).count(), 1);
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByText('Data label format', { exact: true }).click();
      for (const [kind, positions] of [
        ['column', ['', 'ctr', 'inEnd', 'outEnd', 'inBase']],
        ['bar', ['', 'ctr', 'inEnd', 'outEnd', 'inBase']],
        ['line', ['', 'ctr', 't', 'b', 'l', 'r']],
        ['area', ['', 'ctr', 't', 'b', 'l', 'r']],
        ['doughnut', ['', 'ctr', 'inEnd', 'outEnd', 'bestFit']],
      ]) {
        await dialog.getByLabel('Chart type', { exact: true }).selectOption(kind);
        assert.deepEqual(
          await dialog
            .getByLabel('Label position', { exact: true })
            .locator('option')
            .evaluateAll((options) => options.map((option) => option.value)),
          positions,
        );
      }
      await dialog.getByLabel('Chart type', { exact: true }).selectOption('radar');
      assert.equal(await dialog.getByLabel('Label position', { exact: true }).count(), 0);
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal((await read()).kind, 'pie');
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByLabel('Show values', { exact: true }).uncheck();
      await dialog.getByLabel('Show series names', { exact: true }).check();
      await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
      await saved();
      const seriesName = (await read()).series[0].name;
      assert.equal(await editor.locator('.paint text').filter({ hasText: seriesName }).count(), 3);
      for (const kind of ['column', 'bar', 'line', 'area']) {
        await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
        await dialog.getByLabel('Chart type', { exact: true }).selectOption(kind);
        await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
        await saved();
        assert.equal((await read()).kind, kind);
        assert.equal(
          await editor.locator('.paint text').filter({ hasText: seriesName }).count(),
          3,
        );
      }
      for (const kind of ['column', 'bar']) {
        const axis = kind === 'column' ? 'y' : 'x';
        for (const grouping of ['stacked', 'percentStacked']) {
          let center;
          for (const position of ['ctr', 'inEnd']) {
            await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
            await dialog.getByLabel('Chart type', { exact: true }).selectOption(kind);
            await dialog.getByLabel('Series stacking', { exact: true }).selectOption(grouping);
            await dialog.getByText('Data label format', { exact: true }).click();
            await dialog.getByLabel('Label position', { exact: true }).selectOption(position);
            await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
            await saved();
            assert.equal((await read()).dataLabels.position, position);
            const coordinate = await editor
              .locator('.paint text')
              .filter({ hasText: seriesName })
              .first()
              .getAttribute(axis);
            if (position === 'ctr') center = coordinate;
            else assert.notEqual(coordinate, center);
          }
        }
      }
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({
        path: '/tmp/pptx-pr287-chart-label-format-failure.png',
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
