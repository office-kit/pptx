import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeChartSpec, getSlideShapes, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'chart axes support bilingual titles, bounds, intervals, automatic reset and history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-chart-axes-'));
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
      await dialog.getByText('グラフの軸', { exact: true }).click();
      await dialog.getByLabel('項目軸のタイトル', { exact: true }).fill('四半期');
      await dialog.getByLabel('数値軸のタイトル', { exact: true }).fill('売上高');
      await dialog.getByLabel('軸の最小値', { exact: true }).fill('-10');
      await dialog.getByLabel('軸の最大値', { exact: true }).fill('-20');
      assert.equal(
        await dialog.getByRole('button', { name: 'グラフを挿入', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByRole('alert').waitFor();
      await dialog.getByLabel('軸の最大値', { exact: true }).fill('50');
      await dialog.getByLabel('主目盛りの間隔', { exact: true }).fill('0');
      assert.equal(
        await dialog.getByRole('button', { name: 'グラフを挿入', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByLabel('主目盛りの間隔', { exact: true }).fill('10');
      await dialog.getByLabel('補助目盛りの間隔', { exact: true }).fill('2');
      await dialog.getByLabel('軸の数値書式', { exact: true }).fill('0.0');
      await dialog.getByLabel('主目盛線を表示', { exact: true }).check();
      await page.screenshot({ path: '/tmp/pptx-pr287-chart-axes-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: 'グラフを挿入', exact: true }).click();
      await saved();
      let spec = await read();
      assert.equal(spec.categoryAxisTitle, '四半期');
      assert.equal(spec.valueAxisTitle, '売上高');
      assert.equal(spec.valueAxis.min, -10);
      assert.equal(spec.valueAxis.max, 50);
      assert.equal(spec.valueAxis.majorUnit, 10);
      assert.equal(spec.valueAxis.minorUnit, 2);
      assert.equal(spec.valueAxis.numberFormat, '0.0');
      assert.equal(spec.valueAxisMajorGridlines, true);
      assert.ok((await editor.locator('.paint line[stroke-width="0.5"]').count()) > 0);
      assert.match(await editor.locator('.paint').textContent(), /四半期/);
      assert.match(await editor.locator('.paint').textContent(), /売上高/);
      await page.reload();
      await saved();
      await editor.locator('.hit').first().click();
      await editor.locator('select').first().selectOption('en');
      ja = false;
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByText('Chart axes', { exact: true }).click();
      assert.equal(await dialog.getByLabel('Axis minimum', { exact: true }).inputValue(), '-10');
      assert.equal(
        await dialog.getByLabel('Value axis title', { exact: true }).inputValue(),
        '売上高',
      );
      assert.equal(
        await dialog.getByLabel('Show major gridlines', { exact: true }).isChecked(),
        true,
      );
      assert.equal(
        await dialog.getByLabel('Show category axis', { exact: true }).isChecked(),
        true,
      );
      assert.equal(await dialog.getByLabel('Show value axis', { exact: true }).isChecked(), true);
      await dialog.getByLabel('Show value axis', { exact: true }).uncheck();
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal((await read()).valueAxisHidden ?? false, false);
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByText('Chart axes', { exact: true }).click();
      for (const label of ['Show category axis', 'Show value axis', 'Show major gridlines'])
        await dialog.getByLabel(label, { exact: true }).uncheck();
      for (const label of [
        'Axis minimum',
        'Axis maximum',
        'Major tick interval',
        'Minor tick interval',
        'Axis number format',
        'Category axis title',
        'Value axis title',
      ])
        await dialog.getByLabel(label, { exact: true }).fill('');
      await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
      await saved();
      spec = await read();
      for (const key of ['min', 'max', 'majorUnit', 'minorUnit'])
        assert.equal(spec.valueAxis?.[key], undefined);
      assert.equal(spec.categoryAxisTitle, undefined);
      assert.equal(spec.valueAxisTitle, undefined);
      assert.equal(spec.categoryAxisHidden, true);
      assert.equal(spec.valueAxisHidden, true);
      assert.equal(spec.valueAxisMajorGridlines, false);
      assert.equal(await editor.locator('.paint line[stroke-width="0.5"]').count(), 0);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await read()).valueAxis.min, -10);
      assert.equal((await read()).valueAxisHidden ?? false, false);
      assert.equal((await read()).valueAxisMajorGridlines, true);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.equal((await read()).valueAxis?.min, undefined);
      assert.equal((await read()).valueAxisHidden, true);
      assert.equal((await read()).categoryAxisHidden, true);
      assert.equal((await read()).valueAxisMajorGridlines, false);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-chart-axes-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
