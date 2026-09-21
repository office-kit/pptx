import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeChartSpec, getSlideShapes, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'line chart markers and smoothing support bilingual editing and saved history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-chart-markers-'));
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
      await dialog.getByText('線とマーカー', { exact: true }).click();
      await dialog.getByLabel('マーカーの形 1', { exact: true }).selectOption('circle');
      await dialog.getByLabel('マーカーのサイズ (pt) 1', { exact: true }).fill('1');
      assert.equal(
        await dialog.getByRole('button', { name: 'グラフを挿入', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByLabel('マーカーのサイズ (pt) 1', { exact: true }).fill('12');
      await dialog.getByLabel('滑らかな線 1', { exact: true }).check();
      await dialog.getByLabel('系列の積み上げ', { exact: true }).selectOption('stacked');
      await page.screenshot({ path: '/tmp/pptx-pr287-chart-markers-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: 'グラフを挿入', exact: true }).click();
      await saved();
      const first = (await read()).series[0];
      assert.equal(first.markerSymbol, 'circle');
      assert.equal(first.markerSizePt, 12);
      assert.equal(first.smooth, true);
      assert.equal(await editor.locator('.paint circle[r="6.00"]').count(), 3);
      assert.match(
        await editor.locator('.paint path[stroke="#4472C4"]').first().getAttribute('d'),
        /C/,
      );
      await editor.getByRole('button', { name: 'グラフを編集', exact: true }).click();
      await dialog.getByRole('button', { name: '項目を追加', exact: true }).click();
      await dialog.getByLabel('値 4, 1', { exact: true }).fill('25');
      await dialog.getByLabel('値 2, 1', { exact: true }).fill('');
      await dialog.getByRole('button', { name: '変更を適用', exact: true }).click();
      await saved();
      assert.equal((await read()).series[0].values[1], null);
      const gapPath = await editor
        .locator('.paint path[stroke="#4472C4"]')
        .first()
        .getAttribute('d');
      assert.equal(gapPath.match(/M/g)?.length, 2);
      assert.doesNotMatch(gapPath, /C/);
      assert.equal(await editor.locator('.paint circle[r="6.00"]').count(), 3);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await read()).series[0].values[1], 20);
      await editor.locator('select').first().selectOption('en');
      ja = false;
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByText('Line and markers', { exact: true }).click();
      assert.equal(
        await dialog.getByLabel('Marker shape 1', { exact: true }).inputValue(),
        'circle',
      );
      assert.equal(
        await dialog.getByLabel('Marker size (pt) 1', { exact: true }).inputValue(),
        '12',
      );
      assert.equal(await dialog.getByLabel('Smooth line 1', { exact: true }).isChecked(), true);
      await dialog.getByLabel('Marker shape 1', { exact: true }).selectOption('none');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal((await read()).series[0].markerSymbol, 'circle');
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByText('Line and markers', { exact: true }).click();
      await dialog.getByLabel('Marker shape 1', { exact: true }).selectOption('none');
      await dialog.getByLabel('Marker size (pt) 1', { exact: true }).fill('');
      await dialog.getByLabel('Smooth line 1', { exact: true }).uncheck();
      await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
      await saved();
      assert.equal((await read()).series[0].markerSymbol, 'none');
      assert.equal((await read()).series[0].markerSizePt, undefined);
      assert.equal((await read()).series[0].smooth, false);
      assert.equal(await editor.locator('.paint circle').count(), 0);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await read()).series[0].markerSizePt, 12);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.equal((await read()).series[0].markerSymbol, 'none');
      assert.equal((await read()).series[0].smooth, false);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-chart-markers-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
