import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeChartSpec, getSlideShapes, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'pie and doughnut options support bilingual editing, rendering and saved history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-chart-pie-options-'));
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
      await editor.getByRole('tab', { name: '挿入', exact: true }).click();
      await editor.locator('button[title$="— addSlideChart"]').click();
      const dialog = editor.getByRole('dialog');
      await dialog.getByLabel('グラフの種類', { exact: true }).selectOption('doughnut');
      await dialog.getByLabel('最初の扇形の角度 (°)', { exact: true }).fill('90');
      await dialog.getByLabel('ドーナツの穴の大きさ (%)', { exact: true }).fill('9');
      assert.equal(
        await dialog.getByRole('button', { name: 'グラフを挿入', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByLabel('ドーナツの穴の大きさ (%)', { exact: true }).fill('70');
      await dialog.getByRole('button', { name: '扇形を色分け', exact: true }).click();
      await dialog.getByLabel('扇形の色 2', { exact: true }).fill('#cc2299');
      await page.screenshot({ path: '/tmp/pptx-pr287-chart-pie-options-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: 'グラフを挿入', exact: true }).click();
      await saved();
      assert.equal((await read()).firstSliceAngleDeg, 90);
      assert.equal((await read()).holeSizePct, 70);
      const arcs = async () =>
        editor
          .locator('.paint path[stroke="#FFFFFF"][stroke-width="0.6"]')
          .first()
          .getAttribute('d');
      assert.deepEqual((await read()).series[0].pointColors, ['#4472C4', '#CC2299', '#A5A5A5']);
      assert.equal(
        await editor.locator('.paint path[fill="#CC2299"][stroke="#FFFFFF"]').count(),
        1,
      );
      const initialPath = await arcs();
      const radii = [...initialPath.matchAll(/A([\d.]+),/g)].map((match) => Number(match[1]));
      assert.equal(radii.length, 2);
      assert.ok(Math.abs(radii[1] / radii[0] - 0.7) < 0.001);
      await editor.locator('select').first().selectOption('en');
      ja = false;
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      assert.equal(
        await dialog.getByLabel('First slice angle (°)', { exact: true }).inputValue(),
        '90',
      );
      assert.equal(
        await dialog.getByLabel('Doughnut hole size (%)', { exact: true }).inputValue(),
        '70',
      );
      assert.equal(
        await dialog.getByLabel('Slice color 2', { exact: true }).inputValue(),
        '#cc2299',
      );
      await dialog.getByLabel('First slice angle (°)', { exact: true }).fill('180');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal((await read()).firstSliceAngleDeg, 90);
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByLabel('First slice angle (°)', { exact: true }).fill('361');
      assert.equal(
        await dialog.getByRole('button', { name: 'Apply changes', exact: true }).isDisabled(),
        true,
      );
      await dialog.getByLabel('First slice angle (°)', { exact: true }).fill('');
      await dialog.getByLabel('Doughnut hole size (%)', { exact: true }).fill('');
      await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
      await saved();
      assert.equal((await read()).firstSliceAngleDeg ?? 0, 0);
      assert.equal((await read()).holeSizePct, 50);
      assert.notEqual(await arcs(), initialPath);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await read()).firstSliceAngleDeg, 90);
      assert.equal((await read()).holeSizePct, 70);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.equal((await read()).holeSizePct, 50);
      await editor.locator('.hit').first().click();
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByLabel('Chart type', { exact: true }).selectOption('pie');
      assert.equal(await dialog.getByLabel('Doughnut hole size (%)', { exact: true }).count(), 0);
      await dialog.getByLabel('First slice angle (°)', { exact: true }).fill('180');
      await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
      await saved();
      assert.equal((await read()).kind, 'pie');
      assert.equal((await read()).firstSliceAngleDeg, 180);
      assert.equal([...(await arcs()).matchAll(/A/g)].length, 1);
      await editor.getByRole('button', { name: 'Edit chart', exact: true }).click();
      await dialog.getByRole('button', { name: 'Use series color 2', exact: true }).click();
      assert.equal(
        await dialog.getByLabel('Slice color 2', { exact: true }).inputValue(),
        '#4472c4',
      );
      await dialog.getByRole('button', { name: 'Remove category 1', exact: true }).click();
      assert.equal(
        await dialog.getByLabel('Slice color 2', { exact: true }).inputValue(),
        '#a5a5a5',
      );
      await dialog.getByRole('button', { name: 'Apply changes', exact: true }).click();
      await saved();
      assert.deepEqual(
        Array.from((await read()).series[0].pointColors, (color) => color ?? null),
        [null, '#A5A5A5'],
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual((await read()).series[0].pointColors, ['#4472C4', '#CC2299', '#A5A5A5']);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.deepEqual(
        Array.from((await read()).series[0].pointColors, (color) => color ?? null),
        [null, '#A5A5A5'],
      );
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({
        path: '/tmp/pptx-pr287-chart-pie-options-failure.png',
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
