import * as pptx from '@office-kit/pptx';
import { measureTableCellHeight } from '@office-kit/pptx-preview';
import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'Table resize enforces text height, anchors the opposite edge and undoes',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={5} height={1}>Alpha [a] ALPHA</Text></Slide><Slide><Text x={1} y={1} width={5} height={1}>Beta alpha</Text></Slide></Presentation>`,
    );
    const proc = spawn(process.execPath, [
      fileURLToPath(new URL('../../dist/cli.mjs', import.meta.url)),
      'dev',
      file,
      '--port',
      '0',
    ]);
    let browser;
    try {
      const url = await new Promise((resolve, reject) => {
        let output = '';
        proc.stdout.on('data', (data) => {
          output += data;
          const match = output.match(/Preview: (http:\/\/\S+)/);
          if (match) resolve(match[1]);
        });
        proc.stderr.on('data', (data) => process.stderr.write(data));
        proc.on('error', reject);
        proc.on('exit', (code) => reject(new Error('Server exited: ' + code)));
      });
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
      await page.goto(url);
      await page.locator('.shape-hit').first().waitFor();
      await page.getByRole('tab', { name: 'Insert', exact: true }).click();
      await page.locator('[data-edit=table]').click();
      await page.getByRole('gridcell', { name: '3 columns, 2 rows', exact: true }).click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes.some((s) => s.table),
      );
      await page.locator('.shape-hit.selected').waitFor();
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).waitFor();
      const first = await page.locator('#stage [data-pptx-cell="0,0"]').first().boundingBox();
      await page.mouse.click(first.x + first.width / 2, first.y + first.height / 2);
      const field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.waitFor();
      await field.fill('First line\nSecond line\nThird line');
      await field.press('Escape');
      await page.waitForFunction(() => {
        const state = window.office.getState();
        return (
          !state.building &&
          state.editor.slides[0].shapes.find((s) => s.table)?.table.cells[0][0].text ===
            'First line\nSecond line\nThird line'
        );
      });
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
      const height = page.getByRole('spinbutton', { name: 'Table Row Height', exact: true });
      await height.fill('1.5');
      await height.press('Enter');
      await page.waitForFunction(() => {
        const state = window.office.getState();
        return (
          !state.building &&
          state.editor.slides[0].shapes
            .find((s) => s.table)
            .table.rowHeights.every((h) => h === 1371600)
        );
      });
      await height.click({ trial: true });
      const state = () => page.evaluate(() => window.office.getState());
      const before = await state();
      const original = before.editor.slides[0].shapes.find((s) => s.table).bounds;
      const handle = await page.locator('.shape-hit.selected [data-handle="n"]').boundingBox();
      assert.ok(handle);
      await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2);
      await page.mouse.down();
      await page.mouse.move(handle.x + handle.width / 2, handle.y + 120, { steps: 5 });
      await page.mouse.up();
      await page.waitForFunction((revision) => {
        const state = window.office.getState();
        return !state.building && state.revision !== revision;
      }, before.revision);
      const resized = (await state()).editor.slides[0].shapes.find((s) => s.table);
      assert.ok(resized.bounds.h > 0);
      assert.ok(Math.abs(resized.bounds.y + resized.bounds.h - original.y - original.h) <= 1);
      const downloaded = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const table = pptx.getSlideTables(pptx.getSlides(downloaded)[0])[0];
      assert.ok(
        pptx.getTableRowHeights(table)[0] >= measureTableCellHeight(downloaded, table, 0, 0),
      );
      assert.deepEqual(pptx.getShapeBoundsResolved(downloaded, table), resized.bounds);
      await page.screenshot({ path: '/tmp/pptx-table-resize.png' });
      await page.locator('[data-edit=undo]').first().click();
      await page.waitForFunction((expected) => {
        const state = window.office.getState();
        const bounds = state.editor.slides[0].shapes.find((s) => s.table)?.bounds;
        return !state.building && JSON.stringify(bounds) === JSON.stringify(expected);
      }, original);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await new Promise((resolve) => {
        if (proc.exitCode !== null) resolve();
        else proc.once('exit', resolve);
      });
      await rm(dir, { recursive: true, force: true });
    }
  },
);
