import * as pptx from '@office-kit/pptx';
import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'Split Cells dialog preserves content, cancels, exports and undoes',
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
      await field.fill('Split 日本語');
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
      await page.getByRole('button', { name: 'Split Cells', exact: true }).click();
      let dialog = page.getByRole('dialog', { name: 'Split Cells' });
      await dialog.getByLabel('Number of columns:').fill('3');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(await dialog.count(), 0);
      await field.click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Split Cells...', exact: true }).click();
      dialog = page.getByRole('dialog', { name: 'Split Cells' });
      await dialog.getByLabel('Number of columns:').fill('3');
      await dialog.getByLabel('Number of rows:').fill('2');
      await dialog.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.waitForFunction(() => {
        const state = window.office.getState();
        const table = state.editor.slides[0].shapes.find((s) => s.table)?.table;
        return (
          !state.building && table?.columnWidths.length === 5 && table?.rowHeights.length === 3
        );
      });
      await field.waitFor();
      assert.equal(await field.innerText(), 'Split 日本語');
      const downloaded = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const table = pptx.getSlideTables(pptx.getSlides(downloaded)[0])[0];
      assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 0, 0)), 'Split 日本語');
      assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 1, 2)), '');
      assert.equal(pptx.getTableCellSpan(pptx.getTableCell(table, 0, 3)).rowSpan, 2);
      await page.screenshot({ path: '/tmp/pptx-table-split.png' });
      await field.press('Escape');
      await page.locator('[data-edit=undo]').first().click();
      await page.waitForFunction(() => {
        const state = window.office.getState();
        const table = state.editor.slides[0].shapes.find((s) => s.table)?.table;
        return (
          !state.building && table?.columnWidths.length === 3 && table?.rowHeights.length === 2
        );
      });
      await page.locator('[data-edit=table-insert-below]').click({ trial: true });
      const restored = await page.locator('#stage [data-pptx-cell="0,0"]').first().boundingBox();
      await page.mouse.click(restored.x + restored.width / 2, restored.y + restored.height / 2, {
        button: 'right',
      });
      await field.waitFor();
      await page.getByRole('menuitem', { name: 'Insert', exact: true }).hover();
      await page.getByRole('menuitem', { name: 'Insert Rows Below', exact: true }).click();
      await page.waitForFunction(() => {
        const state = window.office.getState();
        return (
          !state.building &&
          state.editor.slides[0].shapes.find((s) => s.table)?.table.rowHeights.length === 3
        );
      });
      const inserted = page.getByRole('textbox', { name: 'Edit cell 2, 1', exact: true });
      await inserted.waitFor();
      await page.getByRole('button', { name: 'Split Cells', exact: true }).click({ trial: true });
      await inserted.click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Delete', exact: true }).hover();
      await page.getByRole('menuitem', { name: 'Delete Rows', exact: true }).click();
      await page.waitForFunction(() => {
        const state = window.office.getState();
        return (
          !state.building &&
          state.editor.slides[0].shapes.find((s) => s.table)?.table.rowHeights.length === 2
        );
      });
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
