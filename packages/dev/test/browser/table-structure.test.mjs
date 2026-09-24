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
  'table insertion and margins retain editing and persist through undo',
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
      const firstCell = await page.locator('#stage [data-pptx-cell="0,0"]').first().boundingBox();
      assert.ok(firstCell);
      await page.mouse.dblclick(
        firstCell.x + firstCell.width / 2,
        firstCell.y + firstCell.height / 2,
      );
      await page.screenshot({ path: '/tmp/pptx-table-structure-start.png' });
      let field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.fill('Original cell');
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
      const waitGrid = (rows, cols) =>
        page.waitForFunction(
          ({ rows, cols }) => {
            const state = window.office.getState();
            const table = state.editor.slides[0].shapes.find((s) => s.table)?.table;
            return (
              !state.building &&
              table?.rowHeights.length === rows &&
              table.columnWidths.length === cols
            );
          },
          { rows, cols },
        );
      await page.getByRole('button', { name: 'Insert Above', exact: true }).click();
      await waitGrid(3, 3);
      await field.waitFor();
      assert.equal(await field.textContent(), '');
      await field.fill('New row');
      await page.getByRole('button', { name: 'Insert Right', exact: true }).click();
      await waitGrid(3, 4);
      field = page.getByRole('textbox', { name: 'Edit cell 1, 2', exact: true });
      await field.waitFor();
      await field.fill('Margin text');
      await field.evaluate((el) => el.setSelectionRange(1, 4));
      await page.getByRole('button', { name: 'Cell Margins', exact: true }).click();
      let dialog = page.getByRole('dialog', { name: 'Cell Margins', exact: true });
      await dialog.getByRole('spinbutton', { name: 'Left margin' }).fill('0.2');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      await page.getByRole('button', { name: 'Cell Margins', exact: true }).click();
      dialog = page.getByRole('dialog', { name: 'Cell Margins', exact: true });
      assert.equal(
        await dialog.getByRole('spinbutton', { name: 'Left margin' }).inputValue(),
        '0.1',
      );
      await dialog.getByRole('spinbutton', { name: 'Left margin' }).fill('0.2');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(() => {
        const state = window.office.getState();
        return (
          !state.building &&
          state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][1].margins.left ===
            182880
        );
      });
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      await page.getByRole('button', { name: 'Insert Below', exact: true }).click();
      await waitGrid(4, 4);
      field = page.getByRole('textbox', { name: 'Edit cell 2, 2', exact: true });
      await field.waitFor();
      await page.getByRole('button', { name: 'Insert Left', exact: true }).click();
      await waitGrid(4, 5);
      await field.waitFor();
      await field.press('Escape');
      await page.locator('[data-edit=undo]').first().click();
      await waitGrid(4, 4);
      await page.locator('[data-edit=redo]').first().click();
      await waitGrid(4, 5);
      await page.screenshot({ path: '/tmp/pptx-table-structure.png' });
      await page.reload();
      await waitGrid(4, 5);
      const saved = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const table = pptx.getSlideTables(pptx.getSlides(saved)[0])[0];
      assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 0, 0)), 'New row');
      assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 0, 2)), 'Margin text');
      assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 2, 0)), 'Original cell');
      assert.equal(pptx.getTableCellMargins(pptx.getTableCell(table, 0, 2)).left, 182880);
      assert.equal(pptx.getTableCellMargins(pptx.getTableCell(table, 0, 0)).left, 91440);
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
