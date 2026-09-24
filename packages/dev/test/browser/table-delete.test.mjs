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
  'table deletion commits cell text, restores editing and supports undo',
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
      const field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      const deleteCommand = async (label) => {
        await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
        await page.locator('[data-edit=table-delete]').click();
        if (label === 'Delete Rows') await page.screenshot({ path: '/tmp/pptx-table-delete.png' });
        await page.getByRole('menuitem', { name: label, exact: true }).click();
      };
      const dimensions = (rows, columns) =>
        page.waitForFunction(
          ({ rows, columns }) => {
            const state = window.office.getState();
            const table = state.editor.slides[0].shapes.find((s) => s.table)?.table;
            return (
              !state.building &&
              table?.cells.length === rows &&
              table.columnWidths.length === columns
            );
          },
          { rows, columns },
        );
      await field.fill('Removed row');
      const families = await field.evaluate((element) =>
        [element, ...element.querySelectorAll('span')].map(
          (node) => getComputedStyle(node).fontFamily,
        ),
      );
      assert.ok(families.length > 1, 'typed text has a formatted run');
      assert.ok(families.every((family) => family.endsWith('Arial, sans-serif')));
      await deleteCommand('Delete Rows');
      await dimensions(1, 3);
      await field.waitFor();
      assert.equal(await field.textContent(), '');
      await field.fill('Removed column');
      await deleteCommand('Delete Columns');
      await dimensions(1, 2);
      await field.waitFor();
      assert.equal(await field.textContent(), '');
      await field.fill('Kept after Undo');
      await deleteCommand('Delete Table');
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          !window.office.getState().editor.slides[0].shapes.some((s) => s.table),
      );
      await page.locator('[data-edit=undo]').first().click();
      await dimensions(1, 2);
      const saved = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const table = pptx.getSlideTables(pptx.getSlides(saved)[0])[0];
      assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 0, 0)), 'Kept after Undo');
      await page.locator('[data-edit=redo]').first().click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          !window.office.getState().editor.slides[0].shapes.some((s) => s.table),
      );
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
