import * as pptx from '@office-kit/pptx';
import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test('cell range dragging merges cells and supports undo', { timeout: 120000 }, async () => {
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
    const last = await page.locator('#stage [data-pptx-cell="1,1"]').first().boundingBox();
    const before = await page.evaluate(
      () => window.office.getState().editor.slides[0].shapes.find((s) => s.table).bounds,
    );
    await page.mouse.move(first.x + first.width / 2, first.y + first.height / 2);
    await page.mouse.down();
    await page.mouse.move(last.x + last.width / 2, last.y + last.height / 2, { steps: 8 });
    await page.mouse.up();
    await page.locator('.table-cell-selection').waitFor();
    assert.equal(
      await page.locator('.table-cell-selection').getAttribute('aria-label'),
      '2 rows, 2 columns selected',
    );
    const gridBefore = await page.evaluate(
      () =>
        window.office.getState().editor.slides[0].shapes.find((s) => s.table).table.columnWidths,
    );
    await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
    await page.locator('#table-size-columns').fill('1.5');
    await page.locator('#table-size-columns').press('Enter');
    await page.waitForFunction(() => {
      const state = window.office.getState();
      const widths = state.editor.slides[0].shapes.find((s) => s.table).table.columnWidths;
      return !state.building && widths[0] === 1371600 && widths[1] === 1371600;
    });
    await page.locator('.table-cell-selection').waitFor();
    assert.equal(await page.locator('#table-size-columns').inputValue(), '1.5');
    assert.equal(
      await page.evaluate(
        () =>
          window.office.getState().editor.slides[0].shapes.find((s) => s.table).table
            .columnWidths[2],
      ),
      gridBefore[2],
    );
    await page.waitForFunction(() => !document.querySelector('[data-edit=undo]').disabled);
    await page.locator('[data-edit=undo]').first().click();
    await page.waitForFunction(() => !document.querySelector('[data-edit=table-delete]').disabled);
    await page.mouse.move(first.x + first.width / 2, first.y + first.height / 2);
    await page.mouse.down();
    await page.mouse.move(last.x + last.width / 2, last.y + last.height / 2, { steps: 8 });
    await page.mouse.up();
    await page.locator('.table-cell-selection').waitFor();
    await page.getByRole('tab', { name: 'Home', exact: true }).click();
    await page.locator('[data-edit=bold]').first().click();
    await page.locator('.table-cell-selection').waitFor();
    await page.waitForFunction(() => !document.querySelector('[data-edit=bold]').disabled);
    await page.getByRole('tab', { name: 'Table Design', exact: true }).click();
    await page.locator('[data-edit=table-shading]').click();
    await page.getByRole('menuitem', { name: 'Red', exact: true }).click();
    await page.locator('#stage [data-pptx-cell="1,1"][fill="#FF0000"]').waitFor();
    await page.locator('.table-cell-selection').waitFor();
    const formatted = await pptx.loadPresentation(
      await (await page.request.get(url + '/deck.pptx')).body(),
    );
    const formattedTable = pptx.getSlideTables(pptx.getSlides(formatted)[0])[0];
    for (const row of [0, 1])
      for (const column of [0, 1, 2]) {
        const cell = pptx.getTableCell(formattedTable, row, column);
        const paragraphs = pptx.getTableCellParagraphs(cell);
        assert.equal(paragraphs[0].endFormat?.bold ?? false, column < 2);
        assert.equal(pptx.getTableCellFill(cell) === '#FF0000', column < 2);
      }
    await page.waitForFunction(() => !document.querySelector('[data-edit=table-borders]').disabled);
    await page.locator('[data-edit=table-borders]').click();
    await page.getByRole('menuitem', { name: 'Outside Borders', exact: true }).click();
    await page.locator('.table-cell-selection').waitFor();
    await page.waitForFunction(() => !document.querySelector('[data-edit=table-borders]').disabled);
    const bordered = await pptx.loadPresentation(
      await (await page.request.get(url + '/deck.pptx')).body(),
    );
    const borderedTable = pptx.getSlideTables(pptx.getSlides(bordered)[0])[0];
    const selectedRight = pptx.getTableCellBorders(
      bordered,
      pptx.getTableCell(borderedTable, 0, 1),
    ).right;
    const neighborLeft = pptx.getTableCellBorders(
      bordered,
      pptx.getTableCell(borderedTable, 0, 2),
    ).left;
    assert.ok(selectedRight);
    assert.deepEqual(selectedRight, neighborLeft);
    assert.notDeepEqual(
      selectedRight,
      pptx.getTableCellBorders(formatted, pptx.getTableCell(formattedTable, 0, 1)).right,
    );
    for (let i = 0; i < 3; i++) {
      await page.waitForFunction(() => !document.querySelector('[data-edit=undo]').disabled);
      await page.locator('[data-edit=undo]').first().click();
      await page.waitForFunction(() => !document.querySelector('[data-edit=undo]').disabled);
    }
    await page.mouse.move(first.x + first.width / 2, first.y + first.height / 2);
    await page.mouse.down();
    await page.mouse.move(last.x + last.width / 2, last.y + last.height / 2, { steps: 8 });
    await page.mouse.up();
    await page.locator('.table-cell-selection').waitFor();
    await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();
    await page.locator('[data-edit=table-merge]').click();
    const field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
    await field.waitFor();
    await field.fill('Merged 日本語');
    await field.press('Escape');
    await page.waitForFunction(() => {
      const s = window.office.getState();
      return (
        !s.building &&
        s.editor.slides[0].shapes.find((s) => s.table)?.table.cells[0][0].text === 'Merged 日本語'
      );
    });
    const saved = await pptx.loadPresentation(
      await (await page.request.get(url + '/deck.pptx')).body(),
    );
    const table = pptx.getSlideTables(pptx.getSlides(saved)[0])[0];
    const span = pptx.getTableCellSpan(pptx.getTableCell(table, 0, 0));
    assert.equal(span.rowSpan, 2);
    assert.equal(span.gridSpan, 2);
    assert.equal(pptx.getTableCellText(pptx.getTableCell(table, 0, 0)), 'Merged 日本語');
    assert.deepEqual(
      await page.evaluate(
        () => window.office.getState().editor.slides[0].shapes.find((s) => s.table).bounds,
      ),
      before,
    );
    await page.screenshot({ path: '/tmp/pptx-table-merge.png' });
    await page.waitForFunction(() => !document.querySelector('[data-edit=table-delete]').disabled);
    const merged = await page.locator('#stage [data-pptx-cell="0,0"]').first().boundingBox();
    await page.mouse.move(merged.x + merged.width / 2, merged.y + merged.height / 2);
    await page.mouse.down();
    await page.mouse.move(merged.x + merged.width / 2 + 10, merged.y + merged.height / 2, {
      steps: 3,
    });
    await page.mouse.up();
    await page.locator('.table-cell-selection').waitFor();
    await page.keyboard.press('Backspace');
    await page.waitForFunction(() => {
      const state = window.office.getState();
      const table = state.editor.slides[0].shapes.find((s) => s.table)?.table;
      return (
        !state.building && table?.cells[0][0].text === '' && table.cells[0][0].span.gridSpan === 2
      );
    });
    await page.locator('.table-cell-selection').waitFor();
    await page.locator('[data-edit=undo]').first().click();
    await page.waitForFunction(
      () =>
        !window.office.getState().building &&
        window.office.getState().editor.slides[0].shapes.find((s) => s.table)?.table.cells[0][0]
          .text === 'Merged 日本語',
    );
    await page.locator('[data-edit=undo]').first().click();
    await page.waitForFunction(
      () =>
        !window.office.getState().building &&
        window.office.getState().editor.slides[0].shapes.find((s) => s.table)?.table.cells[0][0]
          .text === '',
    );
    await page.locator('[data-edit=undo]').first().click();
    await page.waitForFunction(
      () =>
        !window.office.getState().building &&
        window.office.getState().editor.slides[0].shapes.find((s) => s.table)?.table.cells[0][0]
          .span.gridSpan === 1,
    );
    assert.deepEqual(errors, []);
    // A range deletion removes only the selected columns, not the whole table.
    await page.waitForFunction(() => !document.querySelector('[data-edit=table-delete]').disabled);
    const a = await page.locator('#stage [data-pptx-cell="0,0"]').first().boundingBox();
    const b = await page.locator('#stage [data-pptx-cell="1,1"]').first().boundingBox();
    await page.mouse.move(a.x + a.width / 2, a.y + a.height / 2);
    await page.mouse.down();
    await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2, { steps: 8 });
    await page.mouse.up();
    await page.locator('.table-cell-selection').waitFor();
    assert.equal(
      await page.locator('.table-cell-selection').getAttribute('aria-label'),
      '2 rows, 2 columns selected',
    );
    await page.locator('[data-edit=table-delete]').click();
    await page.getByRole('menuitem', { name: 'Delete Columns', exact: true }).click();
    await page.waitForFunction(() => {
      const state = window.office.getState();
      const table = state.editor.slides[0].shapes.find((s) => s.table)?.table;
      return !state.building && table?.columnWidths.length === 1 && table.cells.length === 2;
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
});
