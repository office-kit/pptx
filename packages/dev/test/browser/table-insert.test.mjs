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
  'table picker supports keyboard, dialog, undo and saved table dimensions',
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
      const count = (n) =>
        page.waitForFunction(
          (n) =>
            window.office.getState().editor.slides[0].shapes.length === n &&
            !window.office.getState().building,
          n,
        );
      const checkTable = async (rows, columns) => {
        const p = await pptx.loadPresentation(
          await (await page.request.get(url + '/deck.pptx')).body(),
        );
        const table = pptx.getSlideTables(pptx.getSlides(p)[0])[0];
        assert.equal(pptx.getTableRowHeights(table).length, rows);
        assert.equal(pptx.getTableColumnWidths(table).length, columns);
      };
      assert.equal(await page.locator('#tab-table-layout').isVisible(), false);
      await page.getByRole('tab', { name: 'View', exact: true }).focus();
      await page.keyboard.press('ArrowRight');
      assert.equal(
        await page.locator('#tab-home').evaluate((el) => el === document.activeElement),
        true,
      );
      await page.getByRole('tab', { name: 'Insert', exact: true }).click();
      const button = page.locator('[data-edit=table]');
      await button.click();
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('ArrowDown');
      assert.equal(
        await page
          .getByRole('gridcell', { name: '3 columns, 2 rows', exact: true })
          .evaluate((el) => el === document.activeElement),
        true,
      );
      assert.equal(await page.getByRole('gridcell', { selected: true }).count(), 6);
      await page.screenshot({ path: '/tmp/pptx-table-picker.png' });
      await page.keyboard.press('Escape');
      assert.equal(await button.evaluate((el) => el === document.activeElement), true);
      await count(1);
      await button.click();
      await page.getByRole('gridcell', { name: '3 columns, 2 rows', exact: true }).click();
      await count(2);
      await checkTable(2, 3);
      await page.locator('.shape-hit.selected').waitFor();
      assert.equal(await page.locator('.shape-hit.selected').count(), 1);
      const tableLayout = page.getByRole('tab', { name: 'Table Layout', exact: true });
      assert.equal(await tableLayout.isVisible(), true);
      await tableLayout.click();
      assert.equal(await page.locator('#ribbon-table-layout').isVisible(), true);
      await page.locator('[data-edit=undo]').first().click();
      await count(1);
      assert.equal(await page.locator('#tab-table-layout').isVisible(), false);
      assert.equal(await page.locator('#tab-home').getAttribute('aria-selected'), 'true');
      await page.getByRole('tab', { name: 'Insert', exact: true }).click();

      await page.locator('[data-edit=redo]').first().click();
      await count(2);
      await checkTable(2, 3);
      await page.locator('[data-edit=undo]').first().click();
      await count(1);
      await button.click();
      await page.getByRole('button', { name: 'Insert Table…', exact: true }).click();
      const dialog = page.getByRole('dialog', { name: 'Insert Table', exact: true });
      await dialog.getByLabel('Number of columns:').fill('4');
      await dialog.getByLabel('Number of rows:').fill('0');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      assert.equal(await dialog.isVisible(), true);
      await count(1);
      await dialog.getByLabel('Number of rows:').fill('9');
      await page.screenshot({ path: '/tmp/pptx-insert-table-dialog.png' });
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await count(2);
      await dialog.waitFor({ state: 'detached' });
      await checkTable(9, 4);
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      await count(2);
      await checkTable(9, 4);
      await page.screenshot({ path: '/tmp/pptx-inserted-table.png' });
      const hit = page.locator('.shape-hit').last();
      await hit.dblclick({ position: { x: 30, y: 15 } });
      let field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.waitFor();
      const headerFill = await page
        .locator('[data-pptx-cell="0,0"]')
        .evaluate((cell) => getComputedStyle(cell).fill);
      assert.notEqual(headerFill, 'rgb(255, 255, 255)');
      assert.equal(await field.evaluate((el) => getComputedStyle(el).backgroundColor), headerFill);
      await field.fill('日本語 😀');
      await field.press('Tab');
      field = page.getByRole('textbox', { name: 'Edit cell 1, 2', exact: true });
      await field.fill('Second');
      await field.press('Shift+Tab');
      field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.waitFor();
      assert.equal(await field.textContent(), '日本語 😀');
      await page.screenshot({ path: '/tmp/pptx-table-cell-editing.png' });
      await field.press('Escape');
      const waitText = (expected) =>
        page.waitForFunction((expected) => {
          const table = window.office.getState().editor.slides[0].shapes.find((s) => s.table).table;
          return table.cells[0][1].text === expected && !window.office.getState().building;
        }, expected);
      await waitText('Second');
      await page.locator('[data-edit=undo]').first().click();
      await waitText('');
      await page.locator('[data-edit=redo]').first().click();
      await waitText('Second');
      await page.reload();
      await page.locator('.shape-hit').last().waitFor();
      await waitText('Second');
      const saved = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const savedTable = pptx.getSlideTables(pptx.getSlides(saved)[0])[0];
      assert.equal(pptx.getTableCellText(pptx.getTableCell(savedTable, 0, 0)), '日本語 😀');
      assert.equal(pptx.getTableCellText(pptx.getTableCell(savedTable, 0, 1)), 'Second');
      await page.screenshot({ path: '/tmp/pptx-table-cell-edited.png' });
      await page
        .locator('.shape-hit')
        .last()
        .dblclick({ position: { x: 30, y: 15 } });
      field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.evaluate((el) => el.setSelectionRange(0, 3));
      await page.locator('[data-edit=bold]').first().click();
      await page.waitForFunction(() => {
        const table = window.office.getState().editor.slides[0].shapes.find((s) => s.table).table;
        return (
          table.cells[0][0].paragraphs[0].elements[0].format?.bold === true &&
          !window.office.getState().building
        );
      });
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [0, 3]);
      await field.evaluate((el) => el.setSelectionRange(6, 6));
      await field.press('Control+i');
      await field.pressSequentially('!');
      assert.equal(await field.textContent(), '日本語 😀!');
      await field.press('Escape');
      await page.waitForFunction(
        () =>
          window.office
            .getState()
            .editor.slides[0].shapes.find((s) => s.table)
            .table.cells[0][0].text.endsWith('!') && !window.office.getState().building,
      );
      await page.reload();
      await page.locator('.shape-hit').last().waitFor();
      const formatted = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const formattedCell = pptx.getTableCell(
        pptx.getSlideTables(pptx.getSlides(formatted)[0])[0],
        0,
        0,
      );
      const runs = pptx.getTableCellParagraphs(formattedCell)[0].elements;
      assert.equal(runs[0].text, '日本語');
      assert.equal(runs[0].format.bold, true);
      assert.notEqual(runs[1].format.bold, true);
      assert.equal(runs.at(-1).text, '!');
      assert.equal(runs.at(-1).format.italic, true);
      await page.waitForFunction(() => !window.office.getState().building);
      await page
        .locator('.shape-hit')
        .last()
        .dblclick({ position: { x: 30, y: 15 } });
      field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.fill('First\nSecond\nThird');
      await field.evaluate((el) => el.setSelectionRange(6, 12));
      await page.locator('[data-edit=align-center]').first().click();
      const waitAlignment = (expected) =>
        page.waitForFunction((expected) => {
          const paras = window.office.getState().editor.slides[0].shapes.find((s) => s.table).table
            .cells[0][0].paragraphs;
          return (
            paras.length === 3 && paras[1].align === expected && !window.office.getState().building
          );
        }, expected);
      await waitAlignment('center');
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [6, 12]);
      assert.equal(
        await field
          .locator('.edit-paragraph')
          .nth(1)
          .evaluate((el) => getComputedStyle(el).textAlign),
        'center',
      );
      await field.press('Escape');
      await page.locator('[data-edit=undo]').first().click();
      await waitAlignment(null);
      await page.locator('[data-edit=redo]').first().click();
      await waitAlignment('center');
      await page.reload();
      await page.locator('.shape-hit').last().waitFor();
      await waitAlignment('center');
      const aligned = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const alignedCell = pptx.getTableCell(
        pptx.getSlideTables(pptx.getSlides(aligned)[0])[0],
        0,
        0,
      );
      assert.deepEqual(
        pptx.getTableCellParagraphs(alignedCell).map((p) => p.align),
        [null, 'center', null],
      );
      await page
        .locator('.shape-hit')
        .last()
        .dblclick({ position: { x: 30, y: 15 } });
      field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.evaluate((el) => el.setSelectionRange(6, 12));
      await page.locator('[data-edit=line-spacing]').first().click();
      await page.getByRole('menuitem', { name: '1.5', exact: true }).click();
      const waitSpacing = (expected) =>
        page.waitForFunction((expected) => {
          const paras = window.office.getState().editor.slides[0].shapes.find((s) => s.table).table
            .cells[0][0].paragraphs;
          return (
            (paras[1].properties?.lineSpacing?.value ?? null) === expected &&
            !window.office.getState().building
          );
        }, expected);
      await waitSpacing(1.5);
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [6, 12]);
      assert.equal(
        await field
          .locator('.edit-paragraph')
          .nth(1)
          .evaluate((el) => el.style.lineHeight),
        '1.5',
      );
      await field.press('Escape');
      await page.locator('[data-edit=undo]').first().click();
      await waitSpacing(null);
      await page.locator('[data-edit=redo]').first().click();
      await waitSpacing(1.5);
      await page.reload();
      await waitSpacing(1.5);
      const spaced = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      assert.deepEqual(
        pptx
          .getTableCellParagraphs(
            pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(spaced)[0])[0], 0, 0),
          )
          .map((p) => p.properties?.lineSpacing ?? null),
        [null, { kind: 'pct', value: 1.5 }, null],
      );
      await page.locator('.shape-hit').last().waitFor();
      await page
        .locator('.shape-hit')
        .last()
        .dblclick({ position: { x: 30, y: 15 } });
      field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.evaluate((el) => el.setSelectionRange(0, 12));
      await page.locator('[data-edit=numbering]').first().click();
      const waitBullets = (expected) =>
        page.waitForFunction((expected) => {
          const paras = window.office.getState().editor.slides[0].shapes.find((s) => s.table).table
            .cells[0][0].paragraphs;
          return (
            paras.length === 3 &&
            paras[0].properties?.bullet === expected &&
            paras[1].properties?.bullet === expected &&
            paras[2].properties?.bullet === 'none' &&
            !window.office.getState().building
          );
        }, expected);
      await waitBullets('number');
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [0, 12]);
      assert.deepEqual(
        await field
          .locator('.edit-paragraph')
          .evaluateAll((els) => els.map((el) => el.dataset.bullet ?? null)),
        ['1.', '2.', null],
      );
      await page.locator('[data-edit=bullets]').first().click();
      await waitBullets('bullet');
      await field.waitFor();
      assert.deepEqual(
        await field
          .locator('.edit-paragraph')
          .evaluateAll((els) => els.map((el) => el.dataset.bullet ?? null)),
        ['•', '•', null],
      );
      await page.locator('[data-edit=bullets]').first().click();
      await waitBullets('none');
      await field.press('Escape');
      await page.locator('[data-edit=undo]').first().click();
      await waitBullets('bullet');
      await page.locator('[data-edit=redo]').first().click();
      await waitBullets('none');
      await page.reload();
      await waitBullets('none');
      const bulleted = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      assert.deepEqual(
        pptx
          .getTableCellParagraphs(
            pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(bulleted)[0])[0], 0, 0),
          )
          .map((p) => p.properties?.bullet ?? null),
        ['none', 'none', 'none'],
      );
      await page.locator('.shape-hit').last().waitFor();
      const lastCell = await page.locator('[data-pptx-cell="8,3"]').boundingBox();
      assert.ok(lastCell);
      await page.mouse.dblclick(lastCell.x + lastCell.width / 2, lastCell.y + lastCell.height / 2);
      field = page.getByRole('textbox', { name: 'Edit cell 9, 4', exact: true });
      await field.fill('Last cell');
      await field.press('Tab');
      field = page.getByRole('textbox', { name: 'Edit cell 10, 1', exact: true });
      await field.waitFor();
      await field.fill('New row');
      await field.press('Shift+Tab');
      field = page.getByRole('textbox', { name: 'Edit cell 9, 4', exact: true });
      await field.waitFor();
      assert.equal(await field.textContent(), 'Last cell');
      await field.press('Escape');
      const waitRows = (rows) =>
        page.waitForFunction(
          (rows) =>
            window.office.getState().editor.slides[0].shapes.find((s) => s.table).table.cells
              .length === rows && !window.office.getState().building,
          rows,
        );
      await waitRows(10);
      await page.locator('[data-edit=undo]').first().click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.find((s) => s.table).table.cells[9][0]
            .text === '' && !window.office.getState().building,
      );
      await page.locator('[data-edit=undo]').first().click();
      await waitRows(9);
      await page.locator('[data-edit=redo]').first().click();
      await waitRows(10);
      await page.locator('[data-edit=redo]').first().click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.find((s) => s.table).table.cells[9][0]
            .text === 'New row' && !window.office.getState().building,
      );
      await page.reload();
      await waitRows(10);
      const appended = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const appendedTable = pptx.getSlideTables(pptx.getSlides(appended)[0])[0];
      assert.deepEqual(pptx.getTableDimensions(appendedTable), { rows: 10, cols: 4 });
      assert.equal(pptx.getTableCellText(pptx.getTableCell(appendedTable, 8, 3)), 'Last cell');
      assert.equal(pptx.getTableCellText(pptx.getTableCell(appendedTable, 9, 0)), 'New row');
      await page.locator('.shape-hit').last().waitFor();
      await page
        .locator('.shape-hit')
        .last()
        .dblclick({ position: { x: 30, y: 15 } });
      field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.waitFor();
      await field.evaluate((el) => el.setSelectionRange(1, 4));
      await page.getByRole('tab', { name: 'Table Design', exact: true }).click();
      await page.getByLabel('Pen Weight', { exact: true }).selectOption('3');
      await page.getByLabel('Pen Style', { exact: true }).selectOption('dash');
      const bordersButton = page.locator('[data-edit=table-borders]');
      await bordersButton.click();
      await page.getByRole('menuitem', { name: 'Outside Borders', exact: true }).click();
      await page
        .locator('#stage')
        .getByRole('textbox', { name: 'Edit cell 1, 1', exact: true })
        .waitFor();
      const waitBorder = async (noFill) => {
        await page.waitForFunction(() => !window.office.getState().building);
        const saved = await pptx.loadPresentation(
          await (await page.request.get(url + '/deck.pptx')).body(),
        );
        const cell = pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(saved)[0])[0], 0, 0);
        const top = pptx.getTableCellBorders(saved, cell).top;
        if (noFill) assert.equal(top.noFill, true);
        else {
          assert.equal(top.widthEmu, 38100);
          assert.equal(top.dash, 'dash');
        }
      };
      await waitBorder(false);
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      await bordersButton.click();
      await page.getByRole('menuitem', { name: 'No Border', exact: true }).click();
      await field.waitFor();
      await waitBorder(true);
      const shading = page.locator('[data-edit=table-shading]');
      await shading.click();
      await page.getByRole('menuitem', { name: 'Red', exact: true }).click();
      await page.locator('[data-pptx-cell="0,0"][fill="#FF0000"]').waitFor();
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      await shading.click();
      await page.getByRole('menuitem', { name: 'No Fill', exact: true }).click();
      await page.locator('[data-pptx-cell="0,0"][fill="none"]').waitFor();
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      const headerOption = page.getByRole('checkbox', { name: 'Header Row', exact: true });
      assert.equal(await headerOption.isChecked(), true);
      await headerOption.uncheck();
      const waitHeader = (value) =>
        page.waitForFunction(
          (value) =>
            window.office.getState().editor.slides[0].shapes.find((s) => s.table).table.style
              .firstRow === value && !window.office.getState().building,
          value,
        );
      await waitHeader(false);
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      await field.press('Escape');
      await page.locator('[data-edit=undo]').first().click();
      await waitHeader(true);
      assert.equal(await headerOption.isChecked(), true);
      await page.locator('[data-edit=redo]').first().click();
      await waitHeader(false);
      assert.equal(await headerOption.isChecked(), false);
      await headerOption.click({ trial: true });
      await page.screenshot({ path: '/tmp/pptx-table-design.png' });
      const firstCell = await page.locator('#stage [data-pptx-cell="0,0"]').first().boundingBox();
      assert.ok(firstCell);
      await page.mouse.dblclick(
        firstCell.x + firstCell.width / 2,
        firstCell.y + firstCell.height / 2,
      );
      field = page.getByRole('textbox', { name: 'Edit cell 1, 1', exact: true });
      await field.waitFor();
      await field.evaluate((el) => el.setSelectionRange(1, 4));
      await page.getByRole('tab', { name: 'Table Layout', exact: true }).click();

      for (const name of ['Distribute Rows', 'Distribute Columns']) {
        await page.getByRole('button', { name, exact: true }).click();
        await page.waitForFunction(() => !window.office.getState().building);
        await field.waitFor();
        assert.deepEqual(
          await field.evaluate((el) => [el.selectionStart, el.selectionEnd]),
          [1, 4],
        );
      }
      const heightInput = page.getByRole('spinbutton', { name: 'Table Row Height', exact: true });
      const beforeResize = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const resizeTable = pptx.getSlideTables(pptx.getSlides(beforeResize)[0])[0];
      const expectedHeight = Math.max(
        457200,
        ...pptx
          .getTableCells(resizeTable)[0]
          .map((_, column) => measureTableCellHeight(beforeResize, resizeTable, 0, column) ?? 0),
      );
      await heightInput.fill('0.5');
      await heightInput.press('Enter');
      await page.waitForFunction((expected) => {
        const state = window.office.getState();
        const table = state.editor.slides[0].shapes.find((s) => s.table).table;
        return !state.building && table.rowHeights[0] === expected;
      }, expectedHeight);
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      assert.ok(Math.abs(Number(await heightInput.inputValue()) - expectedHeight / 914400) < 0.001);
      const resized = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      assert.equal(
        pptx.getTableRowHeights(pptx.getSlideTables(pptx.getSlides(resized)[0])[0])[0],
        expectedHeight,
      );
      await page.getByRole('button', { name: 'Cell Margins', exact: true }).click();
      let marginsDialog = page.getByRole('dialog', { name: 'Cell Margins', exact: true });
      await marginsDialog.getByRole('spinbutton', { name: 'Left margin' }).fill('0.2');
      await marginsDialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      await page.getByRole('button', { name: 'Cell Margins', exact: true }).click();
      marginsDialog = page.getByRole('dialog', { name: 'Cell Margins', exact: true });
      assert.equal(
        await marginsDialog.getByRole('spinbutton', { name: 'Left margin' }).inputValue(),
        '0.1',
      );
      await marginsDialog.getByRole('spinbutton', { name: 'Left margin' }).fill('0.2');
      await page.screenshot({ path: '/tmp/pptx-table-margins.png' });
      await marginsDialog.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(() => {
        const state = window.office.getState();
        return (
          !state.building &&
          state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][0].margins.left ===
            182880
        );
      });
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      const inset = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const insetTable = pptx.getSlideTables(pptx.getSlides(inset)[0])[0];
      assert.equal(pptx.getTableCellMargins(pptx.getTableCell(insetTable, 0, 0)).left, 182880);
      assert.equal(pptx.getTableCellMargins(pptx.getTableCell(insetTable, 0, 1)).left, 91440);
      await page.screenshot({ path: '/tmp/pptx-table-distribution.png' });

      const waitCellLayout = (anchor, direction) =>
        page.waitForFunction(
          ({ anchor, direction }) => {
            const cell = window.office.getState().editor.slides[0].shapes.find((s) => s.table).table
              .cells[0][0];
            return (
              cell.anchor === anchor &&
              cell.direction === direction &&
              !window.office.getState().building
            );
          },
          { anchor, direction },
        );
      await page.getByRole('button', { name: 'Align Text', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Bottom', exact: true }).click();
      await waitCellLayout('bottom', null);
      await field.waitFor();
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      await page.locator('#ribbon-table-layout [data-edit=text-direction]').click();
      await page.getByRole('menuitem', { name: 'Rotate all text 270°', exact: true }).click();
      await waitCellLayout('bottom', 'vert270');
      await field.waitFor();
      assert.equal(await field.evaluate((el) => getComputedStyle(el).writingMode), 'vertical-lr');
      assert.deepEqual(await field.evaluate((el) => [el.selectionStart, el.selectionEnd]), [1, 4]);
      await page.screenshot({ path: '/tmp/pptx-table-cell-direction.png' });
      await field.press('Escape');
      await page.locator('[data-edit=undo]').first().click();
      await waitCellLayout('bottom', null);
      await page.locator('[data-edit=redo]').first().click();
      await waitCellLayout('bottom', 'vert270');
      await page.reload();
      await waitCellLayout('bottom', 'vert270');
      const directed = await pptx.loadPresentation(
        await (await page.request.get(url + '/deck.pptx')).body(),
      );
      const directedTable = pptx.getSlideTables(pptx.getSlides(directed)[0])[0];
      assert.equal(pptx.isTableCellNoFill(pptx.getTableCell(directedTable, 0, 0)), true);
      assert.equal(pptx.isTableCellNoFill(pptx.getTableCell(directedTable, 0, 1)), false);
      assert.equal(pptx.getTableStyleFlags(directedTable).firstRow, false);
      assert.equal(pptx.getTableCellAnchor(pptx.getTableCell(directedTable, 0, 0)), 'bottom');
      assert.equal(
        pptx.getTableCellTextDirection(pptx.getTableCell(directedTable, 0, 0)),
        'vert270',
      );
      assert.equal(pptx.getTableCellTextDirection(pptx.getTableCell(directedTable, 0, 1)), null);
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
