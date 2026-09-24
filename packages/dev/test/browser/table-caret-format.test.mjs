import * as pptx from '@office-kit/pptx';
import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test('table cell caret formatting retains immediate text input', { timeout: 120000 }, async () => {
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
    await field.fill('日本語 😀');
    await field.press('Escape');
    await page.waitForFunction(() => {
      const state = window.office.getState();
      return (
        !state.building &&
        state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][0].text === '日本語 😀'
      );
    });
    await page.getByRole('tab', { name: 'Home', exact: true }).click();
    await page.locator('[data-edit=undo]').first().click({ trial: true });
    const box = await page.locator('#stage [data-pptx-cell="0,0"]').first().boundingBox();
    await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
    await field.waitFor();
    await field.evaluate((el) => el.setSelectionRange(0, 3));
    await page.locator('[data-edit=bold]').first().click();
    await page.waitForFunction(() => {
      const state = window.office.getState();
      return (
        !state.building &&
        state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][0].paragraphs[0]
          .elements[0].format?.bold === true
      );
    });
    await field.waitFor();
    await field.evaluate((el) => el.setSelectionRange(6, 6));
    await field.press('Control+i');
    await field.pressSequentially('!');
    assert.equal(await field.textContent(), '日本語 😀!');
    await field.press('Escape');
    await page.waitForFunction(() => {
      const state = window.office.getState();
      return (
        !state.building &&
        state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][0].text === '日本語 😀!'
      );
    });
    const saved = await pptx.loadPresentation(
      await (await page.request.get(url + '/deck.pptx')).body(),
    );
    const cell = pptx.getTableCell(pptx.getSlideTables(pptx.getSlides(saved)[0])[0], 0, 0);
    assert.equal(pptx.getTableCellText(cell), '日本語 😀!');
    const runs = pptx.getTableCellParagraphs(cell)[0].elements;
    assert.equal(runs[0].text, '日本語');
    assert.equal(runs[0].format.bold, true);
    assert.equal(runs.at(-1).text, '!');
    assert.equal(runs.at(-1).format.italic, true);
    await page.locator('[data-edit=undo]').first().click({ trial: true });
    // Only notifications for our own saved history are ignored. Editing the
    // journal externally must still rebuild and update the visible document.
    const historyPath = file + '.edits.json';
    const history = JSON.parse(await readFile(historyPath, 'utf8'));
    history.cursor--;
    await writeFile(historyPath, JSON.stringify(history));
    await page.waitForFunction(() => {
      const state = window.office.getState();
      return (
        !state.building &&
        state.editor.slides[0].shapes.find((s) => s.table).table.cells[0][0].text === '日本語 😀'
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
});
