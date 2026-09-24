import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'presentation search navigates literal matches and keeps edits',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-size-'));
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
      await page.goto(url);
      await page.locator('.shape-hit').first().waitFor();
      const state = () => page.evaluate(() => window.office.getState());
      const search = page.getByRole('searchbox', { name: 'Search in Presentation' });
      const field = page.getByRole('textbox', { name: 'Edit text', exact: true });
      const range = () => field.evaluate((f) => [f.selectionStart, f.selectionEnd]);
      await page.keyboard.press('Meta+f');
      assert.equal(await search.evaluate((f) => f === document.activeElement), true);
      await search.fill('alpha');
      await search.press('Enter');
      assert.deepEqual(await range(), [0, 5]);
      assert.equal(await page.locator('#find-status').textContent(), '1 of 3');
      await page.keyboard.press('Meta+g');
      assert.deepEqual(await range(), [10, 15]);
      await page.keyboard.press('Meta+g');
      assert.equal((await state()).index, 1);
      assert.deepEqual(await range(), [5, 10]);
      await page.keyboard.press('Meta+g');
      assert.equal((await state()).index, 0);
      assert.deepEqual(await range(), [0, 5]);
      await page.keyboard.press('Meta+Shift+g');
      assert.equal((await state()).index, 1);
      await page.keyboard.type('Gamma');
      await page.keyboard.press('Meta+g');
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[1].shapes[0].text === 'Beta Gamma' &&
          !window.office.getState().building,
      );
      await page.waitForFunction(
        () => document.getElementById('find-status').textContent === '1 of 2',
      );
      assert.equal((await state()).editor.slides[1].shapes[0].text, 'Beta Gamma');
      assert.equal(await page.locator('#find-status').textContent(), '1 of 2');
      await page.keyboard.press('Meta+f');
      await search.fill('[a]');
      await search.press('Enter');
      assert.deepEqual(await range(), [6, 9]);
      assert.equal(await page.locator('#find-status').textContent(), '1 of 1');
      await page.keyboard.press('Meta+f');
      await search.fill('missing');
      await search.press('Enter');
      assert.equal(await page.locator('#find-status').textContent(), 'No matches');
      await search.press('Escape');
      assert.equal(await page.locator('#find-status').textContent(), '');
      await page.getByRole('menubar').getByRole('menuitem', { name: 'Edit', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Find', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Find…', exact: true }).click();
      assert.equal(await search.evaluate((f) => f === document.activeElement), true);
      const titleBox = await page.locator('.document-title').boundingBox();
      const searchBox = await search.boundingBox();
      assert.ok(titleBox.x + titleBox.width < searchBox.x);
      await page.getByRole('menubar').getByRole('menuitem', { name: 'Edit', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Find', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Replace…', exact: true }).click();
      const dialog = page.getByRole('dialog', { name: 'Find and Replace', exact: true });
      await dialog.getByLabel('Find what:', { exact: true }).fill('alpha');
      await dialog.getByLabel('Replace with:', { exact: true }).fill('alphabet');
      await dialog.getByRole('button', { name: 'Find Next', exact: true }).click();
      assert.deepEqual(await range(), [0, 5]);
      await dialog.getByRole('button', { name: 'Replace', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'alphabet [a] ALPHA',
      );
      await dialog.getByRole('button', { name: 'Replace', exact: true }).waitFor();
      await page.waitForFunction(
        () => !document.querySelector('.find-replace-dialog [data-replace]').disabled,
      );
      assert.deepEqual(await range(), [13, 18]);
      await dialog.getByLabel('Replace with:', { exact: true }).fill('');
      await dialog.getByRole('button', { name: 'Replace', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'alphabet [a] ',
      );
      await page.waitForFunction(
        () => !document.querySelector('.find-replace-dialog [data-replace]').disabled,
      );
      await page.screenshot({ path: '/tmp/pptx-replace.png' });
      await dialog.getByRole('button', { name: 'Close', exact: true }).click();
      await page.keyboard.press('Escape');
      await page.locator('[data-edit=undo]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'alphabet [a] ALPHA',
      );
      await page.locator('[data-edit=undo]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'Alpha [a] ALPHA',
      );
      await page.getByRole('menubar').getByRole('menuitem', { name: 'Edit', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Find', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Replace…', exact: true }).click();
      await dialog.getByLabel('Find what:', { exact: true }).fill('a');
      await dialog.getByLabel('Replace with:', { exact: true }).fill('aa');
      await dialog.getByRole('button', { name: 'Replace All', exact: true }).click();
      await page.waitForFunction(
        () =>
          document.querySelector('.find-replace-dialog [role=status]').textContent ===
          'Replaced 8 occurrences.',
      );
      assert.deepEqual(
        (await state()).editor.slides.map((s) => s.shapes[0].text),
        ['aalphaa [aa] aaLPHaa', 'Betaa Gaammaa'],
      );
      await dialog.getByRole('button', { name: 'Close', exact: true }).click();
      await page.locator('[data-edit=undo]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'Alpha [a] ALPHA',
      );
      assert.equal((await state()).editor.slides[1].shapes[0].text, 'Beta Gamma');
      await page.getByRole('button', { name: 'Search options', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Find and Replace', exact: true }).click();
      await dialog.getByLabel('Find what:', { exact: true }).fill('Alpha');
      await dialog.getByLabel('Match case', { exact: true }).check();
      await dialog.getByLabel('Find whole words only', { exact: true }).check();
      await dialog.getByRole('button', { name: 'Find Next', exact: true }).click();
      assert.deepEqual(await range(), [0, 5]);
      assert.equal(await page.locator('#find-status').textContent(), '1 of 1');
      await dialog.getByLabel('Replace with:', { exact: true }).fill('Omega');
      await dialog.getByRole('button', { name: 'Replace All', exact: true }).click();
      await page.waitForFunction(
        () =>
          document.querySelector('.find-replace-dialog [role=status]').textContent ===
          'Replaced 1 occurrence.',
      );
      assert.equal((await state()).editor.slides[0].shapes[0].text, 'Omega [a] ALPHA');
      await dialog.getByLabel('Find what:', { exact: true }).fill('meg');
      await dialog.getByRole('button', { name: 'Find Next', exact: true }).click();
      assert.equal(await page.locator('#find-status').textContent(), 'No matches');
      await dialog.getByLabel('Find whole words only', { exact: true }).uncheck();
      await dialog.getByRole('button', { name: 'Find Next', exact: true }).click();
      assert.deepEqual(await range(), [1, 4]);
      await page.screenshot({ path: '/tmp/pptx-search-options.png' });
      await dialog.getByRole('button', { name: 'Close', exact: true }).click();
      await page.screenshot({ path: '/tmp/pptx-search.png' });
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
