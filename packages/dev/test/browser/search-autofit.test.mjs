import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'replace all measures autofit across slides and undoes in one step',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-size-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={2} height={0.5} autoFit="normal">seed</Text></Slide><Slide><Text x={1} y={1} width={2} height={0.5} autoFit="shape">seed</Text></Slide></Presentation>`,
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
      const before = (await state()).editor.slides.map((s) => s.shapes[0]);
      await page.getByRole('button', { name: 'Search options', exact: true }).click();
      await page.getByRole('menuitem', { name: 'Find and Replace', exact: true }).click();
      const dialog = page.getByRole('dialog', { name: 'Find and Replace', exact: true });
      await dialog.getByLabel('Find what:', { exact: true }).fill('seed');
      const replacement = 'A much longer passage with multiple words. '.repeat(12);
      await dialog.getByLabel('Replace with:', { exact: true }).fill(replacement);
      await dialog.getByRole('button', { name: 'Replace All', exact: true }).click();
      await page.waitForFunction(
        () =>
          document.querySelector('.find-replace-dialog [role=status]').textContent ===
          'Replaced 2 occurrences.',
      );
      const after = (await state()).editor.slides.map((s) => s.shapes[0]);
      assert.equal(after[0].text, replacement);
      assert.equal(after[1].text, replacement);
      assert.ok(after[0].autoFitParams.fontScale < 1);
      assert.deepEqual(after[0].bounds, before[0].bounds);
      assert.ok(after[1].bounds.h > before[1].bounds.h);
      await dialog.getByRole('button', { name: 'Close', exact: true }).click();
      await page.locator('[data-edit=undo]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'seed',
      );
      assert.deepEqual(
        (await state()).editor.slides.map((s) => s.shapes[0]),
        before,
      );
      await page.locator('[data-edit=redo]').click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].text !== 'seed',
      );
      assert.deepEqual(
        (await state()).editor.slides.map((s) => s.shapes[0]),
        after,
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
