import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'object clipboard retains copied contents and cuts immediately with undo',
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
      const text = () =>
        page.evaluate(() => window.office.getState().editor.slides[0].shapes.map((s) => s.text));
      await page.locator('.shape-hit').first().click();
      await page.locator('[data-edit=copy]').click();
      await page.waitForFunction(() => !document.querySelector('[data-edit=paste]').disabled);
      await page.locator('.shape-hit').first().dblclick();
      const field = page.getByRole('textbox', { name: 'Edit text', exact: true });
      await field.press('Meta+a');
      await page.keyboard.type('Changed original');
      await field.press('Escape');
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes[0].text === 'Changed original' &&
          !window.office.getState().building,
      );
      await page.locator('[data-edit=paste]').click();
      await count(2);
      assert.deepEqual(await text(), ['Changed original', 'Alpha [a] ALPHA']);
      await page.locator('[data-edit=cut]').click();
      await count(1);
      assert.deepEqual(await text(), ['Changed original']);
      await page.waitForFunction(() => !document.querySelector('[data-edit=paste]').disabled);
      await page.locator('[data-edit=paste]').click();
      await count(2);
      assert.deepEqual(await text(), ['Changed original', 'Alpha [a] ALPHA']);
      await page.locator('[data-edit=undo]').first().click();
      await count(1);
      await page.locator('[data-edit=undo]').first().click();
      await count(2);
      assert.deepEqual(await text(), ['Changed original', 'Alpha [a] ALPHA']);
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      assert.deepEqual(await text(), ['Changed original', 'Alpha [a] ALPHA']);
      // Clipboard contents are intentionally session-local: capture again after reload.
      await page.locator('.shape-hit').last().click();
      await page.locator('[data-edit=copy]').click();
      await page.waitForFunction(() => !document.querySelector('[data-edit=paste]').disabled);
      let releasePaste;
      let pasteStarted;
      const pendingPaste = new Promise((resolve) => {
        releasePaste = resolve;
      });
      const started = new Promise((resolve) => {
        pasteStarted = resolve;
      });
      await page.route('**/edit', async (route) => {
        if (route.request().postDataJSON()?.command?.type === 'paste') {
          pasteStarted();
          await pendingPaste;
        }
        await route.continue();
      });
      await page.locator('[data-edit=paste]').click();
      await started;
      await page.locator('.thumbnail').nth(1).click();
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      releasePaste();
      await count(3);
      await page.waitForFunction(() => !document.querySelector('[data-edit=paste]').disabled);
      assert.equal(await page.evaluate(() => window.office.getState().index), 1);
      assert.equal(await page.locator('.shape-hit.selected').count(), 0);
      assert.equal(
        await page
          .locator('.thumbnail')
          .nth(1)
          .evaluate((el) => el === document.activeElement),
        true,
      );
      assert.deepEqual(
        await page.evaluate(() =>
          window.office.getState().editor.slides[1].shapes.map((s) => s.text),
        ),
        ['Beta alpha'],
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
