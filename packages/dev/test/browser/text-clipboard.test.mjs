import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'Edit menu text clipboard preserves selection, cuts immediately and supports undo',
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
      const field = page.getByRole('textbox', { name: 'Edit text', exact: true });
      await page.locator('.shape-hit').first().dblclick();
      const waitText = async (text) =>
        page.waitForFunction(({ field, text }) => field.value === text, {
          field: await field.elementHandle(),
          text,
        });
      const select = (start, end) =>
        field.evaluate(
          (f, range) => {
            f.focus();
            f.setSelectionRange(...range);
          },
          [start, end],
        );
      const menu = async (name) => {
        await page
          .getByRole('menubar')
          .getByRole('menuitem', { name: 'Edit', exact: true })
          .click();
        await page.getByRole('menuitem', { name, exact: true }).click();
      };
      await select(0, 5);
      await menu('Copy');
      assert.equal(await page.evaluate(() => navigator.clipboard.readText()), 'Alpha');
      assert.equal(await field.evaluate((f) => f.value), 'Alpha [a] ALPHA');
      await select(6, 9);
      await menu('Cut');
      await waitText('Alpha  ALPHA');
      assert.equal(await page.evaluate(() => navigator.clipboard.readText()), '[a]');
      await menu('Undo');
      assert.equal(await field.evaluate((f) => f.value), 'Alpha [a] ALPHA');
      await select(6, 9);
      await page.evaluate(() => navigator.clipboard.writeText('Beta\nGamma'));
      await menu('Paste');
      await waitText('Alpha Beta\nGamma ALPHA');
      await menu('Undo');
      assert.equal(await field.evaluate((f) => f.value), 'Alpha [a] ALPHA');
      await select(6, 9);
      await page.evaluate(() => navigator.clipboard.writeText(''));
      await menu('Paste');
      assert.equal(await field.evaluate((f) => f.value), 'Alpha [a] ALPHA');
      await page.evaluate(() => {
        window.savedWriteText = navigator.clipboard.writeText;
        navigator.clipboard.writeText = async () => {
          throw new Error('Clipboard denied');
        };
      });
      await menu('Cut');
      assert.equal(await field.evaluate((f) => f.value), 'Alpha [a] ALPHA');
      await page.evaluate(() => {
        navigator.clipboard.writeText = window.savedWriteText;
      });
      // Clipboard access may finish after the user moves the selection.
      // Resolve it explicitly so this covers the race without timing assumptions.
      for (const operation of ['Cut', 'Paste']) {
        await select(6, 9);
        await page.evaluate((operation) => {
          const method = operation === 'Cut' ? 'writeText' : 'readText';
          window.savedClipboardMethod = navigator.clipboard[method];
          navigator.clipboard[method] = () =>
            new Promise((resolve) => {
              window.resolveClipboard = resolve;
            });
        }, operation);
        await menu(operation);
        assert.equal(await page.locator('[data-edit=paste]').isDisabled(), true);
        await select(0, 5);
        await page.evaluate(() => window.resolveClipboard('delayed text'));
        await page.waitForFunction(() =>
          document.getElementById('edit-message').textContent.includes('selection changed'),
        );
        assert.equal(await field.evaluate((f) => f.value), 'Alpha [a] ALPHA');
        assert.deepEqual(await field.evaluate((f) => [f.selectionStart, f.selectionEnd]), [0, 5]);
        assert.equal(await page.locator('[data-edit=paste]').isDisabled(), false);
        await page.evaluate((operation) => {
          navigator.clipboard[operation === 'Cut' ? 'writeText' : 'readText'] =
            window.savedClipboardMethod;
        }, operation);
      }
      await select(6, 9);
      await field.dispatchEvent('contextmenu', {
        bubbles: true,
        composed: true,
        clientX: 600,
        clientY: 350,
      });
      assert.equal(await page.getByRole('menuitem', { name: 'Duplicate', exact: true }).count(), 0);
      await page.getByRole('menuitem', { name: 'Delete', exact: true }).click();
      await waitText('Alpha  ALPHA');
      await menu('Undo');
      await waitText('Alpha [a] ALPHA');
      await select(0, 0);
      await page.getByRole('menubar').getByRole('menuitem', { name: 'Edit', exact: true }).click();
      assert.equal(
        await page.getByRole('menuitem', { name: 'Cut', exact: true }).isDisabled(),
        true,
      );
      assert.equal(
        await page.getByRole('menuitem', { name: 'Copy', exact: true }).isDisabled(),
        true,
      );
      await page.keyboard.press('Escape');
      await select(0, 5);
      await page.locator('[data-edit=copy]').click();
      assert.equal(await page.evaluate(() => navigator.clipboard.readText()), 'Alpha');
      await select(6, 9);
      await page.locator('[data-edit=paste]').click();
      await waitText('Alpha Alpha ALPHA');
      await field.press('Escape');
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes[0].text === 'Alpha Alpha ALPHA',
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
