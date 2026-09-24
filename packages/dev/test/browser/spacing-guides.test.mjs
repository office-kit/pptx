import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'equal spacing guides preview, cancel, save and reload a dragged object',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-spacing-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={2} width={1} height={1}>Left</Text><Text x={5} y={2} width={1} height={1}>Right</Text><Text x={3.5} y={2} width={1} height={1}>Move</Text></Slide></Presentation>`,
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
      const moving = page.locator('.shape-hit').nth(2);
      await moving.waitFor();
      const original = await page.evaluate(() => window.office.getState());
      const canvas = await page.locator('#slide').boundingBox();
      const start = await moving.boundingBox();
      const scale = canvas.width / original.editor.width;
      const drag = async () => {
        await page.mouse.move(start.x + start.width / 2, start.y + start.height / 2);
        await page.mouse.down();
        await page.mouse.move(
          start.x + start.width / 2 - 0.5 * 914400 * scale + 2,
          start.y + start.height / 2,
          { steps: 5 },
        );
      };
      const spacing = page.locator('.smart-guide[data-kind="spacing"][data-axis="x"]');
      await drag();
      assert.equal(await spacing.count(), 2);
      const widths = await spacing.evaluateAll((nodes) =>
        nodes.map((node) => node.getBoundingClientRect().width),
      );
      assert.ok(Math.abs(widths[0] - widths[1]) < 0.1);
      assert.equal(await spacing.locator('span').count(), 4);
      await page.keyboard.press('Escape');
      await page.mouse.up();
      assert.equal(await spacing.count(), 0);
      assert.equal(
        (await page.evaluate(() => window.office.getState())).revision,
        original.revision,
      );
      await page.keyboard.down('Meta');
      await drag();
      assert.equal(await spacing.count(), 0);
      await page.keyboard.press('Escape');
      await page.mouse.up();
      await page.keyboard.up('Meta');
      await drag();
      assert.equal(await spacing.count(), 2);
      await page.mouse.up();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          Math.abs(window.office.getState().editor.slides[0].shapes[2].bounds.x - 3 * 914400) < 1,
      );
      await page.reload();
      await moving.waitFor();
      const saved = await page.evaluate(() => window.office.getState());
      assert.equal(saved.editor.slides[0].shapes[2].bounds.x, 3 * 914400);
      assert.equal(await spacing.count(), 0);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      proc.kill('SIGTERM');
      await new Promise((resolve) =>
        proc.exitCode !== null ? resolve() : proc.once('exit', resolve),
      );
      await rm(dir, { recursive: true, force: true });
    }
  },
);
