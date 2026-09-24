import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'Selection pane drag reaches offscreen layers and preserves undo',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-selection-scroll-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide>${Array.from({ length: 45 }, (_, i) => `<Text x={2} y={2} width={3} height={1}>Layer ${i + 1}</Text>`).join('')}</Slide></Presentation>`,
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

      const state = () => page.evaluate(() => window.office.getState());
      const settled = async (revision) =>
        page.waitForFunction((rev) => {
          const s = window.office.getState();
          return (
            s.revision > rev &&
            !s.building &&
            !document.querySelector('[data-edit="undo"]').disabled
          );
        }, revision);
      const fullStage = await page.locator('#stage').boundingBox();
      await page.locator('[data-edit="selection"]:visible').first().click();
      await page.waitForFunction(() => {
        const stage = document.querySelector('#stage').getBoundingClientRect();
        const pane = document.querySelector('#selection-pane').getBoundingClientRect();
        return stage.right <= pane.left;
      });
      await page.getByRole('button', { name: 'Close Selection Pane', exact: true }).click();
      assert.equal((await page.locator('#stage').boundingBox()).width, fullStage.width);
      await page.locator('[data-edit="selection"]:visible').first().click();
      const before = (await state()).editor.slides[0].shapes;
      const pane = page.locator('#selection-pane');
      const box = await pane.boundingBox();
      const source = page.locator(`[data-selection-id="${before.at(-1).id}"]`);
      const sourceBox = await source.boundingBox();
      await page.mouse.move(sourceBox.x + 20, sourceBox.y + sourceBox.height / 2);
      await page.mouse.down();
      await page.mouse.move(sourceBox.x + 25, sourceBox.y + sourceBox.height / 2, { steps: 3 });
      await page.mouse.move(box.x + 70, box.y + box.height - 8, { steps: 10 });
      await page.waitForFunction(() => {
        const pane = document.querySelector('#selection-pane');
        return pane.scrollTop + pane.clientHeight >= pane.scrollHeight - 2;
      });
      const target = page.locator(`[data-selection-id="${before[0].id}"]`);
      const targetBox = await target.boundingBox();
      const revision = (await state()).revision;
      await page.mouse.move(targetBox.x + 20, targetBox.y + targetBox.height - 2, { steps: 3 });
      await page.mouse.move(targetBox.x + 21, targetBox.y + targetBox.height - 3);
      await page.mouse.up();
      await settled(revision);
      const expected = [before.at(-1).id, ...before.slice(0, -1).map((shape) => shape.id)];
      assert.deepEqual(
        (await state()).editor.slides[0].shapes.map((shape) => shape.id),
        expected,
      );
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      assert.deepEqual(
        (await state()).editor.slides[0].shapes.map((shape) => shape.id),
        expected,
      );
      const savedRevision = (await state()).revision;
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction((rev) => {
        const state = window.office.getState();
        return (
          state.revision > rev &&
          !state.building &&
          !document.querySelector('[data-edit="redo"]').disabled
        );
      }, savedRevision);
      assert.deepEqual((await state()).editor.slides[0].shapes, before);
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
