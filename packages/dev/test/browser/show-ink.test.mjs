import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'slideshow pen keeps multi-slide ink, supports undo and discards temporary strokes',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-link-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>Link target</Text></Slide><Slide><Text x={1} y={1} width={8} height={1}>Return</Text></Slide></Presentation>`,
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
        proc.on('error', reject);
        proc.on('exit', (code) => reject(new Error('Server exited ' + code)));
      });
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(url);
      await page.waitForFunction(() => window.office?.getState().editor?.slides.length === 2);
      const start = async () => {
        await page.locator('#present').click();
        await page.waitForFunction(() => window.office.getState().presenting);
        await page.keyboard.press('Meta+p');
      };
      const draw = async () => {
        const box = await page.locator('#slide').boundingBox();
        await page.mouse.move(box.x + box.width * 0.2, box.y + box.height * 0.3);
        await page.mouse.down();
        await page.mouse.move(box.x + box.width * 0.4, box.y + box.height * 0.5, { steps: 8 });
        await page.mouse.up();
      };
      await start();
      await draw();
      assert.equal(await page.evaluate(() => window.office.getState().index), 0);
      assert.equal(await page.locator('#show-ink path').count(), 1);
      const path = await page.locator('#show-ink path').getAttribute('d');
      assert.ok(path.split('L').length >= 8);
      await page.keyboard.press('ArrowRight');
      assert.equal(await page.locator('#show-ink path').count(), 0);
      await draw();
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('#show-ink path').count(), 1);
      await page.screenshot({ path: '/tmp/pptx-show-ink.png' });
      await page.keyboard.press('Escape');
      assert.equal(await page.evaluate(() => window.office.getState().presenting), true);
      assert.equal(await page.locator('#slide').getAttribute('data-show-cursor'), null);
      assert.equal(await page.locator('#show-ink path').count(), 1);
      await page.keyboard.press('Escape');
      await page
        .getByRole('dialog', { name: 'Keep ink annotations' })
        .getByRole('button', { name: 'Keep', exact: true })
        .click();
      await page.waitForFunction(() =>
        window.office.getState().editor.slides.every((s) => s.shapes.some((x) => x.kind === 'ink')),
      );
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await page.waitForFunction(() =>
        window.office
          .getState()
          .editor.slides.every((s) => s.shapes.every((x) => x.kind !== 'ink')),
      );
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await page.waitForFunction(() =>
        window.office.getState().editor.slides.every((s) => s.shapes.some((x) => x.kind === 'ink')),
      );
      await page.waitForFunction(
        () =>
          !document.querySelector('[data-edit=undo]').disabled &&
          !window.office.getState().building,
      );
      const originalInk = await page.evaluate(() =>
        window.office.getState().editor.slides[0].shapes.find((s) => s.kind === 'ink'),
      );
      const inkHit = page.locator('.shape-hit[data-shape-id="' + originalInk.id + '"]');
      const hitBox = await inkHit.boundingBox();
      await page.mouse.move(hitBox.x + hitBox.width / 2, hitBox.y + hitBox.height / 2);
      await page.mouse.down();
      await page.mouse.move(hitBox.x + hitBox.width / 2 + 60, hitBox.y + hitBox.height / 2 + 30, {
        steps: 5,
      });
      await page.mouse.up();
      await page.waitForFunction(
        (x) =>
          window.office.getState().editor.slides[0].shapes.find((s) => s.kind === 'ink').bounds.x >
          x,
        originalInk.bounds.x,
      );
      await page.waitForFunction(
        () =>
          !document.querySelector('[data-edit=undo]').disabled &&
          !window.office.getState().building,
      );
      await inkHit.click({ button: 'right' });
      await page.getByRole('menuitem', { name: 'Size and Position...', exact: true }).click();
      const properties = page.getByRole('complementary', { name: 'Format Shape', exact: true });
      await properties.getByRole('checkbox', { name: 'Lock aspect ratio', exact: true }).check();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes.find((s) => s.kind === 'ink')
            .aspectRatioLocked && !document.querySelector('[data-edit=undo]').disabled,
      );
      await properties.getByRole('button', { name: 'Close Format Shape', exact: true }).click();
      await inkHit.locator('[data-handle="se"]').hover();
      const handle = await inkHit.locator('[data-handle="se"]').boundingBox();
      await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2);
      await page.mouse.down();
      await page.mouse.move(handle.x + 50, handle.y + 30, { steps: 5 });
      await page.mouse.up();
      await page.waitForFunction(
        (w) =>
          window.office.getState().editor.slides[0].shapes.find((s) => s.kind === 'ink').bounds.w >
          w,
        originalInk.bounds.w,
      );
      const resizedInk = await page.evaluate(() =>
        window.office.getState().editor.slides[0].shapes.find((s) => s.kind === 'ink'),
      );
      assert.equal(resizedInk.aspectRatioLocked, true);
      assert.ok(
        Math.abs(
          resizedInk.bounds.w / resizedInk.bounds.h - originalInk.bounds.w / originalInk.bounds.h,
        ) < 0.00001,
      );
      await page.reload();
      await page.waitForFunction(
        () =>
          !window.office?.getState().building &&
          window.office?.getState().editor?.slides[0]?.shapes.some((s) => s.kind === 'ink'),
      );
      assert.deepEqual(
        await page.evaluate(
          () =>
            window.office.getState().editor.slides[0].shapes.find((s) => s.kind === 'ink').bounds,
        ),
        resizedInk.bounds,
      );
      const revision = await page.evaluate(() => window.office.getState().revision);
      await start();
      await draw();
      await page.keyboard.press('Shift+e');
      assert.equal(await page.locator('#show-ink path').count(), 0);
      await draw();
      await page.keyboard.press('Meta+.');
      await page.route('**/edit', (route) =>
        route.fulfill({
          status: 409,
          contentType: 'application/json',
          body: JSON.stringify({ error: 'Source changed; retry after refresh.' }),
        }),
      );
      await page
        .getByRole('dialog', { name: 'Keep ink annotations' })
        .getByRole('button', { name: 'Keep', exact: true })
        .click();
      await page
        .getByRole('dialog')
        .getByRole('alert')
        .filter({ hasText: 'Source changed' })
        .waitFor();
      await page.unroute('**/edit');
      await page.getByRole('button', { name: 'Discard', exact: true }).click();
      assert.equal(await page.evaluate(() => window.office.getState().revision), revision);
      // Browser-controlled fullscreen exit must use the same annotation prompt.
      await start();
      await page.waitForFunction(() => !!document.fullscreenElement);
      await draw();
      await page.evaluate(() => document.exitFullscreen());
      await page.getByRole('dialog', { name: 'Keep ink annotations' }).waitFor();
      assert.equal(await page.evaluate(() => window.office.getState().presenting), false);
      await page.getByRole('button', { name: 'Discard', exact: true }).click();
      // Switching tools while dragging terminates the pen stroke immediately.
      await start();
      const box = await page.locator('#slide').boundingBox();
      await page.mouse.move(box.x + 100, box.y + 100);
      await page.mouse.down();
      await page.mouse.move(box.x + 150, box.y + 150);
      const beforeSwitch = await page.locator('#show-ink path').getAttribute('d');
      const indexBeforeSwitch = await page.evaluate(() => window.office.getState().index);
      await page.keyboard.press('Escape');
      assert.equal(await page.evaluate(() => window.office.getState().presenting), true);
      await page.mouse.move(box.x + 250, box.y + 250);
      assert.equal(await page.locator('#show-ink path').getAttribute('d'), beforeSwitch);
      await page.mouse.up();
      assert.equal(await page.evaluate(() => window.office.getState().index), indexBeforeSwitch);
      await page.keyboard.press('Escape');
      await page.getByRole('button', { name: 'Discard', exact: true }).click();
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
