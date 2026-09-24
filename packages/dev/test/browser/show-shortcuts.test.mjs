import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test('Mac slideshow start, exit and hidden-slide shortcuts', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-link-'));
  const file = join(dir, 'deck.tsx');
  await writeFile(
    file,
    `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1}>Link target</Text></Slide><Slide><Text x={1} y={1} width={8} height={1}>Hidden</Text></Slide><Slide><Text x={1} y={1} width={8} height={1}>Last</Text></Slide></Presentation>`,
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
    await page.waitForFunction(() => window.office?.getState().editor?.slides.length === 3);
    await page.evaluate(async () => {
      const response = await fetch('/edit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          revision: window.office.getState().revision,
          command: { type: 'slide-hidden', slide: 1, hidden: true },
        }),
      });
      if (!response.ok) throw new Error(await response.text());
      await window.office.refresh();
      window.office.selectSlide(2, true);
    });
    await page.keyboard.press('Meta+Enter');
    await page.waitForFunction(
      () => window.office.getState().presenting && !!document.fullscreenElement,
    );
    assert.equal(await page.evaluate(() => window.office.getState().index), 2);
    await page.keyboard.press('Meta+.');
    await page.waitForFunction(
      () => !window.office.getState().presenting && !document.fullscreenElement,
    );
    assert.equal(await page.locator('#show-screen').isVisible(), false);
    await page.keyboard.press('Meta+Shift+Enter');
    await page.waitForFunction(
      () => window.office.getState().presenting && !!document.fullscreenElement,
    );
    assert.equal(await page.evaluate(() => window.office.getState().index), 0);
    await page.keyboard.press('ArrowRight');
    assert.equal(await page.evaluate(() => window.office.getState().index), 2);
    await page.keyboard.press('Home');
    await page.keyboard.press('h');
    assert.equal(await page.evaluate(() => window.office.getState().index), 1);
    await page.keyboard.press('h');
    assert.equal(await page.evaluate(() => window.office.getState().index), 1);
    await page.keyboard.press('ArrowLeft');
    assert.equal(await page.evaluate(() => window.office.getState().index), 0);
    await page.keyboard.press('h');
    await page.keyboard.press('ArrowRight');
    assert.equal(await page.evaluate(() => window.office.getState().index), 2);
    // Jumping backwards to a hidden slide must resume relative to that slide,
    // not the last visible slide from which the jump began.
    await page.keyboard.press('2');
    await page.keyboard.press('Enter');
    assert.equal(await page.evaluate(() => window.office.getState().index), 1);
    assert.equal(await page.locator('#present-prev').isDisabled(), false);
    await page.keyboard.press('ArrowRight');
    assert.equal(await page.evaluate(() => window.office.getState().index), 2);
    assert.equal(await page.locator('#show-screen').isVisible(), false);
    await page.keyboard.press('2');
    await page.keyboard.press('Enter');
    await page.keyboard.press('ArrowLeft');
    assert.equal(await page.evaluate(() => window.office.getState().index), 0);
    await page.keyboard.press('-');
    await page.waitForFunction(
      () => !window.office.getState().presenting && !document.fullscreenElement,
    );
    await page.evaluate(() => window.office.selectSlide(1, true));
    await page.keyboard.press('Meta+Enter');
    await page.waitForFunction(
      () => window.office.getState().presenting && !!document.fullscreenElement,
    );
    assert.equal(await page.evaluate(() => window.office.getState().index), 1);
    await page.keyboard.press('ArrowRight');
    assert.equal(await page.evaluate(() => window.office.getState().index), 2);
    await page.keyboard.press('-');
    await page.waitForFunction(
      () => !window.office.getState().presenting && !document.fullscreenElement,
    );
    // Exit shortcuts retain unsaved annotations for the normal Keep/Discard decision.
    await page.keyboard.press('Meta+Shift+Enter');
    await page.waitForFunction(
      () => window.office.getState().presenting && !!document.fullscreenElement,
    );
    await page.keyboard.press('Meta+p');
    const box = await page.locator('#slide').boundingBox();
    await page.mouse.move(box.x + 100, box.y + 100);
    await page.mouse.down();
    await page.mouse.move(box.x + 200, box.y + 160, { steps: 4 });
    await page.mouse.up();
    assert.equal(await page.locator('#show-ink path').count(), 1);
    await page.keyboard.press('Meta+.');
    const keep = page.getByRole('dialog', { name: 'Keep ink annotations' });
    await keep.waitFor();
    await keep.getByRole('button', { name: 'Discard', exact: true }).click();
    assert.equal(await page.evaluate(() => window.office.getState().presenting), false);
    assert.deepEqual(errors, []);
  } finally {
    await browser?.close();
    proc.kill('SIGTERM');
    await new Promise((resolve) =>
      proc.exitCode !== null ? resolve() : proc.once('exit', resolve),
    );
    await rm(dir, { recursive: true, force: true });
  }
});
