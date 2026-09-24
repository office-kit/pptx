import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'View menu synchronizes views, ribbon shortcuts, guides and zoom',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-view-'));
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
      const context = await browser.newContext({ viewport: { width: 1600, height: 1000 } });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(url);
      await page.locator('.shape-hit').first().waitFor();
      const openView = () => page.locator('#view-menu').click();
      const item = (name) => page.getByRole('menuitem', { name, exact: true });
      const check = (name) => page.getByRole('menuitemcheckbox', { name, exact: true });
      assert.deepEqual(await page.locator('.application-menubar > button').allTextContents(), [
        'File',
        'Edit',
        'View',
        'Format',
        'Arrange',
      ]);
      await openView();
      assert.equal(await check('Normal').getAttribute('aria-checked'), 'true');
      await check('Slide Sorter').click();
      assert.equal(
        await page.locator('body').evaluate((el) => el.classList.contains('slide-sorter')),
        true,
      );
      assert.equal(await page.locator('#view-sorter').getAttribute('aria-pressed'), 'true');
      assert.equal(await page.locator('#status-normal').getAttribute('aria-pressed'), 'false');
      await page.keyboard.press('Meta+Digit1');
      assert.equal(await page.locator('#status-normal').getAttribute('aria-pressed'), 'true');
      await page.keyboard.press('Meta+Digit2');
      await openView();
      assert.equal(await check('Slide Sorter').getAttribute('aria-checked'), 'true');
      await check('Normal').click();
      await openView();
      await check('Ribbon').click();
      assert.equal(
        await page.locator('body').evaluate((el) => el.classList.contains('ribbon-collapsed')),
        true,
      );
      await page.keyboard.press('Meta+Alt+KeyR');
      assert.equal(
        await page.locator('body').evaluate((el) => el.classList.contains('ribbon-collapsed')),
        false,
      );
      await openView();
      await item('Grid and Guides').hover();
      const initialSmart = await check('Smart Guides').getAttribute('aria-checked');
      await check('Smart Guides').click();
      await openView();
      await item('Grid and Guides').hover();
      assert.equal(
        await check('Smart Guides').getAttribute('aria-checked'),
        String(initialSmart !== 'true'),
      );
      await page.keyboard.press('Escape');
      await page.keyboard.press('Escape');
      await openView();
      await item('Zoom').hover();
      await item('Zoom...').click();
      assert.equal(await page.locator('#zoom-dialog').evaluate((el) => el.open), true);
      await page.keyboard.press('Escape');
      const openGrid = async () => {
        await openView();
        await item('Grid and Guides').hover();
        await item('Grid Options...').click();
      };
      const other = await page.context().newPage();
      await other.goto(url);
      await other.locator('.shape-hit').first().waitFor();
      const revisionBeforeGuides = (await page.evaluate(() => window.office.getState())).revision;
      await page.locator('#tab-view').click();
      await page.locator('#show-guides').check();
      await other.locator('.drawing-guide').nth(1).waitFor();
      assert.equal(
        (await page.evaluate(() => window.office.getState())).revision,
        revisionBeforeGuides,
      );
      await other.locator('#tab-view').click();
      await other.locator('#show-guides').uncheck();
      await page.locator('.drawing-guide').waitFor({ state: 'detached' });
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      assert.equal(await page.locator('#show-guides').isChecked(), false);
      await other.close();
      await openGrid();
      await page.locator('#grid-visible').check();
      await page.locator('#grid-dialog button[value=cancel]').click();
      assert.equal(await page.locator('.drawing-grid').count(), 0);
      await openGrid();
      assert.equal(await page.locator('#grid-visible').isChecked(), false);
      await page.locator('#grid-visible').check();
      await page.locator('#grid-snap').check();
      await page.locator('#grid-smart').uncheck();
      await page.locator('#grid-guides').check();
      await page.locator('#grid-spacing').selectOption('360000');
      await page.locator('#grid-dialog button[value=apply]').click();
      await page.waitForFunction(
        () => !document.querySelector('#grid-dialog').open && !window.office.getState().building,
      );
      assert.equal(await page.locator('.drawing-grid').count(), 1);
      assert.deepEqual((await page.evaluate(() => window.office.getState())).editor.gridSpacing, {
        x: 360000,
        y: 360000,
      });
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      assert.equal(await page.locator('.drawing-grid').count(), 1);
      await openGrid();
      assert.equal(await page.locator('#grid-spacing').inputValue(), '360000');
      assert.equal(await page.locator('#grid-snap').isChecked(), true);
      assert.equal(await page.locator('#grid-guides').isChecked(), true);
      await page.locator('#grid-dialog button[value=cancel]').click();
      const beforeDrag = await page.evaluate(() => window.office.getState());
      const hit = await page.locator('.shape-hit').first().boundingBox();
      await page.mouse.move(hit.x + hit.width / 2, hit.y + hit.height / 2);
      await page.mouse.down();
      await page.mouse.move(hit.x + hit.width / 2 + 31, hit.y + hit.height / 2 + 21, { steps: 4 });
      await page.mouse.up();
      await page.waitForFunction(
        (revision) =>
          window.office.getState().revision > revision && !window.office.getState().building,
        beforeDrag.revision,
      );
      const moved = (await page.evaluate(() => window.office.getState())).editor.slides[0].shapes[0]
        .bounds;
      assert.equal(moved.x % 360000, 0);
      assert.equal(moved.y % 360000, 0);
      await page.locator('[data-edit=undo]').click();
      await page.waitForFunction(
        (x) =>
          !window.office.getState().building &&
          window.office.getState().editor.slides[0].shapes[0].bounds.x === x,
        beforeDrag.editor.slides[0].shapes[0].bounds.x,
      );
      await page.locator('[data-edit=undo]').click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.gridSpacing === null,
      );
      assert.equal(
        (await page.evaluate(() => window.office.getState())).editor.guidesVisible,
        false,
      );
      await openGrid();
      await page.locator('#grid-spacing').selectOption('180000');
      await page.locator('#grid-default').click();
      assert.equal(await page.locator('#grid-dialog').evaluate((el) => el.open), true);
      await page.locator('#grid-dialog button[value=cancel]').click();
      assert.equal(await page.evaluate(() => localStorage.getItem('office-grid-defaults')), null);
      await openGrid();
      await page.locator('#grid-spacing').selectOption('180000');
      await page.locator('#grid-snap').check();
      await page.locator('#grid-default').click();
      await page.screenshot({ path: '/tmp/pptx-grid-browser.png' });
      await page.locator('#grid-dialog button[value=apply]').click();
      await page.waitForFunction(() => !document.querySelector('#grid-dialog').open);
      assert.deepEqual(
        await page.evaluate(() => JSON.parse(localStorage.getItem('office-grid-defaults'))),
        { x: 180000, y: 180000, snap: true },
      );
      await page.locator('[data-edit=undo]').click();
      await page.waitForFunction(
        () =>
          !window.office.getState().building &&
          window.office.getState().editor.gridSpacing === null,
      );
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      await openGrid();
      assert.equal(await page.locator('#grid-spacing').inputValue(), '180000');
      assert.equal(await page.locator('#grid-snap').isChecked(), true);
      await page.locator('#grid-snap').uncheck();
      await page.locator('#grid-spacing').selectOption('custom');
      await page.locator('#grid-custom').fill('0.50001');
      await page.locator('#grid-dialog button[value=apply]').click();
      await page.waitForFunction(() => !document.querySelector('#grid-dialog').open);
      assert.equal(
        (await page.evaluate(() => window.office.getState())).editor.gridSpacing.x,
        180004,
      );
      assert.equal((await page.evaluate(() => window.office.getState())).editor.snapToGrid, false);
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      await openGrid();
      assert.equal(await page.locator('#grid-snap').isChecked(), false);
      await page.locator('#grid-dialog button[value=cancel]').click();
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
