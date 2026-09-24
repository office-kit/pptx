import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'custom show definitions persist, undo and play their ordered repeated slides',
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
      const keys = await page.evaluate(() =>
        window.office.getState().editor.slides.map((s) => s.key),
      );
      const openManager = async () => {
        await page.locator('#tab-show').click();
        await page.locator('#custom-shows').click();
        await page.getByRole('menuitem', { name: 'Custom Slide Show...' }).click();
      };
      await openManager();
      const manager = page.getByRole('dialog', { name: 'Custom Shows', exact: true });
      await manager.getByRole('button', { name: 'New custom show' }).click();
      const definition = page.getByRole('dialog', { name: 'Define Custom Show' });
      await definition.getByLabel('Slide show name:').fill('Demo');
      for (const key of [keys[1], keys[0], keys[1]]) {
        await definition.getByLabel('Slides in presentation:').selectOption(key);
        await definition.getByRole('button', { name: 'Add →' }).click();
      }
      await definition.getByRole('button', { name: 'Move slide up' }).click();
      await definition.getByRole('button', { name: 'Move slide down' }).click();
      await definition.getByRole('button', { name: 'OK', exact: true }).click();
      await definition.waitFor({ state: 'detached' });
      assert.deepEqual(
        await page.evaluate(() => window.office.getState().editor.customShows[0].slides),
        [keys[1], keys[0], keys[1]],
      );
      await manager.getByRole('button', { name: 'Copy', exact: true }).click();
      await page.waitForFunction(() => window.office.getState().editor.customShows.length === 2);
      await manager.getByRole('button', { name: 'Delete custom show' }).click();
      await page.waitForFunction(() => window.office.getState().editor.customShows.length === 1);
      await manager.getByRole('button', { name: 'Edit...', exact: true }).click();
      await definition.getByLabel('Slide show name:').fill('Cancelled');
      await definition.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(
        await page.evaluate(() => window.office.getState().editor.customShows[0].name),
        'Demo',
      );
      await manager.getByRole('button', { name: 'Start Show' }).click();
      await page.waitForFunction(
        () => window.office.getState().presenting && window.office.getState().index === 1,
      );
      assert.equal(await page.locator('#present-next').isEnabled(), true);
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 0);
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.keyboard.press('ArrowLeft');
      await page.waitForFunction(() => window.office.getState().index === 0);
      await page.keyboard.press('End');
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.keyboard.press('ArrowRight');
      await page.locator('#show-screen.end').waitFor({ state: 'visible' });
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => !window.office.getState().presenting);
      await page.locator('#present-first').click();
      await page.waitForFunction(
        () => window.office.getState().presenting && window.office.getState().index === 0,
      );
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => !window.office.getState().presenting);
      await page.locator('[data-edit="undo"]').first().click();
      await page.waitForFunction(() => window.office.getState().editor.customShows.length === 2);
      await page.reload();
      await page.waitForFunction(() => window.office?.getState().editor?.customShows.length === 2);
      await page.locator('.thumbnail').first().click();
      await page.locator('.shape-hit').first().click();
      await page.locator('#tab-insert').click();
      await page.locator('[data-edit="action-settings"]').click();
      const settings = page.getByRole('dialog', { name: 'Action Settings' });
      await settings.getByRole('radio', { name: 'Hyperlink to:' }).check();
      await settings.getByLabel('Hyperlink to', { exact: true }).selectOption('customShow');
      const destination = page.getByRole('dialog', { name: 'Hyperlink to Custom Show' });
      await destination.getByLabel('Custom show:').selectOption('0');
      await destination.getByLabel('Show and return').check();
      await destination.getByRole('button', { name: 'OK', exact: true }).click();
      await settings.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes[0].link?.action.kind === 'customShow',
      );
      const savedAction = await page.evaluate(
        () => window.office.getState().editor.slides[0].shapes[0].link.action,
      );
      assert.deepEqual(savedAction, { kind: 'customShow', id: 0, showAndReturn: true });
      await page.locator('[data-edit="link"]').click();
      const link = page.getByRole('dialog', { name: 'Edit Hyperlink' });
      assert.equal(
        await link.getByLabel('Select a place in this document:').inputValue(),
        'customShow:0',
      );
      assert.equal(await link.getByLabel('Show and return').isChecked(), true);
      await link.getByRole('button', { name: 'OK', exact: true }).click();
      await link.waitFor({ state: 'detached' });
      await page.locator('#status-present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      const openDetail = () => page.locator('#slide a[href="#slide-customShow-0-return"]').click();
      await openDetail();
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 0);
      // Enter the same show from its second position, then return to that position.
      await openDetail();
      for (const index of [0, 1, 0]) {
        await page.keyboard.press('ArrowRight');
        await page.waitForFunction((index) => window.office.getState().index === index, index);
      }
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 0);
      assert.equal(await page.evaluate(() => window.office.getState().presenting), true);
      assert.equal(await page.locator('#show-screen').isVisible(), false);
      // Normal sequence resumes on the originating slide.
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => window.office.getState().index === 1);
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => !window.office.getState().presenting);
      await page.locator('.thumbnail').first().click();
      await page.locator('.shape-hit').first().click();
      await page.locator('[data-edit="link"]').click();
      await link.getByLabel('Show and return').uncheck();
      await link.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.slides[0].shapes[0].link?.action.showAndReturn === false,
      );
      await page.locator('#status-present').click();
      await page.waitForFunction(() => window.office.getState().presenting);
      await page.locator('#slide a[href="#slide-customShow-0-exit"]').click();
      await page.waitForFunction(() => window.office.getState().index === 1);
      for (const index of [0, 1]) {
        await page.keyboard.press('ArrowRight');
        await page.waitForFunction((index) => window.office.getState().index === index, index);
      }
      await page.keyboard.press('ArrowRight');
      await page.locator('#show-screen.end').waitFor({ state: 'visible' });
      await page.keyboard.press('ArrowRight');
      await page.waitForFunction(() => !window.office.getState().presenting);
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
