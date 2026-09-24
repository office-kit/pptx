import { chromium } from 'playwright';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test(
  'page setup scales content, cancels, persists and restores through history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-size-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text size={20} x={1} y={2} width={1} height={1}>Left</Text><Text size={20} x={5} y={2} width={1} height={1}>Right</Text><Text size={20} x={3.5} y={2} width={1} height={1}>Move</Text></Slide></Presentation>`,
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
      const before = await state();
      const settled = () =>
        page.waitForFunction(
          () =>
            !window.office.getState().building &&
            !document.querySelector('[data-edit="undo"]').disabled,
        );
      const setup = async () => {
        await page.getByRole('tab', { name: 'Design', exact: true }).click();
        await page.locator('[data-edit="slide-size"]').click();
        await page.getByRole('menuitem', { name: 'Page Setup…', exact: true }).click();
        return page.getByRole('dialog', { name: 'Page Setup', exact: true });
      };
      const fileMenu = page
        .getByRole('menubar')
        .getByRole('menuitem', { name: 'File', exact: true });
      await fileMenu.focus();
      await page.keyboard.press('ArrowDown');
      assert.equal(await fileMenu.getAttribute('aria-expanded'), 'true');
      await page.keyboard.press('Escape');
      assert.equal(await fileMenu.getAttribute('aria-expanded'), 'false');
      assert.equal(await fileMenu.evaluate((element) => element === document.activeElement), true);
      await fileMenu.click();
      await fileMenu.click();
      assert.equal(await page.locator('#office-menu').isHidden(), true);
      await fileMenu.click();
      await page.getByRole('menuitem', { name: 'Page Setup…', exact: true }).click();
      assert.equal(await fileMenu.getAttribute('aria-expanded'), 'false');
      await page.getByRole('dialog', { name: 'Page Setup', exact: true }).waitFor();
      await page.keyboard.press('Escape');
      assert.equal((await state()).revision, before.revision);
      assert.equal(await fileMenu.evaluate((element) => element === document.activeElement), true);
      const editMenu = page
        .getByRole('menubar')
        .getByRole('menuitem', { name: 'Edit', exact: true });
      await fileMenu.click();
      await page.keyboard.press('ArrowRight');
      assert.equal(await fileMenu.getAttribute('aria-expanded'), 'false');
      assert.equal(await editMenu.getAttribute('aria-expanded'), 'true');
      assert.equal(
        await page.getByRole('menuitem', { name: 'Duplicate', exact: true }).isDisabled(),
        true,
      );
      await page.getByRole('menuitem', { name: 'Select All', exact: true }).click();
      await editMenu.click();
      await page.getByRole('menuitem', { name: 'Duplicate', exact: true }).click();
      await settled();
      assert.equal((await state()).editor.slides[0].shapes.length, 6);
      await editMenu.click();
      await page.getByRole('menuitem', { name: 'Undo', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes.length === 3,
      );
      await editMenu.click();
      await page.getByRole('menuitem', { name: 'Redo', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes.length === 6,
      );
      await editMenu.click();
      await page.getByRole('menuitem', { name: 'Undo', exact: true }).click();
      await page.waitForFunction(
        () => window.office.getState().editor.slides[0].shapes.length === 3,
      );
      await page.locator('.shape-hit').first().dblclick();
      const textField = page.getByRole('textbox', { name: 'Edit text', exact: true });
      await textField.evaluate((field) => field.setSelectionRange(1, 3));
      await editMenu.click();
      await page.getByRole('menuitem', { name: 'Delete', exact: true }).click();
      assert.equal(await textField.evaluate((field) => field.value), 'Lt');
      await editMenu.click();
      await page.getByRole('menuitem', { name: 'Undo', exact: true }).click();
      assert.equal(await textField.evaluate((field) => field.value), 'Left');
      assert.deepEqual(
        await textField.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [1, 3],
      );
      await editMenu.click();
      await page.getByRole('menuitem', { name: 'Redo', exact: true }).click();
      assert.equal(await textField.evaluate((field) => field.value), 'Lt');
      await editMenu.click();
      assert.equal(
        await page.getByRole('menuitem', { name: 'Delete', exact: true }).isDisabled(),
        true,
      );
      await page.getByRole('menuitem', { name: 'Undo', exact: true }).click();
      await editMenu.click();
      await page.getByRole('menuitem', { name: 'Select All', exact: true }).click();
      assert.deepEqual(
        await textField.evaluate((field) => [field.selectionStart, field.selectionEnd]),
        [0, 4],
      );
      await page.keyboard.press('Escape');
      before.revision = (await state()).revision;
      let dialog = await setup();
      await page.screenshot({ path: '/tmp/pptx-page-setup.png' });
      await dialog
        .getByRole('group', { name: 'Slides', exact: true })
        .getByLabel('Portrait', { exact: true })
        .check();
      assert.ok(
        Number(await dialog.getByLabel(/Width:/).inputValue()) <
          Number(await dialog.getByLabel(/Height:/).inputValue()),
      );
      await page.keyboard.press('Escape');
      assert.equal((await state()).revision, before.revision);
      dialog = await setup();
      for (const invalid of ['0.5 in', '57 in', '12 cm trailing', 'NaN', '']) {
        await dialog.getByLabel(/Width:/).fill(invalid);
        await dialog.getByRole('button', { name: 'OK', exact: true }).click();
        assert.equal(await dialog.isVisible(), true);
        assert.equal((await state()).revision, before.revision);
        assert.equal(
          await dialog.getByLabel(/Width:/).evaluate((input) => input.checkValidity()),
          false,
        );
      }
      await dialog.getByLabel(/Width:/).fill('16.933333 cm');
      await dialog.getByLabel(/Height:/).fill('3.75 IN');
      await dialog.getByLabel(/Width:/).focus();
      assert.equal(await dialog.getByLabel(/Height:/).inputValue(), '9.525');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await page
        .getByRole('dialog', { name: 'Scale Slide Content' })
        .getByRole('button', { name: 'Scale', exact: true })
        .click();
      await settled();
      await page.waitForFunction(() => window.office.getState().editor.width === 6096000);
      const after = (await state()).editor;
      assert.equal(after.height, before.editor.height / 2);
      for (const [index, shape] of after.slides[0].shapes.entries()) {
        const original = before.editor.slides[0].shapes[index];
        assert.deepEqual(
          shape.bounds,
          Object.fromEntries(
            Object.entries(original.bounds).map(([key, value]) => [key, value / 2]),
          ),
        );
        assert.equal(shape.format.size, original.format.size / 2);
      }
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      assert.deepEqual((await state()).editor, after);
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.width === 12192000 && !window.office.getState().building,
      );
      assert.deepEqual((await state()).editor, before.editor);
      await page.locator('[data-edit="redo"]').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.width === 6096000 && !window.office.getState().building,
      );
      assert.deepEqual((await state()).editor, after);
      dialog = await setup();
      await dialog.getByLabel(/Width:/).fill('25.4');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await page
        .getByRole('dialog', { name: 'Scale Slide Content' })
        .getByRole('button', { name: "Don't Scale", exact: true })
        .click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.width === 9144000 && !window.office.getState().building,
      );
      assert.deepEqual((await state()).editor.slides, after.slides);
      dialog = await setup();
      await dialog
        .getByLabel('Slides sized for:')
        .selectOption({ label: 'Letter Paper (8.5x11 in)' });
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await page
        .getByRole('dialog', { name: 'Scale Slide Content' })
        .getByRole('button', { name: "Don't Scale", exact: true })
        .click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.sizeType === 'letter' &&
          !window.office.getState().building,
      );
      const letter = (await state()).editor;
      dialog = await setup();
      assert.equal(
        await dialog.getByLabel('Slides sized for:').locator('option:checked').textContent(),
        'Letter Paper (8.5x11 in)',
      );
      await dialog.getByLabel('Slides sized for:').selectOption({ label: 'Overhead' });
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.sizeType === 'overhead' &&
          !window.office.getState().building,
      );
      await page
        .getByRole('dialog', { name: 'Page Setup', exact: true })
        .waitFor({ state: 'hidden' });
      assert.equal(await page.getByRole('dialog').count(), 0);
      assert.deepEqual((await state()).editor.slides, letter.slides);
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      dialog = await setup();
      assert.equal(
        await dialog.getByLabel('Slides sized for:').locator('option:checked').textContent(),
        'Overhead',
      );
      await dialog
        .getByRole('group', { name: 'Slides', exact: true })
        .getByLabel('Portrait', { exact: true })
        .check();
      assert.equal(
        await dialog.getByLabel('Slides sized for:').locator('option:checked').textContent(),
        'Overhead',
      );
      await page.keyboard.press('Escape');
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.sizeType === 'letter' &&
          !window.office.getState().building,
      );
      assert.deepEqual((await state()).editor, letter);
      dialog = await setup();
      await dialog.getByLabel('Number slides from:').fill('10');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.firstSlideNumber === 10 &&
          !window.office.getState().building,
      );
      await dialog.waitFor({ state: 'hidden' });
      assert.deepEqual((await state()).editor.slides, letter.slides);
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      dialog = await setup();
      assert.equal(await dialog.getByLabel('Number slides from:').inputValue(), '10');
      await page.keyboard.press('Escape');
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.firstSlideNumber === 1 &&
          !window.office.getState().building,
      );
      assert.deepEqual((await state()).editor, letter);
      dialog = await setup();
      const notes = dialog.getByRole('group', { name: 'Notes, handouts & outline', exact: true });
      assert.equal(await notes.getByLabel('Portrait', { exact: true }).isChecked(), true);
      await notes.getByLabel('Landscape', { exact: true }).check();
      assert.equal(
        await dialog
          .getByRole('group', { name: 'Slides', exact: true })
          .getByLabel('Landscape', { exact: true })
          .isChecked(),
        true,
      );
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.notesOrientation === 'landscape' &&
          !window.office.getState().building,
      );
      await dialog.waitFor({ state: 'hidden' });
      assert.deepEqual((await state()).editor.slides, letter.slides);
      assert.equal((await state()).editor.width, letter.width);
      assert.equal((await state()).editor.height, letter.height);
      await page.reload();
      await page.locator('.shape-hit').first().waitFor();
      dialog = await setup();
      assert.equal(
        await dialog
          .getByRole('group', { name: 'Notes, handouts & outline', exact: true })
          .getByLabel('Landscape', { exact: true })
          .isChecked(),
        true,
      );
      await page.keyboard.press('Escape');
      await page.locator('[data-edit="undo"]').click();
      await page.waitForFunction(
        () =>
          window.office.getState().editor.notesOrientation === 'portrait' &&
          !window.office.getState().building,
      );
      assert.deepEqual((await state()).editor, letter);
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
