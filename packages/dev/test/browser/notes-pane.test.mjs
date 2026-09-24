import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlideNotes, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'notes pane commits drafts to their slide across selection and view changes',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-notes-pane-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide notes="First"/><Slide notes="Second"/></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const state = () => waitForState(preview.url, () => true);
      const notes = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        ).map(getSlideNotes);
      const revision = (await state()).revision;
      await editor.getByRole('button', { name: 'Notes', exact: true }).click();
      const input = editor.getByRole('textbox', { name: 'Notes content', exact: true });
      assert.equal(await input.inputValue(), 'First');
      assert.equal(await input.evaluate((el) => el === document.activeElement), true);
      const resize = editor.getByRole('separator', { name: 'Notes pane height' });
      const height = Number(await resize.getAttribute('aria-valuenow'));
      await resize.focus();
      await resize.press('ArrowUp');
      assert.equal(Number(await resize.getAttribute('aria-valuenow')), height + 10);
      assert.equal((await state()).revision, revision);
      await input.fill('First edited\n日本語');
      await editor.locator('.nav [data-slide-index="1"]').click();
      assert.equal(await input.inputValue(), 'Second');
      assert.equal(await input.evaluate((el) => el === document.activeElement), false);
      await waitForState(preview.url, (s) => s.revision !== revision);
      assert.deepEqual(await notes(), ['First edited\n日本語', 'Second']);
      const secondRevision = (await state()).revision;
      await input.fill('Second edited');
      await input.press('Meta+2');
      await editor.locator('.nav.sorter').waitFor();
      await waitForState(preview.url, (s) => s.revision !== secondRevision);
      assert.deepEqual(await notes(), ['First edited\n日本語', 'Second edited']);
      await editor.getByRole('button', { name: 'Normal', exact: true }).click();
      assert.equal(await input.inputValue(), 'Second edited');
      await input.fill('Pending draft');
      await input.press('Meta+z');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await input.inputValue(), 'Second edited');
      await input.press('Meta+Shift+z');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal(await input.inputValue(), 'Pending draft');
      const hideRevision = (await state()).revision;
      await input.fill('Saved on hide');
      await editor.getByRole('button', { name: 'Notes', exact: true }).click();
      await input.waitFor({ state: 'detached' });
      await waitForState(preview.url, (s) => s.revision !== hideRevision);
      assert.deepEqual(await notes(), ['First edited\n日本語', 'Saved on hide']);
      await editor.getByRole('button', { name: 'Notes', exact: true }).click();
      const idleRevision = (await state()).revision;
      await input.fill('Saved while focused');
      await waitForState(preview.url, (s) => s.revision !== idleRevision);
      assert.deepEqual(await notes(), ['First edited\n日本語', 'Saved while focused']);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
