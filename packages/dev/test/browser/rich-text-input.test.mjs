import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

test(
  'rich text input displays formatting and preserves native multiline edits and history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-rich-input-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={3}>abc</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      const input = editor.locator('.inline-edit');
      const select = async (start, end = start) => {
        await input.focus();
        await input.evaluate(
          (node, range) => window.selectEditorText(node, ...range),
          [start, end],
        );
      };
      await select(1, 2);
      await input.press('Control+b');
      assert.equal(
        await input
          .locator('span')
          .filter({ hasText: /^b$/ })
          .evaluate((n) => getComputedStyle(n).fontWeight),
        '700',
      );
      await select(2);
      await input.press('Enter');
      await page.keyboard.insertText('日本語');
      assert.equal(await input.textContent(), 'ab\n日本語c');
      await input.press('Control+z');
      assert.equal(await input.textContent(), 'ab\nc');
      await input.press('Control+z');
      assert.equal(await input.textContent(), 'abc');
      await input.press('Control+y');
      await input.press('Control+y');
      assert.equal(await input.textContent(), 'ab\n日本語c');
      await select(0, 6);
      await input.press('Backspace');
      assert.equal(await input.textContent(), 'c');
      await input.press('Control+z');
      assert.equal(await input.textContent(), 'ab\n日本語c');
      await select(7);
      await input.press('Enter');
      await input.press('Enter');
      await page.keyboard.insertText('終わり');
      assert.equal(await input.textContent(), 'ab\n日本語c\n\n終わり');
      await input.press('Control+Enter');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      assert.equal(await input.textContent(), 'ab\n日本語c\n\n終わり');
      await page.screenshot({ path: '/tmp/pptx-pr287-rich-text-input.png', fullPage: true });
      await select(0, 12);
      await input.evaluate((node) => {
        node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
        node.dispatchEvent(
          new InputEvent('beforeinput', {
            bubbles: true,
            inputType: 'insertCompositionText',
            isComposing: true,
          }),
        );
        node.textContent = '候';
        window.selectEditorText(node, 1);
        node.dispatchEvent(
          new InputEvent('input', {
            bubbles: true,
            inputType: 'insertCompositionText',
            isComposing: true,
          }),
        );
      });
      assert.equal(
        await input.locator('span').count(),
        0,
        'composition must keep the browser-owned DOM',
      );
      await input.evaluate((node) => {
        node.dispatchEvent(
          new InputEvent('beforeinput', {
            bubbles: true,
            inputType: 'insertCompositionText',
            isComposing: true,
          }),
        );
        node.textContent = '候補';
        window.selectEditorText(node, 2);
        node.dispatchEvent(
          new InputEvent('input', {
            bubbles: true,
            inputType: 'insertCompositionText',
            isComposing: true,
          }),
        );
        node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '候補' }));
      });
      assert.equal(await input.textContent(), '候補');
      await input.press('Control+z');
      assert.equal(await input.textContent(), 'ab\n日本語c\n\n終わり');
      await input.press('Escape');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
