import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getShapeText, getSlideShapes, getSlides, loadPresentation } from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

test(
  'font shortcut cancel restores inline text focus and selection',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-font-dialog-focus-'));
    let preview;
    let browser;
    try {
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={2} paragraphs={[{runs:[{text:'Before '},{text:'Target',format:{size:18}},{text:' After'}]}]} /></Slide></Presentation>`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const input = editor.locator('.inline-edit');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const selectTarget = () =>
        input.evaluate((node) => {
          node.focus({ preventScroll: true });
          window.selectEditorText(node, 7, 13);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(pres)[0])[0];
        return getShapeText(shape);
      };

      await saved();
      await editor.locator('.hit').first().dblclick();
      await input.waitFor();
      await selectTarget();
      await page.keyboard.press('Control+T');
      const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      await dialog.waitFor();
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      await input.waitFor();
      assert.equal(await input.evaluate((node) => node === node.ownerDocument.activeElement), true);
      const beforeTyping = (await waitForState(preview.url, () => true)).revision;
      await page.keyboard.type('X');
      await editor.locator('.floating-text-format-bar summary').click();
      await editor.getByRole('button', { name: 'Done', exact: true }).click();
      await waitForState(preview.url, (state) => state.revision !== beforeTyping);
      await saved();
      assert.equal(await read(), 'Before X After');
      const beforeUndo = (await waitForState(preview.url, () => true)).revision;
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await waitForState(preview.url, (state) => state.revision !== beforeUndo);
      await saved();
      assert.equal(await read(), 'Before Target After');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
