import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeParagraphElements,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

test(
  'font dialog requires a percent sign for superscript and subscript offset',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-font-dialog-offset-'));
    let preview;
    let browser;
    try {
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={2}>Target</Text></Slide></Presentation>`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const input = editor.locator('.canvas-shell .inline-edit');
      await saved();
      await editor.locator('.hit').first().dblclick();
      await input.waitFor();
      await input.evaluate((node) => {
        node.focus({ preventScroll: true });
        window.selectEditorText(node, 0, 6);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      await page.keyboard.press('Control+T');
      const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      await dialog.waitFor();
      const offset = dialog.getByLabel('Offset (%)', { exact: true });
      const superscript = dialog.getByLabel('Superscript', { exact: true });
      const subscript = dialog.getByLabel('Subscript', { exact: true });
      const invalidDialog = editor.getByRole('alertdialog', { name: 'Invalid value', exact: true });
      await offset.fill('10');
      await offset.press('Tab');
      await invalidDialog.waitFor();
      assert.equal(await superscript.isChecked(), false);
      assert.equal(await subscript.isChecked(), false);
      assert.match(await invalidDialog.textContent(), /The value “10” is invalid/);
      await invalidDialog.getByRole('button', { name: 'OK', exact: true }).click();
      assert.equal(await offset.inputValue(), '');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await invalidDialog.waitFor();
      await invalidDialog.getByRole('button', { name: 'Discard Change', exact: true }).click();
      assert.equal(await offset.inputValue(), '0%');
      await superscript.check();
      assert.equal(await offset.inputValue(), '30%');
      assert.equal(await superscript.isChecked(), true);
      assert.equal(await subscript.isChecked(), false);
      await offset.fill('10%');
      await offset.press('Tab');
      assert.equal(await invalidDialog.isVisible(), false);
      assert.equal(await superscript.isChecked(), true);
      assert.equal(await subscript.isChecked(), false);
      await offset.fill('+10.5%');
      await offset.press('Tab');
      await invalidDialog.waitFor();
      assert.match(await invalidDialog.textContent(), /The value “\+10.5%” is invalid/);
      await invalidDialog.getByRole('button', { name: 'Discard Change', exact: true }).click();
      assert.equal(await offset.inputValue(), '10%');
      await offset.fill('-25%');
      await offset.press('Tab');
      assert.equal(await invalidDialog.isVisible(), false);
      assert.equal(await superscript.isChecked(), false);
      assert.equal(await subscript.isChecked(), true);
      await offset.fill('10%');
      await offset.press('Tab');
      assert.equal(await superscript.isChecked(), true);
      assert.equal(await subscript.isChecked(), false);
      await offset.fill('10.5%');
      await offset.press('Tab');
      assert.equal(await offset.inputValue(), '10.5%');
      await offset.press('ArrowUp');
      assert.equal(await offset.inputValue(), '10.5%');
      await dialog.getByRole('button', { name: 'Increase offset', exact: true }).click();
      assert.equal(await offset.inputValue(), '12%');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      const pres = await loadPresentation(
        new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
      );
      const shape = getSlideShapes(getSlides(pres)[0])[0];
      const run = getShapeParagraphElements(shape, 0).find((element) => element.kind === 'r');
      assert.equal(run?.format?.baseline, 0.12);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

for (const discardInvalid of [false, true])
  test(
    `font dialog preserves mixed offsets ${discardInvalid ? 'after discarding invalid input' : 'when the offset field is untouched'}`,
    { timeout: 90000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-font-dialog-mixed-offset-'));
      let preview;
      let browser;
      try {
        await writeFile(
          join(dir, 'deck.tsx'),
          `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={2} paragraphs={[{runs:[{text:'Raised',format:{baseline:0.3}},{text:' Lowered',format:{baseline:-0.25}}]}]} /></Slide></Presentation>`,
        );
        preview = await startPreview(join(dir, 'deck.tsx'));
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        await installRichTextSelection(page);
        await page.goto(preview.url);
        const editor = page.frameLocator('#editor-frame');
        const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
        const input = editor.locator('.canvas-shell .inline-edit');
        await saved();
        await editor.locator('.hit').first().dblclick();
        await input.waitFor();
        await input.evaluate((node) => {
          node.focus({ preventScroll: true });
          window.selectEditorText(node, 0, 14);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
        await page.keyboard.press('Control+T');
        const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
        await dialog.waitFor();
        const offset = dialog.getByLabel('Offset (%)', { exact: true });
        const invalidDialog = editor.getByRole('alertdialog', {
          name: 'Invalid value',
          exact: true,
        });
        assert.equal(await offset.inputValue(), '');
        if (discardInvalid) {
          await offset.fill('10');
          await offset.press('Tab');
          await invalidDialog.waitFor();
          await invalidDialog.getByRole('button', { name: 'Discard Change', exact: true }).click();
        }
        assert.equal(await offset.inputValue(), '');
        await dialog.locator('select').first().selectOption('true:false');
        await dialog.getByRole('button', { name: 'OK', exact: true }).click();
        await saved();
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(pres)[0])[0];
        const runs = getShapeParagraphElements(shape, 0).filter((element) => element.kind === 'r');
        assert.deepEqual(
          runs.map((run) => run.format?.bold),
          [true, true],
        );
        assert.deepEqual(
          runs.map((run) => run.format?.baseline),
          [0.3, -0.25],
        );
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
