import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  cm,
  getShapeText,
  getParagraphPropertiesEffective,
  getShapeParagraphCount,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

test(
  'Home line spacing preserves paragraph selection, pending text and atomic dialog undo',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-line-spacing-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={5} height={3} paragraphs={[{runs:[{text:'First'}]},{runs:[{text:'Second'}]}]} /><Text x={7} y={1} width={2} height={2}>Untouched</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0]).map((shape) => ({
          text: getShapeText(shape),
          paragraphs: Array.from({ length: getShapeParagraphCount(shape) }, (_, index) =>
            getParagraphPropertiesEffective(pres, shape, index),
          ),
        }));
      };
      const trigger = editor
        .locator('.ribbon')
        .getByRole('button', { name: 'Line spacing', exact: true });
      const menu = editor.getByRole('menu', { name: 'Line spacing', exact: true });
      const dialog = editor.getByRole('dialog', { name: 'Paragraph', exact: true });
      const options = async () => {
        await trigger.click();
        await menu.getByRole('menuitem', { name: 'Line Spacing Options...', exact: true }).click();
        await dialog.waitFor();
      };
      await saved();
      const original = await read();
      await editor.locator('.hit').first().click();
      await trigger.click();
      assert.deepEqual(await menu.getByRole('menuitemradio').allTextContents(), [
        '✓1.0',
        '1.5',
        '2.0',
        '2.5',
        '3.0',
      ]);
      await menu.getByRole('menuitemradio', { name: '2.0', exact: true }).click();
      await saved();
      let data = await read();
      assert.deepEqual(
        data[0].paragraphs.map((p) => p.lineSpacing),
        [
          { kind: 'pct', value: 2 },
          { kind: 'pct', value: 2 },
        ],
      );
      assert.deepEqual(data[1], original[1]);
      await trigger.click();
      assert.equal(
        await menu
          .getByRole('menuitemradio', { name: '2.0', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await menu.press('Escape');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), original);

      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.inline-edit');
      await input.evaluate((node) => {
        window.selectEditorText(node, 12, 12);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      await page.keyboard.insertText(' added');
      await options();
      await dialog.getByLabel('Before:', { exact: true }).fill('12');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.deepEqual(await read(), original);
      assert.equal(await input.textContent(), 'First\nSecond added');
      await trigger.click();
      await menu.getByRole('menuitemradio', { name: '1.5', exact: true }).click();
      await saved();
      data = await read();
      assert.equal(data[0].text, 'First\nSecond added');
      assert.deepEqual(data[0].paragraphs[0], original[0].paragraphs[0]);
      assert.deepEqual(data[0].paragraphs[1].lineSpacing, { kind: 'pct', value: 1.5 });
      const beforeDialog = data;
      await options();
      await dialog.getByLabel('Alignment:', { exact: true }).selectOption('right');
      await dialog.getByLabel('Before text:', { exact: true }).fill('1');
      await dialog.getByLabel('Special:', { exact: true }).selectOption('hanging');
      await dialog.getByLabel('By:', { exact: true }).fill('0.5');
      await dialog.getByLabel('Before:', { exact: true }).fill('12');
      await dialog.getByLabel('After:', { exact: true }).fill('6');
      await dialog.getByLabel('Line spacing:', { exact: true }).selectOption('exact');
      await dialog.getByLabel('At:', { exact: true }).fill('24');
      await dialog.getByRole('tab', { name: 'Line Breaks and Alignment', exact: true }).click();
      await dialog.getByLabel('Use Asian typography rules', { exact: true }).uncheck();
      await dialog
        .getByLabel('Allow Latin text to wrap in the middle of a word', { exact: true })
        .check();
      await dialog.getByLabel('Allow hanging punctuation', { exact: true }).uncheck();
      await dialog.getByLabel('Text Alignment:', { exact: true }).selectOption('baseline');
      await dialog
        .getByRole('tab', { name: 'Line Breaks and Alignment', exact: true })
        .press('ArrowLeft');
      assert.equal(await dialog.getByLabel('At:', { exact: true }).inputValue(), '24');

      await page.screenshot({ path: '/tmp/pptx-paragraph-dialog.png' });
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      data = await read();
      assert.deepEqual(data[0].paragraphs[0], beforeDialog[0].paragraphs[0]);
      assert.deepEqual(data[1], original[1]);
      const paragraph = data[0].paragraphs[1];
      assert.equal(paragraph.align, 'right');
      assert.equal(paragraph.marL, cm(1));
      assert.equal(paragraph.indent, -cm(0.5));
      assert.equal(paragraph.spcBefPts, 12);
      assert.equal(paragraph.spcAftPts, 6);
      assert.equal(paragraph.asianLineBreak, false);
      assert.equal(paragraph.latinLineBreak, true);
      assert.equal(paragraph.hangingPunctuation, false);
      assert.equal(paragraph.fontAlignment, 'baseline');
      assert.deepEqual(paragraph.lineSpacing, { kind: 'pts', value: 24 });
      await input.press('Control+z');
      await saved();
      assert.deepEqual(await read(), beforeDialog);
      await input.press('Control+y');
      await saved();
      assert.deepEqual(await read(), data);
      await page.reload();
      await saved();
      assert.deepEqual(await read(), data);
      await editor.locator('.hit').first().click();
      await options();
      assert.equal(await dialog.getByLabel('Line spacing:', { exact: true }).inputValue(), '');
      await dialog.getByRole('tab', { name: 'Line Breaks and Alignment', exact: true }).click();
      assert.equal(
        await dialog
          .getByLabel('Allow Latin text to wrap in the middle of a word', { exact: true })
          .evaluate((node) => node.indeterminate),
        true,
      );
      await dialog.getByRole('tab', { name: 'Indents and Spacing', exact: true }).click();
      await dialog.getByLabel('Before:', { exact: true }).fill('24');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      const mixed = await read();
      assert.deepEqual(
        mixed[0].paragraphs.map((p) => p.lineSpacing),
        data[0].paragraphs.map((p) => p.lineSpacing),
      );
      assert.deepEqual(
        mixed[0].paragraphs.map((p) => p.indent),
        data[0].paragraphs.map((p) => p.indent),
      );
      assert.deepEqual(
        mixed[0].paragraphs.map((p) => p.spcBefPts),
        [24, 24],
      );
      assert.deepEqual(
        mixed[0].paragraphs.map((p) => [
          p.asianLineBreak,
          p.latinLineBreak,
          p.hangingPunctuation,
          p.fontAlignment,
        ]),
        data[0].paragraphs.map((p) => [
          p.asianLineBreak,
          p.latinLineBreak,
          p.hangingPunctuation,
          p.fontAlignment,
        ]),
      );
      assert.deepEqual(mixed[1], data[1]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), data);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
