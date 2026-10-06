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
  'Font dialog edits selected runs, preserves mixed formatting, and supports cancel/undo',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-font-dialog-'));
    let preview;
    let browser;
    try {
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={2} paragraphs={[{runs:[{text:'Before '},{text:'Target',format:{font:'Latin Original',fontEastAsian:'Asian Original',fontComplexScript:'Complex Original',bold:true,italic:true,kern:1200,spc:15}},{text:' After'}]}]} /></Slide></Presentation>`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const input = editor.locator('.canvas-shell .inline-edit');
      const select = (start, end) =>
        input.evaluate(
          (node, range) => {
            node.focus({ preventScroll: true });
            window.selectEditorText(node, range[0], range[1]);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          },
          [start, end],
        );
      const readRuns = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(pres)[0])[0];
        return getShapeParagraphElements(shape, 0)
          .filter((run) => run.kind === 'r')
          .map((run) => ({
            text: run.text,
            font: run.format?.font,
            eastAsian: run.format?.fontEastAsian,
            complexScript: run.format?.fontComplexScript,
            underline: run.format?.underline ?? false,
            size: run.format?.size,
            spc: run.format?.spc ?? 0,
            bold: run.format?.bold ?? false,
            italic: run.format?.italic ?? false,
            kern: run.format?.kern ?? 0,
          }));
      };
      await saved();
      await editor.locator('.hit').first().dblclick();
      await input.waitFor();
      await select(7, 13);
      await page.keyboard.press('Control+T');
      const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      await dialog.waitFor();
      await dialog.getByRole('button', { name: 'Font options', exact: true }).first().click();
      const fontMenu = editor.getByRole('menu', { name: 'Latin text font', exact: true });
      await fontMenu.waitFor();
      await fontMenu.getByRole('menuitemradio', { name: 'Arial', exact: true }).click();
      await dialog.getByLabel('Font size').fill('22');
      await dialog.getByRole('tab', { name: 'Character Spacing', exact: true }).click();
      assert.equal(await dialog.getByLabel('By (pt)').inputValue(), '0.2');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      assert.deepEqual(await readRuns(), [
        {
          text: 'Before ',
          font: undefined,
          eastAsian: undefined,
          complexScript: undefined,
          underline: false,
          size: undefined,
          spc: 0,
          bold: false,
          italic: false,
          kern: 0,
        },
        {
          text: 'Target',
          font: 'Arial',
          eastAsian: 'Asian Original',
          complexScript: 'Complex Original',
          underline: false,
          size: 22,
          spc: 15,
          bold: true,
          italic: true,
          kern: 1200,
        },
        {
          text: ' After',
          font: undefined,
          eastAsian: undefined,
          complexScript: undefined,
          underline: false,
          size: undefined,
          spc: 0,
          bold: false,
          italic: false,
          kern: 0,
        },
      ]);
      await select(7, 13);
      await page.keyboard.press('Control+T');
      await dialog.waitFor();
      await dialog.getByLabel('Underline style').selectOption('wavyHeavy');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      assert.equal((await readRuns())[1].underline, 'wavyHeavy');
      await select(7, 13);
      await page.keyboard.press('Control+T');
      await dialog.waitFor();
      await dialog.getByRole('tab', { name: 'Character Spacing', exact: true }).click();
      assert.equal(await dialog.getByLabel('By (pt)').inputValue(), '0.2');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      await saved();
      await select(0, 19);
      await page.keyboard.press('Control+T');
      await dialog.waitFor();
      await dialog.getByRole('tab', { name: 'Character Spacing', exact: true }).click();
      assert.equal(await dialog.getByLabel('By (pt)').getAttribute('step'), '0.1');
      await dialog.getByLabel('Spacing', { exact: true }).selectOption('expanded');
      const spacingAmount = dialog.getByLabel('By (pt)');
      assert.equal(await spacingAmount.inputValue(), '1');
      await spacingAmount.fill('0.14');
      await spacingAmount.press('Tab');
      assert.equal(await spacingAmount.inputValue(), '0.1');
      await spacingAmount.press('ArrowUp');
      assert.equal(await spacingAmount.inputValue(), '0.2');
      await spacingAmount.fill('0.15');
      await spacingAmount.press('Tab');
      assert.equal(await spacingAmount.inputValue(), '0.2');
      await dialog.getByLabel('Use kerning for fonts', { exact: true }).uncheck();
      await dialog.getByRole('tab', { name: 'Font', exact: true }).click();
      await dialog.getByRole('tab', { name: 'Character Spacing', exact: true }).click();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      const spaced = await readRuns();
      assert.equal(spaced[1].spc, 20);
      assert.equal(spaced[1].bold, true);
      assert.equal(spaced[1].italic, true);
      assert.equal(spaced[1].kern, 0);
      assert.equal(spaced[1].eastAsian, 'Asian Original');
      assert.equal(spaced[1].complexScript, 'Complex Original');
      const beforeCancel = await readRuns();
      await select(7, 13);
      await page.keyboard.press('Control+T');
      await dialog.waitFor();
      await dialog.getByLabel('Font size').fill('30');
      await dialog.press('Escape');
      await saved();
      assert.deepEqual(await readRuns(), beforeCancel);
      await page.keyboard.press('Control+T');
      await dialog.waitFor();
      await dialog.getByLabel('Font size').fill('30');
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      await saved();
      assert.deepEqual(await readRuns(), beforeCancel);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await readRuns())[1].spc, 15);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
