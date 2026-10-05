import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { strFromU8, unzipSync } from 'fflate';
import {
  getShapeParagraphElements,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

test(
  'Font dialog edits underline color without losing the selected run formatting',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-font-dialog-underline-color-'));
    let preview;
    let browser;
    try {
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={2} paragraphs={[{runs:[{text:'Before '},{text:'Target',format:{font:'Arial',fontEastAsian:'Noto Sans CJK JP',fontComplexScript:'Noto Naskh Arabic',bold:true,italic:true,strike:true,underline:'sng',color:'#154687'}},{text:' After'}]}]} /></Slide></Presentation>`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
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
            color: run.format?.color,
            underline: run.format?.underline ?? false,
            underlineColor: run.format?.underlineColor,
            bold: run.format?.bold ?? false,
            italic: run.format?.italic ?? false,
            strike: run.format?.strike ?? false,
          }));
      };
      const readSlideXml = async () => {
        const bytes = new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer());
        return strFromU8(unzipSync(bytes)['ppt/slides/slide1.xml']);
      };

      await saved();
      await editor.locator('.hit').first().dblclick();
      await input.waitFor();
      await select(7, 13);
      await page.keyboard.press('Control+T');
      const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      await dialog.waitFor();

      const underlineStyle = dialog.getByLabel('Underline style');
      const underlineColor = dialog.getByRole('button', { name: 'Underline color', exact: true });
      await underlineStyle.selectOption('none');
      await assert.doesNotReject(async () => assert.equal(await underlineColor.isDisabled(), true));
      await underlineStyle.selectOption('sng');
      assert.equal(await underlineColor.isDisabled(), false);

      await underlineColor.click();
      const colorMenu = editor.getByRole('menu', { name: 'Underline color', exact: true });
      await colorMenu.waitFor();
      const automatic = colorMenu.getByRole('menuitemradio', { name: 'Automatic', exact: true });
      assert.equal(await automatic.getAttribute('aria-checked'), 'true');
      await colorMenu.getByRole('menuitemradio', { name: 'Red', exact: true }).click();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();

      assert.deepEqual(
        (await readRuns()).map(({ text, ...run }) => ({ text, ...run })),
        [
          {
            text: 'Before ',
            font: undefined,
            eastAsian: undefined,
            complexScript: undefined,
            color: undefined,
            underline: false,
            underlineColor: undefined,
            bold: false,
            italic: false,
            strike: false,
          },
          {
            text: 'Target',
            font: 'Arial',
            eastAsian: 'Noto Sans CJK JP',
            complexScript: 'Noto Naskh Arabic',
            color: '#154687',
            underline: true,
            underlineColor: '#FF0000',
            bold: true,
            italic: true,
            strike: true,
          },
          {
            text: ' After',
            font: undefined,
            eastAsian: undefined,
            complexScript: undefined,
            color: undefined,
            underline: false,
            underlineColor: undefined,
            bold: false,
            italic: false,
            strike: false,
          },
        ],
      );
      const redXml = await readSlideXml();
      assert.match(
        redXml,
        /<a:rPr[^>]*u="sng"[^>]*>.*?<a:uFill>.*?<a:solidFill>.*?<a:srgbClr val="FF0000"/s,
      );

      await select(7, 13);
      await page.keyboard.press('Control+T');
      await dialog.waitFor();
      await dialog.getByRole('button', { name: 'Underline color', exact: true }).click();
      const automaticMenu = editor.getByRole('menu', { name: 'Underline color', exact: true });
      await automaticMenu.getByRole('menuitemradio', { name: 'Automatic', exact: true }).click();
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      assert.equal((await readRuns())[1].underlineColor, null);
      assert.match(
        await readSlideXml(),
        /<a:rPr[^>]*u="sng"[^>]*>.*?<a:uFillTx\s*\/?>(?:.*?<\/a:rPr>)?/s,
      );

      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      const undone = await readRuns();
      assert.equal(undone[1].underline, true);
      assert.equal(undone[1].underlineColor, '#FF0000');
      assert.equal(undone[1].strike, true);
      assert.equal(undone[1].font, 'Arial');
      assert.match(await readSlideXml(), /uFill>.*?FF0000/s);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      const original = await readRuns();
      assert.equal(original[1].underline, true);
      assert.equal(original[1].underlineColor, undefined);
      assert.equal(original[1].strike, true);
      assert.equal(original[1].font, 'Arial');
      assert.doesNotMatch(await readSlideXml(), /uFill(?:Tx)?/);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
