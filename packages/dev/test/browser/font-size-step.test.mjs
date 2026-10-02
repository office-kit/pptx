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
  'font size step follows the PowerPoint gallery for selections, caret typing, and objects',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-font-size-step-'));
    let preview;
    let browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={9} height={2} paragraphs={[{runs:[{text:'a',format:{size:10}},{text:'b',format:{size:13}},{text:'c',format:{size:44}},{text:'d',format:{size:96}},{text:'e',format:{size:98}},{text:'f',format:{size:100}},{text:'g',format:{size:7.5}},{text:'h',format:{size:7}},{text:'i',format:{size:20}},{text:'j',format:{size:44}}]}]} /><Text x={1} y={4} width={4} height={1} size={13}>Object</Text><Text x={1} y={6} width={4} height={1} size={10}>Caret</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const readShapes = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0]).map((shape) =>
          getShapeParagraphElements(shape, 0)
            .filter((element) => element.kind === 'r')
            .map((element) => ({ text: element.text, size: element.format?.size })),
        );
      };
      const input = editor.locator('.inline-edit');
      const selectShapeRange = async (start, end) => {
        await input.focus();
        await input.evaluate(
          (node, range) => {
            window.selectEditorText(node, range[0], range[1]);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          },
          [start, end],
        );
      };
      const increase = editor.getByRole('button', { name: 'Increase Font Size', exact: true });
      const decrease = editor.getByRole('button', { name: 'Decrease Font Size', exact: true });
      await saved();
      let shapes = await readShapes();
      assert.equal(shapes[1][0].size, 13);
      assert.equal(shapes[2][0].size, 10);
      await editor.locator('.hit').first().dblclick();
      await input.waitFor();

      const step = async (index, direction, expected) => {
        await selectShapeRange(index, index + 1);
        await (direction > 0 ? increase : decrease).click();
        await saved();
        const shapes = await readShapes();
        assert.equal(shapes[0][index].size, expected);
      };
      await step(0, 1, 10.5);
      await step(1, 1, 14);
      await step(1, -1, 12);
      await step(2, -1, 40);
      await step(3, 1, 115);
      await step(4, 1, 118);
      await step(5, 1, 120);
      await step(5, -1, 100);
      await step(5, -1, 96);
      await step(6, -1, 7);
      await step(7, -1, 6);

      await selectShapeRange(8, 10);
      await increase.click();
      await saved();
      shapes = await readShapes();
      assert.deepEqual(
        shapes[0].slice(8, 10).map((run) => run.size),
        [24, 48],
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      shapes = await readShapes();
      assert.deepEqual(
        shapes[0].slice(8, 10).map((run) => run.size),
        [20, 44],
      );
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      shapes = await readShapes();
      assert.deepEqual(
        shapes[0].slice(8, 10).map((run) => run.size),
        [24, 48],
      );

      await editor.locator('.hit').nth(1).click();
      await input.waitFor();
      await input.press('Escape');
      await input.waitFor({ state: 'detached' });
      await increase.click();
      await saved();
      shapes = await readShapes();
      assert.equal(shapes[1][0].size, 14);

      await editor.locator('.hit').nth(2).dblclick();
      await input.waitFor();
      await selectShapeRange(1, 1);
      await increase.click();
      await saved();
      await input.waitFor();
      await increase.click();
      await saved();
      await input.press('x');
      await input.press('Escape');
      await saved();
      shapes = await readShapes();
      assert.equal(shapes[2].map((run) => run.text).join(''), 'Cxaret');
      assert.equal(shapes[2].find((run) => run.text === 'x').size, 11);

      await page.reload();
      await saved();
      shapes = await readShapes();
      assert.equal(shapes[0][0].size, 10.5);
      assert.deepEqual(
        shapes[0].slice(8, 10).map((run) => run.size),
        [24, 48],
      );
      assert.equal(shapes[1][0].size, 14);
      assert.equal(shapes[2].find((run) => run.text === 'x').size, 11);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
