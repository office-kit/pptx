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

test(
  'Mac font size shortcuts change all selected shapes in one undo step',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-font-multiselect-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={1} size={20}>First</Text><Text x={1} y={3} width={3} height={1} size={44}>Second</Text><Text x={1} y={5} width={3} height={1} size={13}>Unselected</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const sizes = async () => {
        await saved();
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0]).map(
          (shape) =>
            getShapeParagraphElements(shape, 0).find((element) => element.kind === 'r')?.format
              ?.size,
        );
      };
      const selectTwo = async () => {
        await page.keyboard.press('Escape');
        await editor
          .locator('.hit')
          .nth(0)
          .click({ position: { x: 20, y: 2 } });
        await editor
          .locator('.hit')
          .nth(1)
          .click({ position: { x: 20, y: 2 }, modifiers: ['Shift'] });
      };
      assert.deepEqual(await sizes(), [20, 44, 13]);
      await selectTwo();
      await page.keyboard.press('Meta+Shift+Period');
      assert.deepEqual(await sizes(), [24, 48, 13]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      assert.deepEqual(await sizes(), [20, 44, 13]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      assert.deepEqual(await sizes(), [24, 48, 13]);
      await selectTwo();
      await page.keyboard.press('Meta+Shift+Comma');
      assert.deepEqual(await sizes(), [20, 44, 13]);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
