import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getParagraphEndFormat,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'object font stepping updates empty paragraphs and retains them through undo and reload',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-empty-font-step-'));
    let preview;
    let browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={4} paragraphs={[{runs:[{text:'Visible',format:{size:20}}]},{runs:[],endFormat:{size:20}},{runs:[],endFormat:{size:44}}]} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const sizes = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(pres)[0])[0];
        return [1, 2].map((index) => getParagraphEndFormat(shape, index)?.size);
      };
      await saved();
      assert.deepEqual(await sizes(), [20, 44]);
      await editor.locator('.hit').first().click();
      const input = editor.locator('.canvas-shell .inline-edit');
      await input.waitFor();
      await input.press('Escape');
      await input.waitFor({ state: 'detached' });
      await editor.getByRole('button', { name: 'Increase Font Size', exact: true }).click();
      await saved();
      assert.deepEqual(await sizes(), [24, 48]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await sizes(), [20, 44]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await sizes(), [24, 48]);
      await page.reload();
      await saved();
      assert.deepEqual(await sizes(), [24, 48]);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
