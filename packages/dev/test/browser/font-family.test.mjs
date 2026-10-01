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

test('font picker lists all families and applies Arial with undo', { timeout: 60000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-font-family-'));
  let preview;
  let browser;
  try {
    await writeFile(
      join(dir, 'deck.tsx'),
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={6} height={1} font="Calibri">Calibri sample</Text></Slide></Presentation>`,
    );
    preview = await startPreview(join(dir, 'deck.tsx'));
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
    await installRichTextSelection(page);
    await page.goto(preview.url);
    await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
    const editor = page.frameLocator('#editor-frame');
    await editor.getByText('Saved to this project', { exact: true }).waitFor();
    await editor.locator('.hit').first().dblclick();
    const input = editor.locator('.inline-edit');
    await input.evaluate((node) => window.selectEditorText(node, 0, 7));
    await editor.getByRole('tab', { name: 'Home', exact: true }).click();
    await editor.getByRole('button', { name: 'Font options', exact: true }).click();
    const menu = editor.getByRole('menu', { name: 'Font', exact: true });
    assert.equal(await menu.getByRole('menuitemradio', { name: 'Arial', exact: true }).count(), 1);
    await menu.getByRole('menuitemradio', { name: 'Arial', exact: true }).click();
    await editor.getByText('Saved to this project', { exact: true }).waitFor();
    const readFonts = async () => {
      const pres = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      const shape = getSlideShapes(getSlides(pres)[0])[0];
      return getShapeParagraphElements(shape, 0).map((run) => run.format?.font);
    };
    assert.deepEqual(await readFonts(), ['Arial', 'Calibri']);
    await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
    await editor.getByText('Saved to this project', { exact: true }).waitFor();
    assert.deepEqual(await readFonts(), ['Calibri']);
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});
