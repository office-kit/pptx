import assert from 'node:assert/strict';
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeParagraphElements,
  getShapeText,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { installRichTextSelection } from '../helpers/rich-text.mjs';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'outline Font dialog formats only the selected title text and survives save and undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-font-dialog-outline-'));
    let preview;
    let browser;
    try {
      await copyFile(
        new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
        join(dir, 'template.pptx'),
      );
      await writeFile(
        join(dir, 'deck.tsx'),
        `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}}>Heading</Fill><Fill target={{placeholder:{idx:1}}}>Body</Fill></Slide></Presentation>;`,
      );
      preview = await startPreview(join(dir, 'deck.tsx'));
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByRole('tab', { name: 'View', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'View', exact: true })
        .getByRole('button', { name: 'Outline View', exact: true })
        .click();
      await editor.getByRole('tab', { name: 'Home', exact: true }).click();
      const outline = editor.getByRole('navigation', { name: 'Outline View', exact: true });
      const title = outline.getByRole('textbox').first();
      const saved = editor.getByText('Saved to this project', { exact: true });
      const readTitleRuns = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
        );
        const titleShape = getSlideShapes(getSlides(pres)[0]).find(
          (shape) => getShapeText(shape) === 'Heading',
        );
        return getShapeParagraphElements(titleShape, 0)
          .filter((run) => run.kind === 'r')
          .map((run) => ({ text: run.text, size: run.format?.size }));
      };

      await saved.waitFor();
      const before = await readTitleRuns();
      const revision = (await waitForState(preview.url, () => true)).revision;
      await title.focus();
      await title.evaluate((input) => {
        input.focus({ preventScroll: true });
        window.selectEditorText(input, 1, 4);
        input.dispatchEvent(new Event('select', { bubbles: true }));
      });
      await page.keyboard.press('Control+T');
      const dialog = editor.getByRole('dialog', { name: 'Font', exact: true });
      await dialog.waitFor();
      await dialog.getByLabel('Font size', { exact: true }).fill('28');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await waitForState(preview.url, (state) => state.revision !== revision);
      await saved.waitFor();

      assert.deepEqual(await readTitleRuns(), [
        { text: 'H', size: before[0]?.size },
        { text: 'ead', size: 28 },
        { text: 'ing', size: before.at(-1)?.size },
      ]);
      assert.deepEqual((await readTitleRuns()).map((run) => run.text).join(''), 'Heading');

      const undoRevision = (await waitForState(preview.url, () => true)).revision;
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await waitForState(preview.url, (state) => state.revision !== undoRevision);
      await saved.waitFor();
      assert.deepEqual(await readTitleRuns(), before);

      await page.reload();
      await saved.waitFor();
      assert.deepEqual(await readTitleRuns(), before);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
