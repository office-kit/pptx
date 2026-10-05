import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
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
import { startPreview } from '../helpers/server.mjs';

test(
  'canvas Change Case preserves mixed Unicode text, caret words, object text and history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-canvas-case-'));
    let preview;
    let browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={8} height={1} paragraphs={[{runs:[{text:'hello '},{text:'Straße 🌎 world',format:{italic:true}}]}]} /><Text x={1} y={3} width={8} height={1}>whole object</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = editor.getByText('Saved to this project', { exact: true });
      await saved.waitFor();
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0]).map(getShapeText);
      };
      const readRuns = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
        );
        return getShapeParagraphElements(getSlideShapes(getSlides(pres)[0])[0], 0);
      };
      const hits = editor.locator('.hit');
      await hits.nth(0).dblclick();
      const input = editor.locator('.canvas-shell .inline-edit').first();
      await input.waitFor();
      const select = async (start, end = start) => {
        await input.focus();
        await input.evaluate(
          (node, range) => {
            window.selectEditorText(node, ...range);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          },
          [start, end],
        );
      };
      const bar = editor.getByRole('group', { name: 'Selected text formatting', exact: true });

      // Select the mixed-format Unicode run and transform only the selected span.
      await select(6, 15);
      const before = await read();
      const revision = (await (await fetch(`${preview.url}/editor/state`)).json()).revision;
      await bar.getByRole('button', { name: 'Change Case', exact: true }).click();
      await editor.getByRole('menuitem', { name: 'UPPERCASE', exact: true }).click();
      const changed = await waitRevision(preview.url, revision);
      await saved.waitFor();
      assert.deepEqual(await read(), ['hello STRASSE 🌎 world', 'whole object']);
      const changedRuns = await readRuns();
      assert.equal(changedRuns.map((run) => run.text).join(''), 'hello STRASSE 🌎 world');
      assert.equal(changedRuns[1].format.italic, true);
      assert.equal(before[0], 'hello Straße 🌎 world');

      // Undo/redo are one history operation and retain the saved text.
      let historyRevision = changed;
      await page.keyboard.press('Control+z');
      const undoneState = await waitRevision(preview.url, historyRevision);
      await saved.waitFor();
      assert.deepEqual(await read(), ['hello Straße 🌎 world', 'whole object']);
      historyRevision = undoneState;
      await page.keyboard.press('Control+y');
      const redone = await waitRevision(preview.url, historyRevision);
      await saved.waitFor();
      assert.deepEqual(await read(), ['hello STRASSE 🌎 world', 'whole object']);

      // A collapsed caret expands to the current word and keeps the rest intact.
      historyRevision = redone;
      await page.keyboard.press('Control+z');
      await waitRevision(preview.url, historyRevision);
      await select(2, 2);
      const caretRevision = (await (await fetch(`${preview.url}/editor/state`)).json()).revision;
      await bar.getByRole('button', { name: 'Change Case', exact: true }).click();
      await editor.getByRole('menuitem', { name: 'UPPERCASE', exact: true }).click();
      await waitRevision(preview.url, caretRevision);
      await saved.waitFor();
      assert.deepEqual(await read(), ['HELLO Straße 🌎 world', 'whole object']);
      await input.press('Escape');

      // Selecting the second shape as an object transforms its complete text.
      await hits.nth(1).click();
      await editor.locator('.canvas-shell .inline-edit').press('Escape');
      const ribbon = editor.locator('.ribbon .font-ribbon');
      const objectRevision = (await (await fetch(`${preview.url}/editor/state`)).json()).revision;
      await ribbon.getByRole('button', { name: 'Change Case', exact: true }).click();
      await editor.getByRole('menuitem', { name: 'UPPERCASE', exact: true }).click();
      await waitRevision(preview.url, objectRevision);
      await saved.waitFor();
      assert.deepEqual(await read(), ['HELLO Straße 🌎 world', 'WHOLE OBJECT']);
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.deepEqual(await read(), ['HELLO Straße 🌎 world', 'WHOLE OBJECT']);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

async function waitRevision(url, previous) {
  const deadline = Date.now() + 10000;
  while (Date.now() < deadline) {
    const state = await (await fetch(`${url}/editor/state`)).json();
    if (!state.building && state.revision !== previous) return state.revision;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error('Editor state did not settle');
}
