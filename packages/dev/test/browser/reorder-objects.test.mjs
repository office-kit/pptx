import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  renameShape,
  savePresentation,
  loadPresentation,
  getSlides,
  getSlideShapes,
  getShapeName,
  inches,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'layer preview stages keyboard and drag changes, cancels, commits and undoes saved order',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-reorder-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      for (const name of ['A', 'B', 'C']) {
        renameShape(
          addSlideTextBox(slide, {
            x: inches(2),
            y: inches(2),
            w: inches(4),
            h: inches(1),
            text: name,
          }),
          name,
        );
      }
      const source = join(dir, 'source.pptx');
      const file = join(dir, 'deck.tsx');
      await writeFile(source, await savePresentation(pres));
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const state = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        ).map(getShapeName);
      const arrange = editor.getByRole('button', { name: 'Arrange', exact: true });
      await saved();
      await arrange.click();
      assert.equal(
        await editor
          .getByRole('menuitem', { name: 'Reorder Overlapping Objects', exact: true })
          .isDisabled(),
        true,
      );
      await editor
        .getByRole('menuitemcheckbox', { name: 'Selection Pane...', exact: true })
        .click();
      const pane = editor.getByRole('region', { name: 'Selection Pane', exact: true });
      await pane.getByRole('button', { name: 'C', exact: true }).click();
      await pane.getByRole('button', { name: 'A', exact: true }).click({ modifiers: ['Shift'] });
      async function open() {
        await arrange.click();
        await editor
          .getByRole('menuitem', { name: 'Reorder Overlapping Objects', exact: true })
          .click();
        await editor.getByRole('dialog').waitFor();
      }
      await open();
      let dialog = editor.getByRole('dialog');
      assert.equal(await dialog.locator('img').count(), 3);
      await dialog.screenshot({ path: '/tmp/pptx-reorder-preview.png' });
      await dialog.getByRole('button', { name: 'C at position 1', exact: true }).press('End');
      await dialog.getByRole('button', { name: 'C at position 3', exact: true }).waitFor();
      assert.deepEqual(await state(), ['A', 'B', 'C']);
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).press('Enter');
      await dialog.waitFor({ state: 'detached' });
      assert.deepEqual(await state(), ['A', 'B', 'C']);
      await open();
      dialog = editor.getByRole('dialog');
      const first = dialog.getByRole('button', { name: 'C at position 1', exact: true });
      const box = await first.boundingBox();
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.mouse.down();
      await page.mouse.move(box.x + box.width / 2 + 400, box.y + box.height / 2, { steps: 8 });
      await page.mouse.up();
      const moved = dialog.getByRole('button', { name: /C at position [23]/ });
      await moved.waitFor();
      await moved.press('End');
      await dialog.getByRole('button', { name: 'C at position 3', exact: true }).press('Enter');
      await dialog.waitFor({ state: 'detached' });
      await saved();
      assert.deepEqual(await state(), ['C', 'A', 'B']);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await state(), ['A', 'B', 'C']);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
