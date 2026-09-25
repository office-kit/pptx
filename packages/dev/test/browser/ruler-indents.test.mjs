import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlides, getSlideShapes, getParagraphIndent, loadPresentation } from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

test(
  'ruler indentation preserves selected paragraphs, saves and undoes one gesture',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-ruler-indents-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={3} paragraphs={[{runs:[{text:'First paragraph'}]},{runs:[{text:'Second paragraph'}]}]} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      await saved();
      await editor.getByRole('tab', { name: 'View', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'View', exact: true })
        .getByRole('checkbox', { name: 'Ruler', exact: true })
        .check();
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.inline-edit');
      await input.evaluate((node) => {
        window.selectEditorText(node, 17, 22);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      const readIndents = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(pres)[0])[0];
        return [getParagraphIndent(shape, 0), getParagraphIndent(shape, 1)];
      };
      const initial = await readIndents();
      await editor.locator('.rulers').evaluate((node) => {
        const input = node.parentElement.querySelector('.inline-edit');
        const rect = input.getBoundingClientRect();
        const style = getComputedStyle(input);
        const zero = node.querySelector('.horizontal line[data-value="0"]');
        const origin = node.getBoundingClientRect().left;
        const expected =
          rect.left + (parseFloat(style.paddingLeft) * rect.width) / input.offsetWidth;
        if (Math.abs(origin + Number(zero.getAttribute('x1')) - expected) > 1)
          throw new Error('Ruler zero does not follow the text body');
      });
      const revision = async () => (await waitForState(preview.url, () => true)).revision;
      const drag = async (name, pixels, cancel = false) => {
        const before = await revision();
        const handle = await editor.getByRole('button', { name, exact: true }).boundingBox();
        await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2);
        await page.mouse.down();
        await page.mouse.move(handle.x + handle.width / 2 + pixels, handle.y + handle.height / 2, {
          steps: 5,
        });
        if (cancel) await input.press('Escape');
        await page.mouse.up();
        if (!cancel) await waitForState(preview.url, (state) => state.revision > before);
        await saved();
      };
      await drag('First line indent', 35);
      let values = await readIndents();
      assert.deepEqual(values[0], initial[0]);
      assert.ok(values[1].firstLineEmu > 0);
      const first = values[1].firstLineEmu;
      assert.equal(
        await input.evaluate((node) => node.ownerDocument.getSelection().toString()),
        'econd',
      );
      await drag('Hanging indent', 20);
      values = await readIndents();
      assert.ok(values[1].leftEmu > 0);
      assert.equal(values[1].leftEmu + values[1].firstLineEmu, first);
      const hanging = values[1];
      await drag('Left indent', 20);
      values = await readIndents();
      assert.ok(values[1].leftEmu > hanging.leftEmu);
      assert.equal(values[1].firstLineEmu, hanging.firstLineEmu);
      const beforeCancel = await revision();
      const beforeCancelValues = values;
      await drag('First line indent', 50, true);
      assert.equal(await revision(), beforeCancel);
      assert.deepEqual(await readIndents(), beforeCancelValues);
      await page.screenshot({ path: '/tmp/pptx-ruler-indents.png' });
      const beforeUndo = await revision();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await waitForState(preview.url, (state) => state.revision > beforeUndo);
      await saved();
      assert.deepEqual((await readIndents())[1], hanging);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
