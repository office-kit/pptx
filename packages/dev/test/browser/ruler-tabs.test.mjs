import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getParagraphPropertiesEffective,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

test(
  'ruler tabs preserve selected paragraphs, save, move, cancel, delete and undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-ruler-tabs-'));
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
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      await saved();
      await editor.getByRole('tab', { name: 'View', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'View', exact: true })
        .getByRole('checkbox', { name: 'Ruler', exact: true })
        .check();
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.canvas-shell .inline-edit');
      await input.evaluate((node) => {
        window.selectEditorText(node, 17, 22);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      const readTabs = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(pres)[0])[0];
        return [
          getParagraphPropertiesEffective(pres, shape, 0).tabStops ?? [],
          getParagraphPropertiesEffective(pres, shape, 1).tabStops ?? [],
        ];
      };
      const revision = async () => (await waitForState(preview.url, () => true)).revision;
      const change = async (action) => {
        const before = await revision();
        await action();
        await waitForState(preview.url, (state) => state.revision > before);
        await saved();
      };
      for (const [index, alignment] of ['left', 'center', 'right', 'decimal'].entries()) {
        const track = await editor
          .getByRole('button', { name: 'Add tab stop', exact: true })
          .boundingBox();
        await change(() => page.mouse.click(track.x + 80 + index * 55, track.y + 5));
        assert.deepEqual((await readTabs())[0], []);
        assert.equal((await readTabs())[1][index].alignment, alignment);
        if (index < 3) await editor.locator('.tab-selector').click();
      }
      assert.equal(
        await input.evaluate((node) => node.ownerDocument.getSelection().toString()),
        'econd',
      );
      const original = (await readTabs())[1];
      const drag = async (dx, dy, cancel = false) => {
        const box = await editor.locator('.tab-stop').first().boundingBox();
        await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
        await page.mouse.down();
        await page.mouse.move(box.x + box.width / 2 + dx, box.y + box.height / 2 + dy, {
          steps: 5,
        });
        if (cancel) await input.press('Escape');
        await page.mouse.up();
      };
      await change(() => drag(20, 0));
      const moved = (await readTabs())[1];
      assert.ok(moved[0].positionEmu > original[0].positionEmu);
      assert.deepEqual(moved.slice(1), original.slice(1));
      const beforeCancel = await revision();
      await drag(20, 50, true);
      assert.equal(await revision(), beforeCancel);
      assert.deepEqual((await readTabs())[1], moved);
      await change(() => drag(0, 70));
      assert.deepEqual((await readTabs())[1], moved.slice(1));
      await change(() => editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click());
      assert.deepEqual((await readTabs())[1], moved);
      await page.reload();
      await saved();
      assert.deepEqual((await readTabs())[1], moved);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
