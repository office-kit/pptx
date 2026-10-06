import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeBodyPrEffective,
  getShapeText,
  getShapeTextDirection,
  getSlides,
  getSlideSections,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>Hello</Text></Slide><Slide><Text x={1} y={1} width={4} height={1}>Two</Text></Slide><Slide><Text x={1} y={1} width={4} height={1}>Three</Text></Slide></Presentation>`;

test(
  'Home tab matches PowerPoint: Section, Columns, Text Direction, Align Text, Add-ins and Designer',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-home-tab-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const pres = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const changed = async (action) => {
        const before = await (await fetch(preview.url + '/editor/state')).json();
        await action();
        for (let i = 0; i < 200; i += 1) {
          const state = await (await fetch(preview.url + '/editor/state')).json();
          if (!state.building && state.revision !== before.revision) return;
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        throw new Error('The edit was not saved');
      };
      const panel = editor.locator('#ribbon-panel');
      const button = (name) => panel.getByRole('button', { name, exact: true });
      const item = (name) => panel.getByRole('menuitem', { name, exact: true });
      await editor.getByText('Saved to this project', { exact: true }).waitFor();

      for (const name of ['Add-ins', 'Designer', 'Convert to SmartArt'])
        assert.equal(await button(name).isDisabled(), true, name);

      // New Slide ▾ is a gallery of layout thumbnails, as in PowerPoint.
      await button('New Slide options').click();
      const gallery = panel.getByRole('menu', { name: 'New Slide', exact: true });
      assert.ok((await gallery.locator('.layout-item svg.layout-thumbnail').count()) > 0);
      await page.keyboard.press('Escape');

      // Section ▸ Add Section on slide 2 puts slide 1 in "Default Section".
      await editor.locator('.thumb-row').nth(1).click();
      await button('Section').click();
      await item('Add Section').click();
      const rename = editor.getByRole('dialog', { name: 'Rename Section' });
      await rename.getByRole('textbox').fill('Results');
      await changed(() => rename.getByRole('button', { name: 'Rename', exact: true }).click());
      assert.deepEqual(
        getSlideSections(await pres()).map((section) => [section.name, section.slides.length]),
        [
          ['Default Section', 1],
          ['Results', 2],
        ],
      );
      // The thumbnail pane shows each section's name above its first slide.
      assert.deepEqual(await editor.locator('.section-header').allTextContents(), [
        'Default Section',
        'Results',
      ]);
      await button('Section').click();
      await changed(() => item('Remove Section').click());
      assert.deepEqual(
        getSlideSections(await pres()).map((section) => [section.name, section.slides.length]),
        [['Default Section', 3]],
      );
      await button('Section').click();
      await changed(() => item('Remove All Sections').click());
      assert.equal(getSlideSections(await pres()).length, 0);

      // The thumbnail menu is PowerPoint's.
      await editor.locator('.thumb-row').nth(1).click({ button: 'right' });
      const thumbMenu = editor.getByRole('menu').first();
      assert.deepEqual(
        (
          await thumbMenu
            .locator(':scope > .ctx-item, :scope > .branch > .ctx-item')
            .allTextContents()
        ).map((text) => text.replace(/(\s+Del|[›⌘✓]).*$/, '').trim()),
        [
          'Cut',
          'Copy',
          'Paste',
          'New Slide',
          'Duplicate Slide',
          'Delete Slide',
          'Add Section',
          'Layout',
          'Reset Slide',
          'Format Background...',
          'New Comment',
          'Hide Slide',
        ],
      );
      await page.keyboard.press('Escape');

      // Columns, Text Direction and Align Text act on the selected text box.
      await editor.locator('.thumb-row').first().click();
      await editor.locator('.hit').first().click();
      const hello = async () => {
        const deck = await pres();
        const shape = getSlideShapes(getSlides(deck)[0]).find((s) => getShapeText(s) === 'Hello');
        return { deck, shape, body: getShapeBodyPrEffective(deck, shape) };
      };
      await button('Columns').click();
      await changed(() => item('Two Columns').click());
      assert.equal((await hello()).body.columns?.count, 2);
      await button('Text Direction').click();
      await changed(() =>
        panel.getByRole('menuitemradio', { name: 'Rotate all text 90°', exact: true }).click(),
      );
      assert.equal(getShapeTextDirection((await hello()).shape), 'vert');
      await button('Align Text').click();
      await changed(() =>
        panel.getByRole('menuitemradio', { name: 'Bottom', exact: true }).click(),
      );
      assert.equal((await hello()).body.anchor, 'bottom');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
