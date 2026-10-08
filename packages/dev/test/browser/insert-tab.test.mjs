import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  findSlidePlaceholder,
  getShapeMedia,
  getShapeName,
  getShapeText,
  getShapeTextDirection,
  getTableCells,
  isTableShape,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>Hello</Text></Slide><Slide><Text x={1} y={1} width={4} height={1}>Two</Text></Slide></Presentation>`;

test(
  'Insert tab matches the reference desktop app: footer, WordArt, Symbol, Video and unavailable commands',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-insert-tab-'));
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
      const slides = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
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
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.getByRole('tab', { name: 'Insert', exact: true }).click();

      // Native order, with commands the library cannot write shown disabled.
      const labels = await panel
        .locator('.insert button:not(.side)')
        .evaluateAll((nodes) => nodes.map((node) => node.getAttribute('aria-label')));
      assert.deepEqual(labels, [
        'New Slide',
        'Table',
        'Pictures',
        'Screenshot',
        'Cameo',
        'Shapes',
        'Icons',
        '3D Models',
        'SmartArt',
        'Chart',
        'Zoom',
        'Link',
        'Action',
        'Comment',
        'Text Box',
        'Header & Footer',
        'WordArt',
        'Date & Time',
        'Slide Number',
        'Object',
        'Equation',
        'Symbol',
        'Video',
        'Audio',
      ]);
      for (const name of ['Cameo', 'Icons', 'SmartArt', 'Equation', 'Symbol'])
        assert.equal(await button(name).isDisabled(), true, name);

      // The reference desktop app's ▾ menus, with what the browser cannot do disabled.
      const items = async (name) => {
        await button(name).click();
        const menu = panel.getByRole('menu', { name, exact: true });
        const result = await menu
          .getByRole('menuitem')
          .evaluateAll((nodes) =>
            nodes
              .filter((node) => !node.classList.contains('cell'))
              .map((node) => `${node.textContent.trim()}${node.disabled ? ' (off)' : ''}`),
          );
        await page.keyboard.press('Escape');
        return result;
      };
      assert.deepEqual(await items('Pictures'), [
        'Photo Browser... (off)',
        'Picture from File...',
        'Stock Images... (off)',
        'Online Pictures... (off)',
      ]);
      assert.deepEqual(await items('Video'), [
        'Movie Browser... (off)',
        'Movie from File...',
        'Online Movie... (off)',
      ]);
      assert.deepEqual(await items('Audio'), [
        'Audio Browser... (off)',
        'Audio from File...',
        'Record Audio... (off)',
      ]);
      assert.deepEqual((await items('Chart')).slice(0, 5), [
        'Column',
        'Line',
        'Pie',
        'Bar',
        'Area',
      ]);

      // Table ▸ the 3 × 2 cell of the grid inserts a two-column, three-row table.
      await button('Table').click();
      await changed(() => panel.getByRole('menuitem', { name: '2x3 Table', exact: true }).click());
      const table = getSlideShapes((await slides())[0]).find((shape) => isTableShape(shape));
      assert.deepEqual([getTableCells(table).length, getTableCells(table)[0].length], [3, 2]);
      await page.keyboard.press('Escape');
      await page.keyboard.press('Escape');

      // Text Box ▸ Draw Vertical Text Box rotates the text 90°.
      await panel.getByRole('button', { name: 'Text Box options', exact: true }).click();
      await changed(() =>
        panel.getByRole('menuitem', { name: 'Draw Vertical Text Box', exact: true }).click(),
      );
      assert.ok(
        getSlideShapes((await slides())[0]).some(
          (shape) => getShapeTextDirection(shape) === 'vert',
        ),
      );

      // Header & Footer ▸ Footer ▸ Apply to All fills every slide's footer slot.
      await button('Header & Footer').click();
      const dialog = editor.getByRole('dialog', { name: 'Header and Footer' });
      await dialog.getByRole('checkbox', { name: 'Footer', exact: true }).check();
      await dialog.getByRole('textbox', { name: 'Footer text' }).fill('Confidential');
      await changed(() =>
        dialog.getByRole('button', { name: 'Apply to All', exact: true }).click(),
      );
      // The DSL template has no footer slot, so each slide gets a "Footer" box.
      for (const slide of await slides()) {
        const footer =
          findSlidePlaceholder(slide, 'ftr') ??
          getSlideShapes(slide).find((shape) => getShapeName(shape) === 'Footer');
        assert.equal(getShapeText(footer), 'Confidential');
      }

      await button('WordArt').click();
      await changed(() =>
        editor
          .getByRole('menu', { name: 'WordArt', exact: true })
          .getByRole('menuitem', { name: 'Fill: Blue, Accent color 1; Shadow', exact: true })
          .click(),
      );
      assert.ok(
        getSlideShapes((await slides())[0]).some(
          (shape) => getShapeText(shape) === 'Your text here',
        ),
      );

      // Symbol inserts at the text cursor while editing.
      await page.keyboard.press('Escape');
      const hello = editor.locator('.hit').first();
      const inline = editor.locator('.canvas-shell .inline-edit');
      // A click on a selected text box already places the cursor.
      await hello.click();
      if (!(await inline.isVisible())) await hello.dblclick();
      await inline.waitFor();
      await page.keyboard.press('End');
      await button('Symbol').click();
      await editor
        .getByRole('dialog', { name: 'Symbol' })
        .getByRole('button', { name: '→', exact: true })
        .click();
      await changed(() => page.keyboard.press('Escape'));
      assert.ok(
        getSlideShapes((await slides())[0]).some((shape) => getShapeText(shape) === 'Hello→'),
      );

      // Video ▸ from a file on this computer.
      const mp4 = Buffer.concat([
        Buffer.from([0, 0, 0, 0x18]),
        Buffer.from('ftypisom'),
        Buffer.alloc(12),
      ]);
      const chooser = page.waitForEvent('filechooser');
      await button('Video').click();
      await panel.getByRole('menuitem', { name: 'Movie from File...', exact: true }).click();
      await changed(async () =>
        (await chooser).setFiles({ name: 'clip.mp4', mimeType: 'video/mp4', buffer: mp4 }),
      );
      assert.ok(
        getSlideShapes((await slides())[0]).some((shape) => getShapeMedia(shape)?.kind === 'video'),
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
