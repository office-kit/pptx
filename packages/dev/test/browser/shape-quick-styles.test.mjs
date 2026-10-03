import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeParagraphElements,
  getShapeXmlString,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'shape quick styles expose the PowerPoint gallery and preserve text through history',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-shape-quick-styles-'));
    let preview;
    let browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={6} height={1} text="Bold title / regular" format={{bold:true}} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const readDeck = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
      const saved = (editor) =>
        editor.getByText('Saved to this project', { exact: true }).waitFor();

      for (const width of [900, 1500, 2200]) {
        const page = await browser.newPage({ viewport: { width, height: 900 } });
        const pageErrors = [];
        page.on('pageerror', (error) => pageErrors.push(error));
        try {
          await page.goto(preview.url);
          await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
          const editor = page.frameLocator('#editor-frame');
          await saved(editor);
          const panel = editor.locator('#ribbon-panel');
          assert.equal(
            await panel.evaluate((node) => node.scrollWidth <= node.clientWidth + 1),
            true,
            `${width}px Home ribbon overflows horizontally`,
          );
          if (width < 1000) {
            assert.deepEqual(pageErrors, [], pageErrors.map(String).join('\n'));
            continue;
          }
          await editor.locator('.hit').first().click();
          await editor.getByRole('tab', { name: 'Shape Format', exact: true }).click();
          const inlineGallery = editor.locator('.inline-gallery');
          assert.equal(await inlineGallery.locator('.inline-item').count(), 3);
          assert.equal(
            await inlineGallery
              .getByRole('button', { name: 'Colored Outline - Dark 1', exact: true })
              .count(),
            1,
          );
          await inlineGallery
            .getByRole('button', { name: 'Next Quick Styles', exact: true })
            .click();
          assert.equal(
            await inlineGallery
              .getByRole('button', { name: 'Colored Outline - Accent 3', exact: true })
              .count(),
            1,
          );
          await inlineGallery
            .getByRole('button', { name: 'Previous Quick Styles', exact: true })
            .click();
          const firstGalleryWidth = await inlineGallery.evaluate(
            (node) => node.getBoundingClientRect().width,
          );
          for (let pageIndex = 0; pageIndex < 25; pageIndex += 1) {
            await inlineGallery
              .getByRole('button', { name: 'Next Quick Styles', exact: true })
              .click();
          }
          const lastGalleryWidth = await inlineGallery.evaluate(
            (node) => node.getBoundingClientRect().width,
          );
          assert.ok(
            Math.abs(lastGalleryWidth - firstGalleryWidth) < 0.5,
            `Shape Format gallery width changed from ${firstGalleryWidth} to ${lastGalleryWidth}`,
          );
          for (let pageIndex = 0; pageIndex < 25; pageIndex += 1) {
            await inlineGallery
              .getByRole('button', { name: 'Previous Quick Styles', exact: true })
              .click();
          }
          await editor.getByRole('tab', { name: 'Home', exact: true }).click();
          if (width <= 2000) {
            await editor
              .locator('.compact-groups .group-menu-trigger')
              .filter({ hasText: 'Drawing' })
              .click();
          }
          const trigger = editor.getByRole('button', { name: 'Quick Styles', exact: true });
          await trigger.click();
          const menu = editor.getByRole('menu', { name: 'Quick Styles', exact: true });
          assert.equal(await menu.getByRole('menuitem').count(), 77);
          await menu.getByRole('menuitem').first().press('Escape');
          assert.equal(await menu.count(), 0);

          if (width > 2000) {
            assert.deepEqual(pageErrors, [], pageErrors.map(String).join('\n'));
            continue;
          }
          await trigger.click();
          const styleButton = menu.getByRole('menuitem', {
            name: 'Colored Fill - Dark 1',
            exact: true,
          });
          await styleButton.click();
          await saved(editor);
          let deck = await readDeck();
          let shape = getSlideShapes(getSlides(deck)[0])[0];
          let xml = getShapeXmlString(shape);
          assert.match(xml, /<p:style>/);
          assert.match(xml, /<a:lnRef idx="2">/);
          assert.equal(getShapeParagraphElements(shape, 0)[0].format?.bold, true);

          await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
          await saved(editor);
          deck = await readDeck();
          shape = getSlideShapes(getSlides(deck)[0])[0];
          assert.doesNotMatch(getShapeXmlString(shape), /<p:style>/);
          assert.equal(getShapeParagraphElements(shape, 0)[0].format?.bold, true);

          await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
          await saved(editor);
          deck = await readDeck();
          shape = getSlideShapes(getSlides(deck)[0])[0];
          assert.match(getShapeXmlString(shape), /<p:style>/);
          assert.equal(getShapeParagraphElements(shape, 0)[0].format?.bold, true);
          assert.deepEqual(pageErrors, [], pageErrors.map(String).join('\n'));
        } finally {
          await page.close();
        }
      }
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
