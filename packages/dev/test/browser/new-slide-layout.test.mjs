import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getShapeText,
  getSlideLayout,
  getSlideLayoutName,
  isShapePlaceholder,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'new slides use the chosen layout with editable placeholders, history and bilingual persistence',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-new-layout-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['Before','After'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const thumbs = editor.locator('.thumb-row');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const slides = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        );
      await saved();
      await thumbs.nth(0).click();
      // Home ▸ New Slide ▾ lists the layouts, as in the reference desktop app. A narrow
      // ribbon first collapses the Slides group into a single button.
      const newSlideFrom = async (slidesLabel, optionsLabel, layout) => {
        const group = editor.getByRole('button', { name: slidesLabel, exact: true });
        if (await group.isVisible()) await group.click();
        await editor.getByRole('button', { name: optionsLabel, exact: true }).click();
        await editor.getByRole('menuitem', { name: layout, exact: true }).click();
      };
      await newSlideFrom('Slides', 'New Slide options', 'Title and Content');
      await saved();
      assert.equal(await thumbs.count(), 3);
      assert.match(await thumbs.nth(1).getAttribute('class'), /active/);
      let deck = await slides();
      assert.equal(getSlideLayoutName(getSlideLayout(deck[1])), 'Title and Content');
      assert.ok(getSlideShapes(deck[1]).filter(isShapePlaceholder).length >= 2);
      assert.equal(getShapeText(getSlideShapes(deck[0])[0]), 'Before');
      assert.equal(getShapeText(getSlideShapes(deck[2])[0]), 'After');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await slides()).length, 2);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal((await slides()).length, 3);
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      const input = editor.locator('.canvas-shell .inline-edit');
      await input.fill('新しいスライド / New slide');
      await input.press('Control+Enter');
      await saved();
      assert.ok(
        getSlideShapes((await slides())[1]).some(
          (shape) => getShapeText(shape) === '新しいスライド / New slide',
        ),
      );
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await newSlideFrom('スライド', '新しいスライドのオプション', 'タイトルスライド');
      await page.screenshot({ path: '/tmp/pptx-new-slide-layout-ja.png', fullPage: true });
      await saved();
      assert.equal(getSlideLayoutName(getSlideLayout((await slides())[2])), 'Title Slide');
      await page.reload();
      await saved();
      deck = await slides();
      assert.equal(deck.length, 4);
      assert.ok(
        getSlideShapes(deck[1]).some(
          (shape) => getShapeText(shape) === '新しいスライド / New slide',
        ),
      );
      assert.equal(getSlideLayoutName(getSlideLayout(deck[2])), 'Title Slide');
      assert.equal(getShapeText(getSlideShapes(deck[3])[0]), 'After');
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-new-slide-layout-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'generic New Slide advances Title Slide to Title and Content from thumbnails and keyboard',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-generic-new-layout-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide layout={{name:'Title Slide'}} /></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const slides = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        );
      const thumbs = editor.locator('.thumb-row');
      await saved();

      await thumbs.first().click({ button: 'right' });
      await editor.getByRole('menuitem', { name: 'New Slide', exact: true }).click();
      await saved();
      let deck = await slides();
      assert.equal(deck.length, 2);
      assert.equal(getSlideLayoutName(getSlideLayout(deck[0])), 'Title Slide');
      assert.equal(getSlideLayoutName(getSlideLayout(deck[1])), 'Title and Content');

      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await slides()).length, 1);

      await thumbs.first().click();
      await editor.locator('body').press('Meta+Shift+n');
      await saved();
      deck = await slides();
      assert.equal(deck.length, 2);
      assert.equal(getSlideLayoutName(getSlideLayout(deck[1])), 'Title and Content');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await slides()).length, 1);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
