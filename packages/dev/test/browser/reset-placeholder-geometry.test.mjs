import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addSlide,
  createPresentation,
  findSlideLayout,
  getShapeBoundsResolved,
  getShapeRotation,
  getShapeText,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapePosition,
  setShapeRotation,
  setShapeSize,
  setShapeText,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'reset placeholder positions applies to selected slides with bilingual history and persistence',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-reset-layout-'));
    let preview, browser, page;
    try {
      const pres = createPresentation();
      const expected = [],
        moved = [];
      for (let index = 0; index < 3; index++) {
        const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content') });
        const title = getSlideShapes(slide)[0];
        expected.push(getShapeBoundsResolved(pres, title));
        setShapeText(title, `日本語 / English ${index}`);
        setShapePosition(title, inches(3), inches(2));
        setShapeSize(title, inches(4), inches(1));
        setShapeRotation(title, 30);
        moved.push(getShapeBoundsResolved(pres, title));
      }
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(pres));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const state = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const titles = getSlides(deck).map((slide) => getSlideShapes(slide)[0]);
        return {
          bounds: titles.map((shape) => getShapeBoundsResolved(deck, shape)),
          rotations: titles.map(getShapeRotation),
          texts: titles.map(getShapeText),
        };
      };
      await saved();
      const thumbs = editor.locator('.thumb-row');
      await thumbs.nth(0).click();
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      await editor
        .getByRole('region', { name: 'Slide options', exact: true })
        .getByRole('button', { name: 'Reset placeholder positions', exact: true })
        .click();
      await saved();
      assert.deepEqual((await state()).bounds, [expected[0], expected[1], moved[2]]);
      assert.deepEqual((await state()).rotations, [0, 0, 30]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual((await state()).bounds, moved);
      assert.deepEqual((await state()).rotations, [30, 30, 30]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual((await state()).bounds, [expected[0], expected[1], moved[2]]);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await thumbs.nth(2).click();
      await editor
        .locator('.ribbon')
        .getByRole('button', { name: 'プレースホルダーの配置を戻す', exact: true })
        .click();
      await saved();
      assert.deepEqual((await state()).bounds, expected);
      await page.screenshot({ path: '/tmp/pptx-reset-layout-ja.png', fullPage: true });
      await page.reload();
      await saved();
      const result = await state();
      assert.deepEqual(result.bounds, expected);
      assert.deepEqual(result.rotations, [0, 0, 0]);
      assert.deepEqual(
        result.texts,
        [0, 1, 2].map((index) => `日本語 / English ${index}`),
      );
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-reset-layout-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
