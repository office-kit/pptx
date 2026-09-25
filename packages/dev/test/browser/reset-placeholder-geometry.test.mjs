import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addSlide,
  removeShape,
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
  setShapeTextFormat,
  setShapeFill,
  getShapeFill,
  getShapeXmlString,
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

test(
  'restore deleted placeholders on selected slides with bilingual undo and reload',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-restore-slots-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      for (let index = 0; index < 3; index++) {
        const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content') });
        const [title, body] = getSlideShapes(slide);
        setShapeText(title, `Keep ${index}`);
        removeShape(body);
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
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
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
        return getSlides(deck).map((slide) => getSlideShapes(slide).map(getShapeText));
      };
      await saved();
      const thumbs = editor.locator('.thumb-row');
      await thumbs.nth(0).click();
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      await editor
        .getByRole('region', { name: 'Slide options', exact: true })
        .getByRole('button', { name: 'Restore deleted placeholders', exact: true })
        .click();
      await saved();
      assert.deepEqual(await state(), [['Keep 0', ''], ['Keep 1', ''], ['Keep 2']]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await state(), [['Keep 0'], ['Keep 1'], ['Keep 2']]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await state(), [['Keep 0', ''], ['Keep 1', ''], ['Keep 2']]);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await thumbs.nth(2).click();
      const restore = editor
        .locator('.ribbon')
        .getByRole('button', { name: '削除したプレースホルダーを復元', exact: true });
      await restore.click();
      await saved();
      await restore.click();
      await saved();
      await page.screenshot({ path: '/tmp/pptx-restore-slots-ja.png', fullPage: true });
      await page.reload();
      await saved();
      assert.deepEqual(await state(), [
        ['Keep 0', ''],
        ['Keep 1', ''],
        ['Keep 2', ''],
      ]);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'reset placeholder text formatting applies to selected slides with bilingual history and persistence',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-reset-text-'));
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
        setShapeTextFormat(title, { size: 44, bold: true, color: '#CC1122' });
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
          formatted: titles.map((shape) => getShapeXmlString(shape).includes('sz="4400"')),
          texts: titles.map(getShapeText),
        };
      };
      await saved();
      const thumbs = editor.locator('.thumb-row');
      await thumbs.nth(0).click();
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      await editor
        .getByRole('region', { name: 'Slide options', exact: true })
        .getByRole('button', { name: 'Reset placeholder text formatting', exact: true })
        .click();
      await saved();
      assert.deepEqual((await state()).formatted, [false, false, true]);
      assert.deepEqual((await state()).bounds, moved);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual((await state()).bounds, moved);
      assert.deepEqual((await state()).formatted, [true, true, true]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual((await state()).formatted, [false, false, true]);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await thumbs.nth(2).click();
      await editor
        .locator('.ribbon')
        .getByRole('button', { name: 'プレースホルダーの文字書式を戻す', exact: true })
        .click();
      await saved();
      assert.deepEqual((await state()).formatted, [false, false, false]);
      await page.screenshot({ path: '/tmp/pptx-reset-text-ja.png', fullPage: true });
      await page.reload();
      await saved();
      const result = await state();
      assert.deepEqual(result.bounds, moved);
      assert.deepEqual(result.rotations, [30, 30, 30]);
      assert.deepEqual(result.formatted, [false, false, false]);
      assert.deepEqual(
        result.texts,
        [0, 1, 2].map((index) => `日本語 / English ${index}`),
      );
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-reset-text-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'reset layout restores geometry, styling and missing slots to selected slides with bilingual history and persistence',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-reset-all-'));
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
        setShapeTextFormat(title, { size: 48, bold: true });
        setShapeFill(title, '#FF0000');
        removeShape(getSlideShapes(slide)[1]);
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
          fills: titles.map((shape) => getShapeFill(shape).kind),
          counts: getSlides(deck).map((slide) => getSlideShapes(slide).length),
          formatted: titles.map((shape) => getShapeXmlString(shape).includes('sz="4800"')),
          texts: titles.map(getShapeText),
        };
      };
      await saved();
      const thumbs = editor.locator('.thumb-row');
      await thumbs.nth(0).click();
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      await editor
        .getByRole('region', { name: 'Slide options', exact: true })
        .getByRole('button', { name: 'Reset layout', exact: true })
        .click();
      await saved();
      assert.deepEqual((await state()).bounds, [expected[0], expected[1], moved[2]]);
      assert.deepEqual((await state()).rotations, [0, 0, 30]);
      assert.deepEqual((await state()).fills, ['inherit', 'inherit', 'solid']);
      assert.deepEqual((await state()).counts, [2, 2, 1]);
      assert.deepEqual((await state()).formatted, [false, false, true]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual((await state()).bounds, moved);
      assert.deepEqual((await state()).rotations, [30, 30, 30]);
      assert.deepEqual((await state()).fills, ['solid', 'solid', 'solid']);
      assert.deepEqual((await state()).counts, [1, 1, 1]);
      assert.deepEqual((await state()).formatted, [true, true, true]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual((await state()).bounds, [expected[0], expected[1], moved[2]]);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await thumbs.nth(2).click();
      await editor
        .locator('.ribbon')
        .getByRole('button', { name: 'レイアウトをリセット', exact: true })
        .click();
      await saved();
      assert.deepEqual((await state()).bounds, expected);
      await page.screenshot({ path: '/tmp/pptx-reset-all-ja.png', fullPage: true });
      await page.reload();
      await saved();
      const result = await state();
      assert.deepEqual(result.bounds, expected);
      assert.deepEqual(result.rotations, [0, 0, 0]);
      assert.deepEqual(result.fills, ['inherit', 'inherit', 'inherit']);
      assert.deepEqual(result.counts, [2, 2, 2]);
      assert.deepEqual(result.formatted, [false, false, false]);
      assert.deepEqual(
        result.texts,
        [0, 1, 2].map((index) => `日本語 / English ${index}`),
      );
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-reset-all-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
