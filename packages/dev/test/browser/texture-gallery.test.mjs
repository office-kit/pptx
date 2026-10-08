import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeFill,
  getShapeImageFillBytes,
  getShapeImageFillLayout,
  getSlideBackgroundImageFillLayout,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { expandFormatSections } from '../helpers/format-pane.mjs';
import { startPreview } from '../helpers/server.mjs';
import { slideBackgroundImageBytes, textureIdOf } from '../helpers/textures.mjs';

const NAMES = [
  'Papyrus',
  'Canvas',
  'Denim',
  'Woven mat',
  'Water droplets',
  'Paper bag',
  'Fish fossil',
  'Sand',
  'Green marble',
  'White marble',
  'Brown marble',
  'Granite',
  'Newsprint',
  'Recycled paper',
  'Parchment',
  'Stationery',
  'Blue tissue paper',
  'Pink tissue paper',
  'Purple mesh',
  'Bouquet',
  'Cork',
  'Walnut',
  'Oak',
  'Medium wood',
];
const JA_NAMES = [
  'パピルス',
  'キャンバス',
  'デニム',
  '麻',
  '水滴',
  '紙袋',
  '化石',
  '砂',
  '緑の大理石',
  '白の大理石',
  '茶色の大理石',
  'みかげ石',
  '新聞紙',
  '再生紙',
  'セーム皮',
  '便箋',
  '青い画用紙',
  'ピンクの画用紙',
  '紫のメッシュ',
  'ブーケ',
  'コルク',
  'ウォールナット',
  'オーク',
  '木目',
];
// The reference desktop app's texture tiling: no offset, 100%, top left, no mirror, rotate with shape.
const TILE = {
  mode: 'tile',
  offsetX: 0,
  offsetY: 0,
  scaleX: 1,
  scaleY: 1,
  alignment: 'tl',
  flip: 'none',
  rotateWithShape: true,
};

const DECK = `<Slide><Text x={1} y={1} width={3} height={2} fill="#00FF00">First</Text><Text x={5} y={1} width={3} height={2} fill="#0000FF">Second</Text></Slide><Slide><Text x={1} y={1} width={3} height={2}>Other</Text></Slide>`;

async function withEditor(locale, run, slides = DECK) {
  const dir = await mkdtemp(join(tmpdir(), 'office-texture-gallery-'));
  let preview, browser;
  try {
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>${slides}</Presentation>`,
    );
    preview = await startPreview(file);
    browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
    const errors = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.addInitScript(
      (language) => localStorage.setItem('ok-editor-locale', language),
      locale,
    );
    await page.goto(preview.url);
    const editor = page.frameLocator('#editor-frame');
    const deck = async () =>
      loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
    await run({ page, editor, deck });
    assert.deepEqual(errors, []);
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
}

const swatchNames = (gallery) =>
  gallery
    .locator('.grid [role="menuitem"]')
    .evaluateAll((items) => items.map((item) => item.getAttribute('aria-label')));

test(
  'Texture gallery tiles shapes and backgrounds from the Format pane and Shape Fill',
  { timeout: 180000 },
  async () => {
    await withEditor('en', async ({ page, editor, deck }) => {
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      // Waits for the edit to reach the project file.
      const commit = async (action) => {
        const persisted = page.waitForResponse(
          (response) =>
            response.url().endsWith('/editor/document') &&
            response.request().method() === 'PUT' &&
            response.ok(),
        );
        await action();
        await persisted;
        await saved();
      };
      const undo = () => commit(() => editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click());
      const shapes = async () =>
        getSlideShapes(getSlides(await deck())[0]).map((shape) => ({
          fill: getShapeFill(shape).kind,
          texture: textureIdOf(getShapeImageFillBytes(shape)),
          layout: getShapeImageFillLayout(shape),
        }));
      const gallery = editor.getByRole('menu', { name: 'Texture', exact: true });
      const tiled = editor.getByRole('checkbox', { name: 'Tile picture as texture', exact: true });
      const picture = editor.getByRole('radio', { name: 'Picture or texture fill', exact: true });
      const texture = editor.getByRole('button', { name: 'Texture', exact: true });
      // Like the reference desktop app's, the gallery never marks the current texture.
      const assertNothingMarked = async () => {
        await texture.click();
        assert.equal(await gallery.locator('[aria-checked], [aria-selected]').count(), 0);
        await page.keyboard.press('Escape');
      };
      // The reference desktop app never opens a file chooser here: a fill with no picture to
      // restore gets the default texture, Papyrus, tiled.
      page.on('filechooser', () => assert.fail('Picture or texture fill opened a file chooser'));
      await saved();

      // Both shapes: Format Shape ▸ Picture or texture fill, then Texture ▾ ▸ Canvas.
      await editor
        .locator('.hit')
        .first()
        .click({ position: { x: 2, y: 2 } });
      await editor
        .locator('.hit')
        .nth(1)
        .click({ position: { x: 2, y: 2 }, modifiers: ['Shift'] });
      await editor
        .locator('.hit')
        .first()
        .click({ button: 'right', position: { x: 2, y: 2 } });
      await editor.getByRole('menuitem', { name: 'Format Shape...', exact: true }).click();
      await expandFormatSections(editor, 'Fill', 'Line');
      const original = await shapes();
      const paneTitle = () => editor.locator('.panel-head strong').textContent();
      assert.equal(await paneTitle(), 'Format Shape');
      await commit(() => picture.click());
      const before = await shapes();
      for (const shape of before) {
        assert.deepEqual(shape, { fill: 'image', texture: 'papyrus', layout: TILE });
      }
      assert.equal(await picture.isChecked(), true);
      assert.equal(await tiled.isChecked(), true);
      assert.equal(await paneTitle(), 'Format Picture');
      await assertNothingMarked();
      // Choosing the fill is one undo step for both shapes.
      await undo();
      assert.deepEqual(await shapes(), original);
      assert.equal(await paneTitle(), 'Format Shape');
      await commit(() => editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click());
      assert.deepEqual(await shapes(), before);
      await texture.click();
      assert.deepEqual(await swatchNames(gallery), NAMES);
      // As in the reference desktop app, the swatch button sits at the right of the Texture row,
      // flush with the other value controls, and the gallery hangs from its right edge.
      const box = async (locator) => (await locator.boundingBox()) ?? assert.fail('not visible');
      const [button, alignment, menu] = await Promise.all([
        box(texture),
        box(editor.locator('#format-panel select').filter({ visible: true }).first()),
        box(gallery),
      ]);
      assert.ok(Math.abs(button.x + button.width - (alignment.x + alignment.width)) <= 1);
      assert.ok(Math.abs(button.x + button.width - (menu.x + menu.width)) <= 1);
      assert.ok(menu.y >= button.y + button.height);
      // Five to a row, with Insert... beside it instead of More Textures...
      const rows = await gallery
        .locator('.grid [role="menuitem"]')
        .evaluateAll((items) => items.map((item) => Math.round(item.getBoundingClientRect().top)));
      assert.deepEqual(
        rows.map((top) => [...new Set(rows)].indexOf(top)),
        NAMES.map((_, index) => Math.floor(index / 5)),
      );
      assert.equal(
        await gallery.getByRole('menuitem', { name: 'More Textures...', exact: true }).count(),
        0,
      );
      const focused = () =>
        gallery.locator(':focus').evaluate((item) => item.getAttribute('aria-label'));
      // Keyboard: focus starts on Papyrus; Down reaches the next row of five,
      // Up returns, then Right and Enter pick Canvas.
      assert.equal(await focused(), 'Papyrus');
      await page.keyboard.press('ArrowDown');
      assert.equal(await focused(), 'Paper bag');
      await page.keyboard.press('ArrowUp');
      assert.equal(await focused(), 'Papyrus');
      await commit(async () => {
        await page.keyboard.press('ArrowRight');
        await page.keyboard.press('Enter');
      });
      assert.equal(await gallery.count(), 0);
      for (const shape of await shapes()) {
        assert.deepEqual(shape, { fill: 'image', texture: 'canvas', layout: TILE });
      }
      assert.equal(await tiled.isChecked(), true);
      await assertNothingMarked();
      assert.equal(await gallery.count(), 0);

      // One undo step restores both shapes.
      await undo();
      assert.deepEqual(await shapes(), before);
      await commit(() => editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click());

      // Shape Fill ▾ ▸ Texture ▸ Oak on the ribbon.
      await editor.getByRole('tab', { name: 'Home', exact: true }).click();
      await editor.getByRole('button', { name: 'Shape Fill', exact: true }).click();
      await editor.getByRole('menuitem', { name: 'Texture', exact: true }).click();
      assert.deepEqual(await swatchNames(gallery), NAMES);
      assert.equal(await gallery.locator('[aria-checked], [aria-selected]').count(), 0);
      await gallery.getByRole('menuitem', { name: 'More Textures...', exact: true }).waitFor();
      await commit(() => gallery.getByRole('menuitem', { name: 'Oak', exact: true }).click());
      for (const shape of await shapes()) {
        assert.deepEqual(shape, { fill: 'image', texture: 'oak', layout: TILE });
      }
      await undo();
      for (const shape of await shapes()) assert.equal(shape.texture, 'canvas');

      // Format Background ▸ Picture or texture fill ▸ Texture ▾ ▸ Walnut, then Apply to All.
      // Apply to All moves the background to the master, so read the inherited one.
      const backgrounds = async () => {
        const pres = await deck();
        return getSlides(pres).map((slide) => ({
          texture: textureIdOf(slideBackgroundImageBytes(pres, slide)),
          layout: getSlideBackgroundImageFillLayout(slide),
        }));
      };
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor.getByRole('button', { name: 'Background Styles', exact: true }).click();
      await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();
      const plain = await backgrounds();
      assert.deepEqual(
        plain.map((background) => background.texture),
        [null, null],
      );
      // Oak, last picked for a shape, is the session's default texture.
      await commit(() => picture.click());
      assert.deepEqual(await backgrounds(), [{ texture: 'oak', layout: TILE }, plain[1]]);
      assert.equal(await picture.isChecked(), true);
      await assertNothingMarked();
      await undo();
      assert.deepEqual(await backgrounds(), plain);
      assert.equal(await picture.isChecked(), false);
      await commit(() => editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click());
      assert.equal((await backgrounds())[0].texture, 'oak');
      await texture.click();
      assert.deepEqual(await swatchNames(gallery), NAMES);
      await commit(() => gallery.getByRole('menuitem', { name: 'Walnut', exact: true }).click());
      let state = await backgrounds();
      assert.deepEqual(state[0], { texture: 'walnut', layout: TILE });
      assert.equal(state[1].texture, null);
      assert.equal(await tiled.isChecked(), true);
      await assertNothingMarked();
      await commit(() => editor.getByRole('button', { name: 'Apply to All', exact: true }).click());
      state = await backgrounds();
      assert.deepEqual(state, [
        { texture: 'walnut', layout: TILE },
        { texture: 'walnut', layout: TILE },
      ]);
      await undo();
      assert.equal((await backgrounds())[1].texture, null);
      await undo();
      assert.equal((await backgrounds())[0].texture, 'oak');
    });
  },
);

test(
  'Picture or texture fill defaults to the last texture picked for a shape until reload',
  { timeout: 180000 },
  async () => {
    // Replays the native capture: one in-memory "last texture" per session,
    // starting at Papyrus, set only by shape texture picks, used by shapes and
    // backgrounds alike and reset by relaunching (here: reloading the page).
    await withEditor(
      'en',
      async ({ page, editor, deck }) => {
        const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
        const commit = async (action) => {
          const persisted = page.waitForResponse(
            (response) =>
              response.url().endsWith('/editor/document') &&
              response.request().method() === 'PUT' &&
              response.ok(),
          );
          await action();
          await persisted;
          await saved();
        };
        const picture = editor.getByRole('radio', { name: 'Picture or texture fill', exact: true });
        const gallery = editor.getByRole('menu', { name: 'Texture', exact: true });
        const pick = (name) =>
          commit(async () => {
            await editor.getByRole('button', { name: 'Texture', exact: true }).click();
            await gallery.getByRole('menuitem', { name, exact: true }).click();
          });
        const shapeTexture = async (slide, index) =>
          textureIdOf(
            getShapeImageFillBytes(getSlideShapes(getSlides(await deck())[slide])[index]),
          );
        const backgroundTexture = async (slide) => {
          const pres = await deck();
          return textureIdOf(slideBackgroundImageBytes(pres, getSlides(pres)[slide]));
        };
        const formatShape = async (slide, index) => {
          await editor.locator('.thumb-row').nth(slide).click();
          await editor
            .locator('.hit')
            .nth(index)
            .click({ button: 'right', position: { x: 2, y: 2 } });
          await editor.getByRole('menuitem', { name: 'Format Shape...', exact: true }).click();
          await expandFormatSections(editor, 'Fill');
        };
        const formatBackground = async (slide) => {
          await editor.locator('.thumb-row').nth(slide).click();
          await editor.getByRole('tab', { name: 'Design', exact: true }).click();
          await editor.getByRole('button', { name: 'Background Styles', exact: true }).click();
          await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();
        };
        await saved();

        await formatShape(0, 0);
        await commit(() => picture.click());
        assert.equal(await shapeTexture(0, 0), 'papyrus');
        await pick('Canvas');
        assert.equal(await shapeTexture(0, 0), 'canvas');
        await formatShape(0, 1);
        await commit(() => picture.click());
        assert.equal(await shapeTexture(0, 1), 'canvas');
        await formatBackground(0);
        await commit(() => picture.click());
        assert.equal(await backgroundTexture(0), 'canvas');
        // A texture picked for a background does not become the default.
        await pick('Denim');
        assert.equal(await backgroundTexture(0), 'denim');
        await formatShape(1, 0);
        await commit(() => picture.click());
        assert.equal(await shapeTexture(1, 0), 'canvas');
        await pick('Sand');
        await formatBackground(1);
        await commit(() => picture.click());
        assert.equal(await backgroundTexture(1), 'sand');
        await pick('Granite');
        await formatBackground(2);
        await commit(() => picture.click());
        assert.equal(await backgroundTexture(2), 'sand');
        // Undo does not forget the session texture either.
        await commit(() => editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click());
        await commit(() => picture.click());
        assert.equal(await backgroundTexture(2), 'sand');

        // Nothing about it is saved: a reload starts again from Papyrus.
        await page.reload();
        await saved();
        await formatShape(0, 2);
        await commit(() => picture.click());
        assert.equal(await shapeTexture(0, 2), 'papyrus');
        await formatBackground(3);
        await commit(() => picture.click());
        assert.equal(await backgroundTexture(3), 'papyrus');
      },
      `<Slide><Text x={1} y={1} width={3} height={2}>A</Text><Text x={5} y={1} width={3} height={2}>B</Text><Text x={9} y={1} width={3} height={2}>C</Text></Slide><Slide><Text x={1} y={1} width={3} height={2}>D</Text></Slide><Slide><Text x={1} y={1} width={3} height={2}>E</Text></Slide><Slide><Text x={1} y={1} width={3} height={2}>F</Text></Slide>`,
    );
  },
);

test('Texture gallery shows Japanese names', { timeout: 120000 }, async () => {
  await withEditor('ja', async ({ page, editor, deck }) => {
    const saved = () => editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
    await saved();
    // 図形の書式設定 ▸ 塗りつぶし（図またはテクスチャ） inserts パピルス.
    await editor
      .locator('.hit')
      .first()
      .click({ button: 'right', position: { x: 2, y: 2 } });
    await editor.getByRole('menuitem', { name: '図形の書式設定...', exact: true }).click();
    await expandFormatSections(editor, '塗りつぶし');
    const persisted = page.waitForResponse(
      (response) =>
        response.url().endsWith('/editor/document') &&
        response.request().method() === 'PUT' &&
        response.ok(),
    );
    await editor
      .getByRole('radio', { name: '塗りつぶし（図またはテクスチャ）', exact: true })
      .click();
    await persisted;
    await saved();
    assert.equal(await editor.locator('.panel-head strong').textContent(), '図の書式設定');
    const shape = getSlideShapes(getSlides(await deck())[0])[0];
    assert.equal(textureIdOf(getShapeImageFillBytes(shape)), 'papyrus');
    assert.deepEqual(getShapeImageFillLayout(shape), TILE);
    await editor.getByRole('button', { name: 'テクスチャ', exact: true }).click();
    const paneGallery = editor.getByRole('menu', { name: 'テクスチャ', exact: true });
    assert.deepEqual(await swatchNames(paneGallery), JA_NAMES);
    assert.equal(await paneGallery.locator('[aria-checked], [aria-selected]').count(), 0);
    await page.keyboard.press('Escape');
    await editor
      .locator('.hit')
      .first()
      .click({ position: { x: 2, y: 2 } });
    // Home's Drawing group collapses at this width with Japanese labels on
    // Linux fonts; the contextual tab always shows Shape Fill.
    await editor.getByRole('tab', { name: '図形の書式', exact: true }).click();
    await editor.getByRole('button', { name: '図形の塗りつぶし', exact: true }).click();
    await editor.getByRole('menuitem', { name: 'テクスチャ', exact: true }).click();
    const gallery = editor.getByRole('menu', { name: 'テクスチャ', exact: true });
    assert.deepEqual(await swatchNames(gallery), JA_NAMES);
    await gallery.getByRole('menuitem', { name: 'その他のテクスチャ...', exact: true }).waitFor();
  });
});
