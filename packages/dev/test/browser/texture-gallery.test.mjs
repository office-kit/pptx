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
import { readSlideBackgroundImageBytes } from '../../../../site/src/lib/editor/core/slide-background.ts';
import { textureIdOf } from '../../../../site/src/lib/editor/core/textures.ts';
import { startPreview } from '../helpers/server.mjs';

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
// PowerPoint's texture tiling: no offset, 100%, top left, no mirror, rotate with shape.
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
const PIXEL = {
  name: 'pixel.png',
  mimeType: 'image/png',
  buffer: Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==',
    'base64',
  ),
};

async function withEditor(locale, run) {
  const dir = await mkdtemp(join(tmpdir(), 'office-texture-gallery-'));
  let preview, browser;
  try {
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={2} fill="#00FF00">First</Text><Text x={5} y={1} width={3} height={2} fill="#0000FF">Second</Text></Slide><Slide><Text x={1} y={1} width={3} height={2}>Other</Text></Slide></Presentation>`,
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
    .getByRole('menuitemradio')
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
      const insertPixel = () =>
        commit(async () => {
          const chooser = page.waitForEvent('filechooser');
          await picture.click();
          await (await chooser).setFiles(PIXEL);
        });
      await saved();

      // Both shapes: Format Shape ▸ Picture or texture fill ▸ Texture ▾ ▸ Canvas.
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
      await insertPixel();
      const before = await shapes();
      assert.deepEqual(
        before.map((shape) => [shape.fill, shape.texture]),
        [
          ['image', null],
          ['image', null],
        ],
      );
      const texture = editor.getByRole('button', { name: 'Texture', exact: true });
      await texture.click();
      assert.deepEqual(await swatchNames(gallery), NAMES);
      assert.equal(await gallery.locator('[aria-checked="true"]').count(), 0);
      await gallery.getByRole('menuitem', { name: 'More Textures...', exact: true }).waitFor();
      // Keyboard: focus starts on Papyrus; Right then Enter picks Canvas.
      await commit(async () => {
        await page.keyboard.press('ArrowRight');
        await page.keyboard.press('Enter');
      });
      assert.equal(await gallery.count(), 0);
      for (const shape of await shapes()) {
        assert.deepEqual(shape, { fill: 'image', texture: 'canvas', layout: TILE });
      }
      assert.equal(await tiled.isChecked(), true);
      await texture.click();
      assert.equal(
        await gallery
          .getByRole('menuitemradio', { name: 'Canvas', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await page.keyboard.press('Escape');
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
      assert.equal(
        await gallery
          .getByRole('menuitemradio', { name: 'Canvas', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await commit(() => gallery.getByRole('menuitemradio', { name: 'Oak', exact: true }).click());
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
          texture: textureIdOf(readSlideBackgroundImageBytes(pres, slide)),
          layout: getSlideBackgroundImageFillLayout(slide),
        }));
      };
      await editor.getByRole('tab', { name: 'Design', exact: true }).click();
      await editor.getByRole('button', { name: 'Background Styles', exact: true }).click();
      await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();
      await insertPixel();
      assert.equal((await backgrounds())[0].texture, null);
      await texture.click();
      assert.deepEqual(await swatchNames(gallery), NAMES);
      await commit(() =>
        gallery.getByRole('menuitemradio', { name: 'Walnut', exact: true }).click(),
      );
      let state = await backgrounds();
      assert.deepEqual(state[0], { texture: 'walnut', layout: TILE });
      assert.equal(state[1].texture, null);
      assert.equal(await tiled.isChecked(), true);
      await texture.click();
      assert.equal(
        await gallery
          .getByRole('menuitemradio', { name: 'Walnut', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await page.keyboard.press('Escape');
      await commit(() => editor.getByRole('button', { name: 'Apply to All', exact: true }).click());
      state = await backgrounds();
      assert.deepEqual(state, [
        { texture: 'walnut', layout: TILE },
        { texture: 'walnut', layout: TILE },
      ]);
      await undo();
      assert.equal((await backgrounds())[1].texture, null);
      await undo();
      assert.equal((await backgrounds())[0].texture, null);
    });
  },
);

test('Texture gallery shows Japanese names', { timeout: 120000 }, async () => {
  await withEditor('ja', async ({ editor }) => {
    await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
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
