import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getShapeKind,
  getShapeImageCrop,
  getShapeImageBytes,
  getShapeBounds,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'visual image crop supports handles, movement, keyboard, cancel and bilingual saved history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-image-crop-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide /></Presentation>`,
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
      const picture = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        ).find((shape) => getShapeKind(shape) === 'picture');
      await saved();
      const png = Buffer.from(
        await page.evaluate(() => {
          const canvas = document.createElement('canvas');
          canvas.width = 400;
          canvas.height = 200;
          const ctx = canvas.getContext('2d');
          for (const [index, color] of ['#e45252', '#45b36b', '#387cdc', '#ffcd49'].entries()) {
            ctx.fillStyle = color;
            ctx.fillRect((index % 2) * 200, Math.floor(index / 2) * 100, 200, 100);
          }
          ctx.fillStyle = '#111';
          ctx.font = '28px sans-serif';
          ctx.fillText('Crop / トリミング', 60, 110);
          return canvas.toDataURL('image/png').split(',')[1];
        }),
        'base64',
      );
      await editor.getByRole('button', { name: 'Insert', exact: true }).click();
      await editor.getByTitle(/— addSlideImage$/).click();
      let dialog = editor.getByRole('dialog');
      await dialog
        .getByLabel('Image file', { exact: true })
        .setInputFiles({ name: 'quadrants.png', mimeType: 'image/png', buffer: png });
      await dialog.getByRole('button', { name: 'Insert image', exact: true }).click();
      await dialog.waitFor({ state: 'hidden' });
      await saved();
      const originalBounds = getShapeBounds(await picture());
      await editor.locator('.hit').dblclick();
      dialog = editor.getByRole('dialog', { name: 'Crop image', exact: true });
      const surface = dialog.locator('.crop-surface');
      const drag = async (name, xFraction, yFraction) => {
        const box = await surface.boundingBox();
        const handle = await dialog.getByRole('button', { name, exact: true }).boundingBox();
        const x = handle.x + handle.width / 2,
          y = handle.y + handle.height / 2;
        await page.mouse.move(x, y);
        await page.mouse.down();
        await page.mouse.move(x + box.width * xFraction, y + box.height * yFraction, { steps: 10 });
        await page.mouse.up();
      };
      const rect = () =>
        dialog.locator('.crop-box').evaluate((node) => ({
          left: parseFloat(node.style.left),
          top: parseFloat(node.style.top),
          width: parseFloat(node.style.width),
          height: parseFloat(node.style.height),
        }));
      await dialog.getByRole('button', { name: 'Crop left edge', exact: true }).waitFor();
      await drag('Crop left edge', 0.2, 0);
      assert.ok(Math.abs((await rect()).left - 20) < 0.5);
      assert.equal(getShapeImageCrop(await picture()), null);
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.equal(getShapeImageCrop(await picture()), null);
      await editor.getByRole('button', { name: 'Crop image', exact: true }).click();
      await dialog
        .getByRole('button', { name: 'Crop top left', exact: true })
        .press('Shift+ArrowRight');
      await dialog
        .getByRole('button', { name: 'Crop top left', exact: true })
        .press('Shift+ArrowDown');
      await dialog
        .getByRole('button', { name: 'Crop bottom right', exact: true })
        .press('Shift+ArrowLeft');
      await dialog
        .getByRole('button', { name: 'Crop bottom right', exact: true })
        .press('Shift+ArrowUp');
      await drag('Move crop selection', 0.3, 0.3);
      assert.ok(Math.abs((await rect()).left - 20) < 0.5);
      assert.ok(Math.abs((await rect()).top - 20) < 0.5);
      await dialog.getByRole('button', { name: 'Reset crop', exact: true }).click();
      await dialog
        .getByRole('button', { name: 'Crop top left', exact: true })
        .press('Shift+ArrowRight');
      await dialog
        .getByRole('button', { name: 'Crop top left', exact: true })
        .press('Shift+ArrowDown');
      await dialog
        .getByRole('button', { name: 'Crop bottom right', exact: true })
        .press('Shift+ArrowLeft');
      await dialog
        .getByRole('button', { name: 'Crop bottom right', exact: true })
        .press('Shift+ArrowUp');
      await dialog
        .getByRole('button', { name: 'Move crop selection', exact: true })
        .press('ArrowRight');
      await dialog
        .getByRole('button', { name: 'Move crop selection', exact: true })
        .press('ArrowDown');
      await page.screenshot({ path: '/tmp/pptx-pr287-crop-en.png', fullPage: true });
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await saved();
      const expected = { left: 0.11, top: 0.11, right: 0.09, bottom: 0.09 };
      assert.deepEqual(getShapeImageCrop(await picture()), expected);
      assert.deepEqual(getShapeBounds(await picture()), originalBounds);
      assert.deepEqual(Buffer.from(getShapeImageBytes(await picture())), png);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapeImageCrop(await picture()), null);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(getShapeImageCrop(await picture()), expected);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor.getByRole('button', { name: '画像をトリミング', exact: true }).click();
      dialog = editor.getByRole('dialog', { name: '画像をトリミング', exact: true });
      await dialog.getByRole('button', { name: 'トリミング範囲を移動', exact: true }).waitFor();
      await page.screenshot({ path: '/tmp/pptx-pr287-crop-ja.png', fullPage: true });
      await dialog.getByRole('button', { name: 'トリミングをリセット', exact: true }).click();
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      assert.equal(getShapeImageCrop(await picture()), null);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.deepEqual(getShapeImageCrop(await picture()), expected);
      await editor.locator('.hit').dblclick();
      await dialog.getByRole('button', { name: 'トリミング範囲を移動', exact: true }).waitFor();
      assert.ok(Math.abs((await rect()).left - 11) < 0.01);
      await page.keyboard.press('Escape');
      await dialog.waitFor({ state: 'hidden' });
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-crop-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
