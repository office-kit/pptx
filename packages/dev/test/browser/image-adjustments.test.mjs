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
  getShapePreset,
  getShapeImageOpacity,
  getShapeImageBrightness,
  getShapeImageContrast,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'image adjustments reset together, undo and persist in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-image-adjustments-'));
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
      const reset = () =>
        editor.getByRole('button', {
          name: ja ? '画像の調整をリセット' : 'Reset image adjustments',
          exact: true,
        });
      assert.equal(await reset().isDisabled(), true);
      const change = async (label, value) => {
        await editor.getByLabel(label, { exact: true }).fill(value);
        await editor.getByLabel(label, { exact: true }).press('Tab');
        await saved();
      };
      await change('Crop left (%)', '10');
      await editor.getByLabel('Image shape', { exact: true }).selectOption('ellipse');
      await saved();
      const before = await picture();
      await change('Opacity (%)', '40');
      await change('Brightness (%)', '20');
      await change('Contrast (%)', '-30');
      const adjustments = (shape) => [
        getShapeImageOpacity(shape) ?? 1,
        getShapeImageBrightness(shape) ?? 0,
        getShapeImageContrast(shape) ?? 0,
      ];
      assert.deepEqual(adjustments(await picture()), [0.4, 0.2, -0.3]);
      await reset().click();
      await saved();
      assert.deepEqual(adjustments(await picture()), [1, 0, 0]);
      assert.equal(await reset().isDisabled(), true);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(adjustments(await picture()), [0.4, 0.2, -0.3]);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await reset().click();
      await saved();
      assert.deepEqual(adjustments(await picture()), [1, 0, 0]);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(adjustments(await picture()), [0.4, 0.2, -0.3]);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      await editor.locator('.hit').click();
      assert.equal(await reset().isDisabled(), true);
      const after = await picture();
      assert.deepEqual(adjustments(after), [1, 0, 0]);
      assert.deepEqual(getShapeBounds(after), getShapeBounds(before));
      assert.deepEqual(getShapeImageCrop(after), getShapeImageCrop(before));
      assert.equal(getShapePreset(after), 'ellipse');
      assert.deepEqual(Buffer.from(getShapeImageBytes(after)), png);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
