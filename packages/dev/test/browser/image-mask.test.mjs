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
  getShapeStroke,
  getShapeStrokeDash,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'image masks render, undo and persist without altering source bytes or crop',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-image-mask-'));
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
      await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
      await editor.locator('#ribbon-panel [data-menu="pictures"]').click();
      await editor.getByTitle(/— addSlideImage$/).click();
      let dialog = editor.getByRole('dialog');
      await dialog
        .getByLabel('Image file', { exact: true })
        .setInputFiles({ name: 'quadrants.png', mimeType: 'image/png', buffer: png });
      await dialog.getByRole('button', { name: 'Insert image', exact: true }).click();
      await dialog.waitFor({ state: 'hidden' });
      await saved();
      const originalBounds = getShapeBounds(await picture());
      await editor.locator('.hit').click({ button: 'right' });
      await editor.getByRole('menuitem', { name: 'Format Picture...', exact: true }).click();
      await editor.getByLabel('Offset X', { exact: true }).fill('0.5');
      await editor.getByLabel('Offset X', { exact: true }).press('Tab');
      await saved();
      const crop = getShapeImageCrop(await picture());
      // Every mask is the preset's own outline, so a new preset changes it.
      const clip = editor.locator('.paint clipPath path').first();
      const nextOutline = async (previous) => {
        for (let i = 0; i < 200; i += 1) {
          const d = await clip.getAttribute('d').catch(() => null);
          if (d && d !== previous) return d;
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        throw new Error('The mask outline did not change');
      };
      let mask = null;
      for (const preset of [
        'ellipse',
        'roundRect',
        'triangle',
        'diamond',
        'pentagon',
        'hexagon',
        'star5',
        'heart',
      ]) {
        await editor.getByLabel('Image shape', { exact: true }).selectOption(preset);
        mask = await nextOutline(mask);
        await saved();
        assert.equal(getShapePreset(await picture()), preset);
        assert.deepEqual(getShapeBounds(await picture()), originalBounds);
        assert.deepEqual(getShapeImageCrop(await picture()), crop);
      }
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor.getByLabel('画像の形状', { exact: true }).selectOption('ellipse');
      await saved();
      await page.screenshot({ path: '/tmp/pptx-pr287-image-mask-ja.png', fullPage: true });
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapePreset(await picture()), 'heart');
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal(getShapePreset(await picture()), 'ellipse');
      assert.deepEqual(Buffer.from(getShapeImageBytes(await picture())), png);
      await page.reload();
      await saved();
      assert.equal(getShapePreset(await picture()), 'ellipse');
      await editor.locator('.hit').click();
      await editor.locator('.hit').click({ button: 'right' });
      await editor.getByRole('menuitem', { name: '図の書式設定...', exact: true }).click();
      await editor.getByLabel('画像の形状', { exact: true }).selectOption('rect');
      await saved();
      assert.equal(getShapePreset(await picture()), 'rect');
      assert.deepEqual(getShapeImageCrop(await picture()), crop);
      await editor.getByLabel('画像の形状', { exact: true }).selectOption('ellipse');
      await editor.getByLabel('画像の枠線の色', { exact: true }).fill('#7c3aed');
      await editor.getByLabel('画像の枠線の太さ（ポイント）', { exact: true }).fill('6');
      await editor.getByLabel('画像の枠線の太さ（ポイント）', { exact: true }).press('Tab');
      await saved();
      assert.equal(getShapeStroke(await picture()).color, '#7C3AED');
      await editor.getByLabel('画像の枠線の種類', { exact: true }).selectOption('dash');
      await saved();
      assert.equal(getShapeStroke(await picture()).widthEmu, 76200);
      assert.equal(getShapeStrokeDash(await picture()), 'dash');
      const outline = editor.locator('.paint g[fill="none"][stroke="#7C3AED"]');
      await outline.waitFor({ state: 'attached' });
      assert.equal(await outline.getAttribute('stroke-width'), '8.00');
      assert.equal(await outline.getAttribute('clip-path'), null);
      assert.equal(await outline.locator('path').count(), 1);
      assert.ok(await outline.getAttribute('stroke-dasharray'));
      await page.screenshot({ path: '/tmp/pptx-pr287-image-border-ja.png', fullPage: true });
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapeStrokeDash(await picture()) ?? 'solid', 'solid');
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.equal(getShapeStrokeDash(await picture()), 'dash');
      await page.reload();
      await saved();
      await editor.locator('.hit').click();
      await editor.locator('.hit').click({ button: 'right' });
      await editor.getByRole('menuitem', { name: '図の書式設定...', exact: true }).click();
      assert.equal(getShapeStroke(await picture()).widthEmu, 76200);
      await editor.getByLabel('画像の枠線の種類', { exact: true }).selectOption('none');
      await saved();
      assert.equal(getShapeStroke(await picture()).kind, 'none');
      assert.equal(await outline.count(), 0);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapeStrokeDash(await picture()), 'dash');
      assert.equal(getShapeStroke(await picture()).widthEmu, 76200);
      assert.deepEqual(Buffer.from(getShapeImageBytes(await picture())), png);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
