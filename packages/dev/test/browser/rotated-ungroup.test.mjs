import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getShapeBounds,
  getShapeRotation,
  getShapeKind,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'rotated groups keep child geometry when ungrouped in English and Japanese',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-rotated-ungroup-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={2} width={2} height={1}>English</Text><Text x={5} y={2} width={2} height={1} rotation={30}>日本語</Text></Slide></Presentation>`,
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
      const shapes = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        );
      await saved();
      await editor.locator('.hit').first().click();
      await page.keyboard.press('Control+a');
      await editor.getByRole('tab', { name: 'Size & Properties' }).click();
      await editor
        .getByRole('region', { name: 'Arrange', exact: true })
        .getByRole('button', { name: 'Group', exact: true })
        .click();
      await saved();
      assert.equal(await editor.locator('.hit').count(), 1);
      const frame = await editor.locator('.hit.selected').boundingBox();
      const button = await editor
        .getByRole('button', { name: 'Rotate', exact: true })
        .boundingBox();
      const cx = frame.x + frame.width / 2,
        cy = frame.y + frame.height / 2;
      const x = button.x + button.width / 2,
        y = button.y + button.height / 2;
      await page.mouse.move(x, y);
      await page.mouse.down();
      await page.keyboard.down('Shift');
      await page.mouse.move(cx - (y - cy), cy + (x - cx), { steps: 8 });
      await page.mouse.up();
      await page.keyboard.up('Shift');
      await saved();
      const grouped = await shapes();
      assert.equal(getShapeRotation(grouped[0]), 90);
      const verify = async () => {
        const result = await shapes();
        assert.equal(result.length, 2);
        for (const [index, shape] of result.entries()) {
          const bounds = getShapeBounds(shape);
          assert.ok(Math.abs(bounds.x - 3 * 914400) <= 2);
          assert.ok(Math.abs(bounds.y - index * 4 * 914400) <= 2);
          assert.equal(bounds.w, 2 * 914400);
          assert.equal(bounds.h, 914400);
          assert.equal(getShapeRotation(shape), index ? 120 : 90);
        }
      };
      await editor
        .getByRole('region', { name: 'Arrange', exact: true })
        .getByRole('button', { name: 'Ungroup', exact: true })
        .click();
      await saved();
      await verify();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapeKind((await shapes())[0]), 'group');
      assert.equal(await editor.locator('.hit').count(), 1);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor
        .getByRole('region', { name: '配置', exact: true })
        .getByRole('button', { name: 'グループ解除', exact: true })
        .click();
      await saved();
      await verify();
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await verify();
      await page.reload();
      await saved();
      await verify();
      assert.equal(await editor.locator('.hit').count(), 2);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
