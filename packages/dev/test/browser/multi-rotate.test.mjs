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
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

test(
  'multiple objects rotate around their own centres, with bilingual history and cancellation',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-multi-rotate-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={2} y={2} width={2} height={1}>English</Text><Text x={6} y={2} width={2} height={1} rotation={30}>日本語</Text><Text x={10} y={5} width={1} height={1}>Other</Text></Slide></Presentation>`,
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
      const geometry = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        ).map((shape) => ({ ...getShapeBounds(shape), rotation: getShapeRotation(shape) }));
      await saved();
      const original = await geometry();
      await editor.locator('.hit').nth(0).click();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      assert.equal(await editor.locator('.hit.selected').count(), 2);
      const handle = () =>
        editor
          .locator('.hit')
          .nth(0)
          .getByRole('button', {
            name: ja ? '回転' : 'Rotate',
            exact: true,
          });
      assert.equal(await handle().count(), 1, 'each selected object has its own rotation handle');
      const beginTurn = async (degrees, single = false) => {
        const frame = await editor.locator('.hit.selected').first().boundingBox();
        const button = await (
          single ? editor.getByRole('button', { name: '回転', exact: true }) : handle()
        ).boundingBox();
        const cx = frame.x + frame.width / 2,
          cy = frame.y + frame.height / 2;
        const x = button.x + button.width / 2 + (single ? 4 : 0),
          y = button.y + button.height / 2;
        const radians = (degrees * Math.PI) / 180;
        await page.mouse.move(x, y);
        await page.mouse.down();
        if (!single) await page.keyboard.down('Shift');
        await page.mouse.move(
          cx + (x - cx) * Math.cos(radians) - (y - cy) * Math.sin(radians),
          cy + (x - cx) * Math.sin(radians) + (y - cy) * Math.cos(radians),
          { steps: 8 },
        );
      };
      const beforeTurn = (await waitForState(preview.url, () => true)).revision;
      await beginTurn(89);
      await page.mouse.up();
      await page.keyboard.up('Shift');
      await waitForState(preview.url, (state) => state.revision !== beforeTurn);
      await saved();
      const turned = await geometry();
      assert.equal(turned[0].rotation, 90);
      assert.equal(turned[1].rotation, 120);
      assert.deepEqual(turned[2], original[2]);
      for (let i = 0; i < 2; i++) {
        assert.equal(turned[i].x, original[i].x);
        assert.equal(turned[i].y, original[i].y);
        assert.equal(turned[i].w, original[i].w);
        assert.equal(turned[i].h, original[i].h);
      }
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await geometry(), original);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor.locator('.hit').nth(0).click();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await beginTurn(-44);
      await page.keyboard.press('Escape');
      await page.mouse.up();
      await page.keyboard.up('Shift');
      await saved();
      assert.deepEqual(await geometry(), original);
      assert.equal(await editor.locator('.hit.selected').count(), 2);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await geometry(), turned);
      await page.reload();
      await saved();
      assert.deepEqual(await geometry(), turned);
      await editor
        .locator('.hit')
        .nth(2)
        .click({ position: { x: 2, y: 2 } });
      await beginTurn(45, true);
      await page.mouse.up();
      await saved();
      const singleTurn = await geometry();
      assert.equal(
        singleTurn[2].rotation,
        45,
        'grabbing the edge of the handle does not offset the rotation',
      );
      assert.deepEqual({ ...singleTurn[2], rotation: 0 }, original[2]);
      assert.deepEqual(singleTurn.slice(0, 2), turned.slice(0, 2));
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
