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
import { startPreview } from '../helpers/server.mjs';

test(
  'Escape and interrupted gestures restore geometry while preserving redo history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-gesture-cancel-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={2} y={2} width={3} height={1}>Move / 移動</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const shape = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        )[0];
      const hit = editor.locator('.hit');
      await saved();
      const original = getShapeBounds(await shape());
      await hit.click();
      await page.keyboard.press('ArrowRight');
      await saved();
      const nudged = getShapeBounds(await shape());
      assert.notDeepEqual(nudged, original);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      await hit.click();
      for (const mode of ['move', 'resize', 'rotate', 'interrupt', 'blur']) {
        const style = await hit.getAttribute('style');
        const target =
          mode === 'resize'
            ? editor.getByRole('button', { name: 'Resize se', exact: true })
            : mode === 'rotate'
              ? editor.getByRole('button', { name: 'Rotate', exact: true })
              : hit;
        const box = await target.boundingBox();
        const x = box.x + box.width / 2,
          y = box.y + box.height / 2;
        await page.mouse.move(x, y);
        await page.mouse.down();
        await page.mouse.move(x + 65, y + 45, { steps: 8 });
        assert.notEqual(await hit.getAttribute('style'), style);
        if (mode === 'interrupt')
          await hit.dispatchEvent('pointercancel', { pointerId: 1, bubbles: true });
        else if (mode === 'blur') await hit.evaluate(() => window.dispatchEvent(new Event('blur')));
        else await page.keyboard.press('Escape');
        await page.mouse.up();
        await hit.and(editor.locator('.selected')).waitFor();
        await hit.evaluate(
          (node, expected) =>
            new Promise((resolve, reject) => {
              const deadline = performance.now() + 5000;
              const check = () => {
                if (node.getAttribute('style') === expected) resolve();
                else if (performance.now() > deadline)
                  reject(new Error('Canceled gesture was not restored'));
                else requestAnimationFrame(check);
              };
              check();
            }),
          style,
        );
        await saved();
        assert.deepEqual(getShapeBounds(await shape()), original);
        assert.equal(getShapeRotation(await shape()) ?? 0, 0);
        assert.equal(
          await editor
            .getByTitle(ja ? 'やり直し (Ctrl+Y)' : 'Redo (Ctrl+Y)', { exact: true })
            .isEnabled(),
          true,
        );
      }
      const stage = await editor.locator('.stage').boundingBox();
      await page.mouse.move(stage.x + 8, stage.y + 8);
      await page.mouse.down();
      await page.mouse.move(stage.x + 65, stage.y + 45, { steps: 4 });
      assert.equal(await editor.locator('.hit.selected').count(), 0);
      await page.keyboard.press('Escape');
      await page.mouse.up();
      assert.equal(await editor.locator('.hit.selected').count(), 1);
      assert.equal(await editor.locator('.marquee').count(), 0);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(getShapeBounds(await shape()), nudged);
      await page.reload();
      await saved();
      assert.deepEqual(getShapeBounds(await shape()), nudged);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
