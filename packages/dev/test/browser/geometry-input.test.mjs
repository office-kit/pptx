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
  inches,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'geometry inputs preserve untouched precision and reject invalid values in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-geometry-input-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1.23456} y={1.34567} width={3.45678} height={1.56789}>日本語 English</Text></Slide></Presentation>`,
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
      const read = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const shape = getSlideShapes(getSlides(deck)[0])[0];
        return { bounds: getShapeBounds(shape), rotation: getShapeRotation(shape) };
      };
      await saved();
      await editor.locator('.hit').first().click();
      const original = await read();
      const field = (name) =>
        editor.locator('.bespoke').getByRole('spinbutton', { name, exact: true });
      const change = async (name, value) => {
        await field(name).fill(value);
        await field(name).press('Tab');
      };
      await change('X', '2.125');
      await saved();
      const moved = { ...original, bounds: { ...original.bounds, x: inches(2.125) } };
      assert.deepEqual(await read(), moved);
      for (const [name, value] of [
        ['W', '-1'],
        ['H', ''],
        ['X', '1e20'],
        ['Y', ''],
      ]) {
        const before = await field(name).inputValue();
        await change(name, value);
        assert.equal(await field(name).inputValue(), before);
        assert.deepEqual(await read(), moved);
      }
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), original);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), moved);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await change('Y', '-0.125');
      await saved();
      await change('回転', '-30.5');
      await saved();
      const rotated = await read();
      assert.deepEqual(rotated.bounds, { ...moved.bounds, y: inches(-0.125) });
      assert.equal(rotated.rotation, 329.5);
      await change('回転', '');
      assert.equal(await field('回転').inputValue(), '329.5');
      assert.deepEqual(await read(), rotated);
      await page.reload();
      await saved();
      assert.deepEqual(await read(), rotated);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
