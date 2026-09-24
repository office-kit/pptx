import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeGradientFill,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'gradient type changes reset native geometry while retaining stops and rotation',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-gradient-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={3} height={2} fill={{stops:[{offset:0,color:'accent1',brightness:0.95},{offset:1,color:'accent1',brightness:0.7}],angleDeg:90,scaled:false}}>Gradient</Text><Text x={6} y={1} width={3} height={2} fill="#00FF00">Solid</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const gradient = async (index = 0) => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeGradientFill(getSlideShapes(getSlides(deck)[0])[index]);
      };
      await saved();
      await editor.locator('.hit').nth(0).click();
      const initialGradient = await gradient();
      const type = editor.getByRole('combobox', { name: 'Gradient type', exact: true });
      const zero = { left: 0, top: 0, right: 0, bottom: 0 };
      const corner = { left: 1, top: 1, right: 0, bottom: 0 };
      const center = { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 };
      const expanded = { left: 0, top: 0, right: -1, bottom: -1 };
      for (const path of ['circle', 'rect', 'circle', 'linear', 'shape']) {
        const before = await gradient();
        await type.selectOption(path);
        await saved();
        const changed = await gradient();
        assert.equal(changed.path ?? 'linear', path);
        assert.deepEqual(changed.stops, initialGradient.stops);
        assert.equal(changed.rotateWithShape, initialGradient.rotateWithShape);
        if (path === 'linear') {
          assert.equal(changed.angleDeg, 45);
          assert.equal(changed.scaled, true);
          assert.deepEqual(changed.tileRect, zero);
          assert.equal(changed.focus, undefined);
        } else {
          assert.deepEqual(changed.focus, path === 'shape' ? center : corner);
          assert.deepEqual(changed.tileRect, path === 'shape' ? zero : expanded);
        }
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await saved();
        assert.deepEqual(await gradient(), before);
        await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
        await saved();
        assert.deepEqual(await gradient(), changed);
        if (path === 'circle' || path === 'rect') {
          await editor.getByRole('button', { name: 'Gradient direction', exact: true }).click();
          await editor
            .getByRole('menuitemradio', { name: 'From Top Left Corner', exact: true })
            .click();
          await saved();
          assert.deepEqual((await gradient()).focus, { left: 0, top: 0, right: 1, bottom: 1 });
        }
      }
      const last = await gradient();
      await page.reload();
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      await saved();
      await editor.locator('.hit').nth(0).click();
      assert.deepEqual(await gradient(), last);
      assert.equal(
        await editor.getByRole('button', { name: 'Gradient direction', exact: true }).isDisabled(),
        true,
      );
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
