import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeBoundsResolved,
  getShapeKind,
  getShapePreset,
  getSlides,
  getSlideShapes,
  getSlideSize,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide} from '@office-kit/pptx-dsl';export default <Presentation><Slide /></Presentation>`;

test(
  'Shapes opens PowerPoint’s gallery and the chosen shape is drawn by dragging or clicking',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-shape-gallery-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url + '/editor');
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const shapes = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return { pres, shapes: getSlideShapes(getSlides(pres)[0]) };
      };
      const changed = async (action) => {
        const before = await (await fetch(preview.url + '/editor/state')).json();
        await action();
        for (let i = 0; i < 200; i += 1) {
          const state = await (await fetch(preview.url + '/editor/state')).json();
          if (!state.building && state.revision !== before.revision) return;
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        throw new Error('The edit was not saved');
      };
      const panel = page.locator('#ribbon-panel');
      const gallery = page.getByRole('menu', { name: 'Shapes', exact: true });

      await panel.getByRole('button', { name: 'Shapes', exact: true }).click();
      assert.deepEqual(await gallery.locator('.heading').allTextContents(), [
        'Lines',
        'Rectangles',
        'Basic Shapes',
        'Block Arrows',
        'Equation Shapes',
        'Stars and Banners',
      ]);
      await gallery.getByRole('menuitem', { name: 'Oval', exact: true }).click();
      const layer = page.getByRole('application', { name: 'Draw shape' });
      const box = await layer.boundingBox();
      await changed(async () => {
        await page.mouse.move(box.x + box.width * 0.25, box.y + box.height * 0.25);
        await page.mouse.down();
        await page.mouse.move(box.x + box.width * 0.5, box.y + box.height * 0.5, { steps: 5 });
        await page.mouse.up();
      });
      let { pres, shapes: drawn } = await shapes();
      const oval = drawn.find((shape) => getShapePreset(shape) === 'ellipse');
      const size = getSlideSize(pres);
      const bounds = getShapeBoundsResolved(pres, oval);
      assert.ok(Math.abs(bounds.x / size.width - 0.25) < 0.02);
      assert.ok(Math.abs(bounds.w / size.width - 0.25) < 0.02);
      assert.equal(await layer.count(), 0);

      // A click drops a one-inch shape; a line is drawn as a connector.
      await panel.getByRole('button', { name: 'Shapes', exact: true }).click();
      await gallery.getByRole('menuitem', { name: 'Line', exact: true }).click();
      await changed(async () => {
        await page.mouse.move(box.x + box.width * 0.6, box.y + box.height * 0.6);
        await page.mouse.down();
        await page.mouse.move(box.x + box.width * 0.8, box.y + box.height * 0.7, { steps: 5 });
        await page.mouse.up();
      });
      ({ shapes: drawn } = await shapes());
      assert.ok(drawn.some((shape) => getShapeKind(shape) === 'connector'));
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
