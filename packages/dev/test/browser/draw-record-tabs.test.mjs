import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeCustomGeometry,
  getShapeName,
  getShapePreset,
  getShapeStrokeColor,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={1} height={1} text="One" /></Slide><Slide><Shape preset="rect" x={1} y={1} width={3} height={1} text="Two" /></Slide></Presentation>`;

test(
  'Draw and Record tabs ink, erase, recognize, lasso and play like PowerPoint',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-draw-tab-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(file, DECK);
    let preview;
    let browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const shapes = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        );
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
      const panel = editor.locator('#ribbon-panel');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const stage = await editor.locator('.stage').boundingBox();
      // Slide fractions → page coordinates.
      const at = (fx, fy) => ({ x: stage.x + stage.width * fx, y: stage.y + stage.height * fy });
      const stroke = async (points) => {
        await page.mouse.move(at(...points[0]).x, at(...points[0]).y);
        await page.mouse.down();
        for (const point of points.slice(1))
          await page.mouse.move(at(...point).x, at(...point).y, { steps: 4 });
        await page.mouse.up();
      };

      await editor.getByRole('tab', { name: 'Draw', exact: true }).click();
      await panel.getByRole('radio', { name: 'Pen: Red, 1 mm', exact: true }).click();
      await changed(() =>
        stroke([
          [0.5, 0.5],
          [0.6, 0.55],
          [0.7, 0.5],
          [0.8, 0.6],
        ]),
      );
      let ink = (await shapes()).find((shape) => getShapeName(shape).startsWith('Ink '));
      assert.ok(ink, 'the stroke is saved as an ink freeform');
      assert.equal(getShapeStrokeColor(ink), '#C00000');
      assert.equal(getShapeCustomGeometry(ink).paths[0].fill, 'none');

      // The eraser removes a stroke it touches and nothing else.
      await panel.getByRole('button', { name: 'Eraser', exact: true }).click();
      await changed(() =>
        stroke([
          [0.6, 0.45],
          [0.6, 0.6],
        ]),
      );
      ink = (await shapes()).find((shape) => getShapeName(shape).startsWith('Ink '));
      assert.equal(ink, undefined);
      assert.equal((await shapes()).length, 1);

      // Ink to Shape turns a closed four-corner stroke into a rectangle.
      await panel.getByRole('button', { name: 'Ink to Shape', exact: true }).click();
      await changed(() =>
        stroke([
          [0.5, 0.3],
          [0.8, 0.3],
          [0.8, 0.7],
          [0.5, 0.7],
          [0.5, 0.31],
        ]),
      );
      const recognized = (await shapes()).at(-1);
      assert.equal(getShapePreset(recognized), 'rect');
      assert.equal(getShapeCustomGeometry(recognized), null);
      await panel.getByRole('button', { name: 'Ink to Shape', exact: true }).click();

      // Lasso Select picks the shapes wholly inside the loop and returns to selecting.
      await panel.getByRole('button', { name: 'Lasso Select', exact: true }).click();
      await stroke([
        [0.45, 0.25],
        [0.85, 0.25],
        [0.85, 0.75],
        [0.45, 0.75],
        [0.45, 0.26],
      ]);
      await editor.locator('[data-ink-tool]').waitFor({ state: 'detached' });
      assert.equal(await editor.locator('.hit.selected').count(), 1);
      assert.equal(
        await panel
          .getByRole('button', { name: 'Lasso Select', exact: true })
          .getAttribute('aria-pressed'),
        'false',
      );

      await editor.getByRole('tab', { name: 'Record', exact: true }).click();
      assert.equal(
        await panel.getByRole('button', { name: 'Clear Recording', exact: true }).isDisabled(),
        true,
      );
      await editor.locator('.thumb').nth(1).click();
      await panel.getByRole('button', { name: 'From Current Slide', exact: true }).click();
      await page.waitForFunction(
        () => document.querySelector('#present-count')?.textContent === 'Slide 2 of 2',
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
