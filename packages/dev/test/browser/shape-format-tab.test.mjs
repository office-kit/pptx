import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  cm,
  getShapeBoundsResolved,
  getShapePreset,
  getShapeRunFormat,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={6} height={1} text="Title" /></Slide></Presentation>`;

test(
  'Shape Format tab matches the reference desktop app: Edit Shape, WordArt styles, text effects, size and Format Pane',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-shape-format-'));
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
      const shape = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const found = getSlideShapes(getSlides(pres)[0])[0];
        return { pres, shape: found };
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
      await page.locator('.hit').first().click();
      await page.getByRole('tab', { name: 'Shape Format', exact: true }).click();
      const panel = page.locator('#ribbon-panel');
      const groups = await panel
        .locator('.shape-format > section')
        .evaluateAll((nodes) => nodes.map((node) => node.getAttribute('aria-label')));
      assert.deepEqual(groups, [
        'Insert Shapes',
        'Shape Styles',
        'WordArt Styles',
        'Accessibility',
        'Arrange',
        'Size',
        'Format Pane',
      ]);
      assert.equal(await panel.getByRole('button', { name: 'Merge Shapes' }).isDisabled(), true);

      await panel.getByRole('button', { name: 'Edit Shape', exact: true }).click();
      await changed(() => panel.getByRole('menuitem', { name: 'Oval', exact: true }).click());
      assert.equal(getShapePreset((await shape()).shape), 'ellipse');

      await panel.getByRole('button', { name: 'Text Effects', exact: true }).click();
      await changed(() => panel.getByRole('menuitem', { name: 'Glow', exact: true }).click());
      assert.ok(getShapeRunFormat((await shape()).shape, 0, 0).glow);

      await panel.getByRole('button', { name: 'WordArt Quick Styles', exact: true }).click();
      await changed(() =>
        panel
          .getByRole('menuitem', {
            name: 'Fill: Red, Accent color 2; Outline: Red, Accent color 2',
          })
          .click(),
      );
      const styled = getShapeRunFormat((await shape()).shape, 0, 0);
      assert.equal(styled.glow, undefined);
      assert.ok(styled.outline);

      const width = panel.getByRole('spinbutton', { name: 'Width', exact: true });
      await width.fill('8');
      await changed(() => width.press('Enter'));
      const { pres, shape: resized } = await shape();
      assert.equal(Math.round(getShapeBoundsResolved(pres, resized).w / cm(1)), 8);

      await panel.getByRole('button', { name: 'Format Pane', exact: true }).click();
      await page.getByRole('button', { name: 'Close Format Shape', exact: true }).waitFor();
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
