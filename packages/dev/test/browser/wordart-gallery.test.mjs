import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeRunFormat,
  getShapeText,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={6} height={1} text="Title" /></Slide></Presentation>`;

// Mac PowerPoint's gallery, in order (test/fixtures/native/wordart-capture.md).
const PRESETS = [
  'Fill: Black, Text color 1; Shadow',
  'Fill: Blue, Accent color 1; Shadow',
  'Fill: Red, Accent color 2; Outline: Red, Accent color 2',
  'Fill: White; Outline: Aqua, Accent color 5; Shadow',
  'Gradient Fill, Gray',
  'Fill: Purple, Accent color 4; Soft Bevel',
  'Gradient Fill: Aqua, Accent color 5; Reflection',
  'Gradient Fill: Purple, Accent color 4; Outline: Purple, Accent color 4',
  'Fill: White; Outline: Blue, Accent color 1; Glow: Blue, Accent color 1',
  'Fill: Olive Green, Accent color 3; Sharp Bevel',
  'Fill: Black, Text color 1; Outline: White, Background color 1; Hard Shadow: White, Background color 1',
  'Fill: Black, Text color 1; Outline: White, Background color 1; Hard Shadow: Aqua, Accent color 5',
  'Fill: Aqua, Accent color 5; Outline: White, Background color 1; Hard Shadow: Aqua, Accent color 5',
  'Fill: White; Outline: Red, Accent color 2; Hard Shadow: Red, Accent color 2',
  'Fill: Tan, Background color 2; Inner Shadow',
  'Pattern Fill: White; Dark Upward Diagonal Stripe; Shadow',
  'Pattern Fill: Olive Green, Accent color 3, Narrow Horizontal Stripe; Inner Shadow',
  'Pattern Fill: Blue, Accent color 1, 50%; Hard Shadow: Blue, Accent color 1',
  'Pattern Fill: Aqua, Accent color 5, Light Downward Diagonal Stripe; Outline: Aqua, Accent color 5',
  'Pattern Fill: Dark Blue, Dark Upward Diagonal Stripe; Hard Shadow',
];
const BEVELS = [
  'Fill: Purple, Accent color 4; Soft Bevel',
  'Fill: Olive Green, Accent color 3; Sharp Bevel',
];

test(
  'WordArt gallery matches PowerPoint: twenty swatches in Quick Styles and Insert ▸ WordArt',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-wordart-gallery-'));
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
      const swatches = (menu) =>
        menu
          .locator('.grid [role="menuitem"]')
          .evaluateAll((nodes) =>
            nodes.map((node) => [node.getAttribute('aria-label'), node.getAttribute('title')]),
          );

      await page.locator('.hit').first().click();
      await page.getByRole('tab', { name: 'Shape Format', exact: true }).click();
      const panel = page.locator('#ribbon-panel');
      await panel.getByRole('button', { name: 'WordArt Quick Styles', exact: true }).click();
      const quickStyles = panel.getByRole('menu', { name: 'WordArt Quick Styles', exact: true });
      const listed = await swatches(quickStyles);
      assert.deepEqual(
        listed.map(([label]) => label),
        PRESETS,
      );
      for (const [label, title] of listed)
        assert.equal(
          title,
          BEVELS.includes(label) ? 'Text bevels are not supported by the library yet.' : label,
        );
      for (const bevel of BEVELS)
        assert.equal(
          await quickStyles.getByRole('menuitem', { name: bevel, exact: true }).isDisabled(),
          true,
        );
      assert.equal(
        await quickStyles.getByRole('menuitem', { name: 'Clear WordArt', exact: true }).count(),
        1,
      );

      // Arrow keys move through the 5-column grid, stepping over the disabled bevels.
      const focused = () =>
        page.evaluate(
          () =>
            document.activeElement?.getAttribute('aria-label') ??
            document.activeElement?.textContent,
        );
      assert.equal(await focused(), PRESETS[0]);
      await page.keyboard.press('ArrowRight');
      assert.equal(await focused(), PRESETS[1]);
      await page.keyboard.press('ArrowLeft');
      await page.keyboard.press('ArrowDown');
      assert.equal(await focused(), PRESETS[10]);
      await page.keyboard.press('ArrowDown');
      await page.keyboard.press('ArrowDown');
      assert.equal(await focused(), 'Clear WordArt');
      await page.keyboard.press('ArrowUp');
      assert.equal(await focused(), PRESETS[15]);
      await page.keyboard.press('Home');
      assert.equal(await focused(), PRESETS[0]);

      await changed(() =>
        quickStyles
          .getByRole('menuitem', { name: 'Fill: White; Outline: Aqua, Accent color 5; Shadow' })
          .click(),
      );
      const white = getShapeRunFormat((await shapes())[0], 0, 0);
      assert.equal(white.bold, true);
      assert.equal(white.color, '#FFFFFF');
      assert.deepEqual(white.outline, { color: 'accent5', widthEmu: 10160 });
      assert.equal(white.shadow.opacity, 0.3);

      // A second preset replaces the first, as PowerPoint does.
      await panel.getByRole('button', { name: 'WordArt Quick Styles', exact: true }).click();
      await changed(() =>
        quickStyles
          .getByRole('menuitem', { name: 'Fill: Black, Text color 1; Shadow', exact: true })
          .click(),
      );
      const black = getShapeRunFormat((await shapes())[0], 0, 0);
      assert.equal(black.color, undefined);
      assert.equal(black.bold, false);
      assert.deepEqual(black.outline, { widthEmu: 0 });
      assert.equal(black.shadow.opacity, 0.4);

      await changed(() => page.getByRole('button', { name: 'Undo', exact: true }).click());
      assert.equal(getShapeRunFormat((await shapes())[0], 0, 0).color, '#FFFFFF');

      // Insert ▸ WordArt opens the same gallery and inserts "Your text here" in the pick.
      await page.getByRole('tab', { name: 'Insert', exact: true }).click();
      await panel.getByRole('button', { name: 'WordArt', exact: true }).click();
      const insertGallery = page.getByRole('menu', { name: 'WordArt', exact: true });
      assert.deepEqual(
        (await swatches(insertGallery)).map(([label]) => label),
        PRESETS,
      );
      assert.equal(await insertGallery.getByRole('menuitem', { name: 'Clear WordArt' }).count(), 0);
      await changed(() =>
        insertGallery.getByRole('menuitem', { name: 'Gradient Fill, Gray', exact: true }).click(),
      );
      assert.equal(await insertGallery.count(), 0);
      const inserted = (await shapes()).find((shape) => getShapeText(shape) === 'Your text here');
      assert.ok(inserted);
      const format = getShapeRunFormat(inserted, 0, 0);
      assert.equal(format.size, 54);
      assert.equal(format.textFill.kind, 'gradient');
      assert.deepEqual(
        format.textFill.stops.map((stop) => [stop.offset, stop.color]),
        [
          [0.21, '#53575C'],
          [0.88, '#C5C7CA'],
        ],
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
