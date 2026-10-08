import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getShapeText,
  getShapeText3D,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={6} height={1.5} text="Title" /></Slide></Presentation>`;

// The bevel layer drawing `text` inside an SVG subtree, with its filter's
// lighting primitives.
const bevelsOf = (root, text) =>
  root.evaluate((node, text) => {
    const doc = node.ownerDocument;
    return [...node.querySelectorAll('text[data-pptx-bevel]')]
      .filter((el) => el.textContent === text)
      .map((el) => {
        const id = /^url\(#(.+)\)$/.exec(el.getAttribute('filter') ?? '')?.[1];
        const filter = id ? doc.getElementById(id) : null;
        const light = filter?.querySelector('feDiffuseLighting feDistantLight');
        return {
          filter: filter?.localName ?? null,
          blur: filter?.querySelector('feGaussianBlur')?.getAttribute('stdDeviation') ?? null,
          surfaceScale:
            filter?.querySelector('feDiffuseLighting')?.getAttribute('surfaceScale') ?? null,
          azimuth: light?.getAttribute('azimuth') ?? null,
          elevation: light?.getAttribute('elevation') ?? null,
          visible: getComputedStyle(el).visibility !== 'hidden',
        };
      });
  }, text);

test(
  'Text Art Soft Bevel shades the canvas text and survives editing',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-text-art-bevel-'));
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
      const shapePaint = page.locator('.paint [data-pptx-shape-id]').first();
      assert.deepEqual(await bevelsOf(shapePaint, 'Title'), []);

      await page.locator('.hit').first().click();
      await page.getByRole('tab', { name: 'Shape Format', exact: true }).click();
      const panel = page.locator('#ribbon-panel');
      await panel.getByRole('button', { name: 'Text Art Quick Styles', exact: true }).click();
      const menu = panel.getByRole('menu', { name: 'Text Art Quick Styles', exact: true });
      const before = await (await fetch(preview.url + '/editor/state')).json();
      await menu
        .getByRole('menuitem', { name: 'Fill: Purple, Accent color 4; Soft Bevel', exact: true })
        .click();
      await waitForState(
        preview.url,
        (state) => !state.building && state.revision !== before.revision,
      );
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      await page.waitForFunction(() => document.querySelector('.paint text[data-pptx-bevel]'));

      // Soft Bevel: w 25400 EMU (blur = half of 2.67 px), h 38100 EMU, soft
      // rig from the top turned by its 260° revolution.
      const [bevel] = await bevelsOf(shapePaint, 'Title');
      assert.equal(bevel?.filter, 'filter', JSON.stringify(bevel));
      assert.equal(bevel.blur, '1.33');
      assert.equal(bevel.azimuth, '170');
      assert.equal(bevel.elevation, '45');
      assert.ok(Number(bevel.surfaceScale) > 0);

      const input = page.locator('.canvas-shell .inline-edit').first();
      if (!(await input.isVisible())) await page.locator('.hit').first().dblclick();
      await input.waitFor();
      const overlay = page.locator('.inline-effects').first();
      await overlay.waitFor();
      assert.equal((await bevelsOf(overlay, 'Title')).length, 1, 'the editing overlay bevels');
      // The static bevel hides with the static glyphs while editing.
      assert.ok((await bevelsOf(shapePaint, 'Title')).every((layer) => !layer.visible));

      const beforeEdit = (await waitForState(preview.url, () => true)).revision;
      await input.fill('Beveled');
      await page.waitForFunction(() =>
        [...document.querySelectorAll('.inline-effects text[data-pptx-bevel]')].some(
          (t) => t.textContent === 'Beveled',
        ),
      );
      await page.keyboard.press('ControlOrMeta+Enter');
      await waitForState(preview.url, (state) => state.revision > beforeEdit);
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const [saved] = await shapes();
      assert.equal(getShapeText(saved), 'Beveled');
      assert.deepEqual(getShapeText3D(saved)?.bevelTop, { widthEmu: 25400, heightEmu: 38100 });
      await page.waitForFunction(() =>
        [...document.querySelectorAll('.paint text[data-pptx-bevel]')].some(
          (t) => t.textContent === 'Beveled',
        ),
      );
      if (process.env.BEVEL_SCREENSHOT) {
        await page
          .locator('.hit')
          .first()
          .click({ position: { x: 1, y: 1 } });
        await page.keyboard.press('Escape');
        await shapePaint.screenshot({ path: process.env.BEVEL_SCREENSHOT });
      }
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
