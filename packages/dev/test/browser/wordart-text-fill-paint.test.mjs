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
import { startPreview, waitForState } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Shape} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Shape preset="rect" x={1} y={1} width={6} height={1.5} text="Title" /></Slide></Presentation>`;

// The paint server each glyph run of `text` uses, read from an SVG subtree.
const paintsOf = (root, text) =>
  root.evaluate((node, text) => {
    const doc = node.ownerDocument;
    return [...node.querySelectorAll('tspan')]
      .filter((tspan) => tspan.textContent === text)
      .map((tspan) => {
        const id = /^url\(#(.+)\)$/.exec(tspan.getAttribute('fill') ?? '')?.[1];
        const server = id ? doc.getElementById(id) : null;
        return {
          fill: tspan.getAttribute('fill'),
          kind: server?.localName ?? null,
          stops: server
            ? [...server.querySelectorAll('stop')].map((stop) => [
                stop.getAttribute('offset'),
                stop.getAttribute('stop-color'),
              ])
            : [],
          visible: getComputedStyle(tspan).visibility !== 'hidden',
        };
      });
  }, text);

test(
  'WordArt gradient and pattern fills paint the canvas text and survive editing',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-wordart-text-fill-'));
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
        await waitForState(
          preview.url,
          (state) => !state.building && state.revision !== before.revision,
        );
        await page.getByText('Saved to this project', { exact: true }).waitFor();
      };
      const panel = page.locator('#ribbon-panel');
      const applyPreset = async (name) => {
        await panel.getByRole('button', { name: 'WordArt Quick Styles', exact: true }).click();
        const menu = panel.getByRole('menu', { name: 'WordArt Quick Styles', exact: true });
        await changed(() => menu.getByRole('menuitem', { name, exact: true }).click());
      };
      const shapePaint = page.locator('.paint [data-pptx-shape-id]').first();
      // The static HTML glyphs keep layout and selection but draw nothing.
      const htmlColor = () =>
        shapePaint.evaluate((node) => {
          const span = [...node.querySelectorAll('foreignObject span')].find(
            (el) => el.textContent === 'Title',
          );
          return span ? getComputedStyle(span).color : null;
        });

      await page.locator('.hit').first().click();
      await page.getByRole('tab', { name: 'Shape Format', exact: true }).click();

      await applyPreset('Gradient Fill, Gray');
      await page.waitForFunction(() =>
        [...document.querySelectorAll('.paint tspan')].some((t) =>
          t.getAttribute('fill')?.startsWith('url('),
        ),
      );
      const gradient = await paintsOf(shapePaint, 'Title');
      assert.equal(gradient.length, 1, JSON.stringify(gradient));
      assert.equal(gradient[0].kind, 'linearGradient');
      assert.deepEqual(gradient[0].stops, [
        ['0.2100', '#53575C'],
        ['0.8800', '#C5C7CA'],
      ]);
      assert.equal(await htmlColor(), 'rgba(0, 0, 0, 0)');

      await applyPreset('Pattern Fill: White; Dark Upward Diagonal Stripe; Shadow');
      await page.waitForFunction(() =>
        [...document.querySelectorAll('.paint tspan')].some((t) => {
          const id = /^url\(#(.+)\)$/.exec(t.getAttribute('fill') ?? '')?.[1];
          return id && document.getElementById(id)?.localName === 'pattern';
        }),
      );
      const pattern = await paintsOf(shapePaint, 'Title');
      assert.ok(
        pattern.some((paint) => paint.kind === 'pattern'),
        JSON.stringify(pattern),
      );
      assert.equal(await htmlColor(), 'rgba(0, 0, 0, 0)');

      // While editing, the overlay paints the pattern and the editable glyphs
      // stay transparent with a visible caret.
      const input = page.locator('.canvas-shell .inline-edit').first();
      // Selecting a text box already places the caret in it.
      if (!(await input.isVisible())) await page.locator('.hit').first().dblclick();
      await input.waitFor();
      const overlay = page.locator('.inline-effects').first();
      await overlay.waitFor();
      assert.ok(
        (await paintsOf(overlay, 'Title')).some((paint) => paint.kind === 'pattern'),
        'the editing overlay paints the pattern',
      );
      const editing = await input.evaluate((node) => {
        const span = [...node.querySelectorAll('span')].find((el) => el.textContent === 'Title');
        const css = span ? getComputedStyle(span) : null;
        return { color: css?.color, caret: css?.caretColor };
      });
      assert.equal(editing.color, 'rgba(0, 0, 0, 0)');
      assert.notEqual(editing.caret, 'rgba(0, 0, 0, 0)');
      // Static glyphs hide while the editable layer is up.
      assert.ok((await paintsOf(shapePaint, 'Title')).every((paint) => !paint.visible));

      const beforeEdit = (await waitForState(preview.url, () => true)).revision;
      await input.fill('Patterned');
      await page.waitForFunction(() =>
        [...document.querySelectorAll('.inline-effects tspan')].some(
          (t) => t.textContent === 'Patterned' && t.getAttribute('fill')?.startsWith('url('),
        ),
      );
      await page.keyboard.press('ControlOrMeta+Enter');
      await waitForState(preview.url, (state) => state.revision > beforeEdit);
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const [saved] = await shapes();
      assert.equal(getShapeText(saved), 'Patterned');
      const format = getShapeRunFormat(saved, 0, 0);
      assert.equal(format.textFill?.kind, 'pattern');
      assert.equal(format.textFill?.preset, 'dkUpDiag');
      await page.waitForFunction(() =>
        [...document.querySelectorAll('.paint tspan')].some(
          (t) => t.textContent === 'Patterned' && t.getAttribute('fill')?.startsWith('url('),
        ),
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
