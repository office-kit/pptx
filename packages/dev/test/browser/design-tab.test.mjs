import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getPresentationFonts,
  getPresentationTheme,
  getSlideSize,
  loadPresentation,
  SLIDE_SIZE_4_3,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={4} height={1}>Hello</Text></Slide></Presentation>`;

test(
  'Design tab matches PowerPoint: Themes gallery, Colors, Fonts, Slide Size and Design Suggestions',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-design-tab-'));
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
      const pres = async () =>
        loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
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
      await page.getByRole('tab', { name: 'Design', exact: true }).click();
      const panel = page.locator('#ribbon-panel');
      // The ribbon panel scrolls horizontally, which also clips vertically: a
      // dropdown must open below its button over the slide, not inside the ribbon.
      const assertMenuUnclipped = async (label) => {
        const menu = panel.getByRole('menu', { name: label, exact: true });
        await menu.waitFor();
        const layout = await menu.evaluate((node) => {
          const ribbon = document.getElementById('ribbon-panel');
          const box = node.getBoundingClientRect();
          return {
            ribbonGrew: ribbon.scrollHeight > ribbon.clientHeight,
            below: box.top >= ribbon.getBoundingClientRect().bottom - 40,
            topmost:
              document
                .elementFromPoint(box.left + 10, box.bottom - 10)
                ?.closest('[role="menu"]') === node,
          };
        });
        assert.deepEqual(layout, { ribbonGrew: false, below: true, topmost: true }, label);
      };
      assert.deepEqual(
        await panel
          .locator('.design > section')
          .evaluateAll((nodes) => nodes.map((node) => node.getAttribute('aria-label'))),
        ['Themes', 'Variants', 'Customize', 'Designer'],
      );
      for (const name of ['Variants', 'Design Suggestions'])
        assert.equal(await panel.getByRole('button', { name, exact: true }).isDisabled(), true);

      // A theme rewrites the deck theme's colors and fonts.
      await changed(() =>
        panel.getByRole('option', { name: 'Office 2013 - 2022 Theme', exact: true }).click(),
      );
      let deck = await pres();
      assert.equal(getPresentationTheme(deck).accent1.toUpperCase(), '#4472C4');
      assert.equal(getPresentationFonts(deck).majorLatin, 'Calibri Light');
      assert.equal(
        await panel
          .getByRole('option', { name: 'Office 2013 - 2022 Theme', exact: true })
          .getAttribute('aria-selected'),
        'true',
      );

      await panel.getByRole('button', { name: 'Colors', exact: true }).click();
      await assertMenuUnclipped('Colors');
      await changed(() => panel.getByRole('menuitemradio', { name: 'Green' }).click());
      assert.equal(getPresentationTheme(await pres()).accent1.toUpperCase(), '#549E39');

      await panel.getByRole('button', { name: 'Fonts', exact: true }).click();
      await assertMenuUnclipped('Fonts');
      await changed(() => panel.getByRole('menuitemradio', { name: /^Georgia/ }).click());
      assert.equal(getPresentationFonts(await pres()).minorLatin, 'Georgia');

      await panel.getByRole('button', { name: 'Slide Size', exact: true }).click();
      await assertMenuUnclipped('Slide Size');
      await changed(() => panel.getByRole('menuitem', { name: 'Standard (4:3)' }).click());
      deck = await pres();
      assert.equal(getSlideSize(deck).width, SLIDE_SIZE_4_3.width);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
