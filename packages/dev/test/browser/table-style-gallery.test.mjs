import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getTableCellFill,
  getTableCells,
  getTableStyleFlags,
  getTableStyleId,
  isTableShape,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const DECK = `import { Presentation, Slide, Table } from '@office-kit/pptx-dsl';
export default (
  <Presentation>
    <Slide><Table x={0.5} y={1} width={8} height={2} rows={[["A","B","C"],["D","E","F"],["G","H","I"]]} cellStyle={{ fill: '#FFFF00' }} /></Slide>
  </Presentation>
);
`;

// Table Design ▸ Table Styles offers the reference desktop app's 74 built-in styles: a strip
// showing the current style's row, and › opening the whole gallery under the
// four headings the reference desktop app uses, then Clear Table.
test(
  'Table Design offers every built-in table style, named as the reference desktop app names them',
  { timeout: 180000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-table-styles-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(file, DECK);
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1512, height: 900 } });
      await page.goto(preview.url + '/editor');
      await page.getByText('Saved to this project', { exact: true }).waitFor();
      const panel = page.locator('#ribbon-panel');
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
      const saved = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0]).find(isTableShape);
      };
      await page.locator('.hit').nth(0).click();
      await page.getByRole('tab', { name: 'Table Design', exact: true }).click();

      // The strip shows the row holding the table's style (Medium Style 2).
      const strip = panel.locator('.table-strip button');
      await strip.first().waitFor();
      assert.deepEqual(await strip.evaluateAll((items) => items.map((item) => item.title)), [
        'Medium Style 2',
        'Medium Style 2 - Accent 1',
        'Medium Style 2 - Accent 2',
        'Medium Style 2 - Accent 3',
        'Medium Style 2 - Accent 4',
        'Medium Style 2 - Accent 5',
        'Medium Style 2 - Accent 6',
      ]);
      assert.equal(
        await panel
          .getByRole('button', { name: 'Medium Style 2 - Accent 1', exact: true })
          .getAttribute('aria-pressed'),
        'true',
      );
      // Swatches are the preview renderer's pictures of a styled table.
      const swatch = await strip.nth(1).locator('img').getAttribute('src');
      assert.match(
        decodeURIComponent(swatch),
        /^data:image\/svg\+xml,<svg[\s\S]*data-pptx-cell="0,0"/,
      );

      // The style options redraw the swatches with the new flags.
      const before = await strip.nth(2).locator('img').getAttribute('src');
      await changed(() =>
        panel.getByRole('checkbox', { name: 'First Column', exact: true }).check(),
      );
      assert.equal(getTableStyleFlags(await saved()).firstCol, true);
      assert.ok(
        (await strip.nth(2).locator('img').getAttribute('src')) !== before,
        'the swatches show the first column',
      );

      // › opens the whole gallery under the reference desktop app's headings.
      const more = panel.getByRole('button', { name: 'Next Table Styles gallery', exact: true });
      await more.click();
      const gallery = page.getByRole('menu', { name: 'Table Styles', exact: true });
      assert.deepEqual(await gallery.locator('.heading').allTextContents(), [
        'Best Match for Document',
        'Light',
        'Medium',
        'Dark',
      ]);
      const items = gallery.getByRole('menuitemradio');
      assert.equal(await items.count(), 74);
      assert.equal(
        await gallery
          .getByRole('group', { name: 'Dark', exact: true })
          .getByRole('menuitemradio')
          .count(),
        11,
      );
      assert.equal(await items.locator('img').count(), 74);
      // Themed Style 1's swatch carries the theme gradient with the reference desktop app's
      // over-saturated stop: accent 6 (#F79646) with tint 50% + satMod 300%
      // is #FFBE87 in the reference desktop app's own export (2 levels of sampling slack).
      const themed = decodeURIComponent(
        await gallery
          .getByRole('menuitemradio', { name: 'Themed Style 1 - Accent 6', exact: true })
          .locator('img')
          .getAttribute('src'),
      );
      const firstStop = /<linearGradient[\s\S]*?stop-color="#([0-9A-Fa-f]{6})"/.exec(themed)?.[1];
      assert.ok(firstStop, 'the swatch draws the theme gradient');
      const reference = [0xff, 0xbe, 0x87];
      for (const [i, level] of reference.entries())
        assert.ok(
          Math.abs(Number.parseInt(firstStop.slice(i * 2, i * 2 + 2), 16) - level) <= 2,
          `gradient stop #${firstStop} is the reference desktop app's #FFBE87`,
        );
      // Picking one applies it and writes the reference desktop app's definition.
      await changed(() =>
        gallery
          .getByRole('menuitemradio', { name: 'Dark Style 2 - Accent 3/Accent 4', exact: true })
          .click(),
      );
      assert.equal(await gallery.count(), 0, 'the gallery closes after a pick');
      assert.equal(getTableStyleId(await saved()), '{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}');
      assert.equal(
        await strip.first().getAttribute('title'),
        'Dark Style 2',
        'the strip follows the new style’s row',
      );

      // Clear Table: No Style, No Grid, and no cell formatting of its own.
      await more.click();
      await changed(() =>
        gallery.getByRole('menuitem', { name: 'Clear Table', exact: true }).click(),
      );
      const cleared = await saved();
      assert.equal(getTableStyleId(cleared), '{2D5ABB26-0587-4C30-8999-92F81FD0307C}');
      assert.equal(getTableCellFill(getTableCells(cleared)[1][1]), null);

      // Japanese names follow the reference desktop app's.
      await page.locator('.lang select').selectOption('ja');
      await page.locator('.hit').nth(0).click();
      await page.getByRole('tab', { name: 'テーブル デザイン', exact: true }).click();
      await panel.getByRole('button', { name: '次の表のスタイル ギャラリー', exact: true }).click();
      const ja = page.getByRole('menu', { name: '表のスタイル', exact: true });
      assert.deepEqual(await ja.locator('.heading').allTextContents(), [
        'ドキュメントに最適なスタイル',
        '淡色',
        '中間',
        '濃色',
      ]);
      for (const name of [
        'スタイルなし、表のグリッド線なし',
        'テーマ スタイル 1 - アクセント 1',
        'スタイル (淡色) 1',
        '中間スタイル 2 - アクセント 1',
        '濃色スタイル 2 - アクセント 5/アクセント 6',
      ])
        assert.equal(await ja.getByRole('menuitemradio', { name, exact: true }).count(), 1, name);
      assert.equal(await ja.getByRole('menuitem', { name: '表のクリア', exact: true }).count(), 1);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
