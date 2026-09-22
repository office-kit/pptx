import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  getTableCells,
  setShapeTextMargins,
  setShapeTextAnchor,
  setTableCellMargins,
  setTableCellAnchor,
  savePresentation,
  inches,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

for (const table of [false, true])
  for (const anchor of ['top', 'center', 'bottom']) {
    test(
      `inline body margins and ${anchor} anchor (table=${table})`,
      { timeout: 60000 },
      async () => {
        const dir = await mkdtemp(join(tmpdir(), 'office-text-body-'));
        let preview, browser;
        try {
          const pres = createPresentation(),
            slide = addBlankSlide(pres);
          const bounds = { x: inches(2), y: inches(1), w: inches(5), h: inches(3) };
          const margins = {
            left: inches(0.3),
            right: inches(0.2),
            top: inches(0.4),
            bottom: inches(0.1),
          };
          const shape = table
            ? addSlideTable(slide, { ...bounds, rows: [['日本語 English']] })
            : addSlideTextBox(slide, { ...bounds, text: '日本語 English' });
          if (table) {
            const cell = getTableCells(shape)[0][0];
            setTableCellMargins(cell, margins);
            setTableCellAnchor(cell, anchor);
          } else {
            setShapeTextMargins(shape, margins);
            setShapeTextAnchor(shape, anchor);
          }
          const source = join(dir, 'source.pptx'),
            file = join(dir, 'deck.tsx');
          await writeFile(source, await savePresentation(pres));
          await writeFile(
            file,
            `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
          );
          preview = await startPreview(file);
          browser = await chromium.launch({ headless: true });
          const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
          await page.goto(preview.url);
          await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
          const editor = page.frameLocator('#editor-frame');
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
          await editor.locator('.hit').dblclick();
          const input = editor.locator('.inline-edit');
          await input.waitFor();
          const layout = await input.evaluate((node) => {
            const css = getComputedStyle(node),
              box = node.getBoundingClientRect();
            const paragraph = node.querySelector('[data-text-paragraph]').getBoundingClientRect();
            const zoom = Number(css.getPropertyValue('--text-zoom'));
            return {
              zoom,
              padding: [css.paddingLeft, css.paddingRight, css.paddingTop, css.paddingBottom].map(
                parseFloat,
              ),
              top: paragraph.top - box.top,
              height: paragraph.height,
              boxHeight: box.height,
            };
          });
          [0.3, 0.2, 0.4, 0.1].forEach((inch, i) =>
            assert.ok(
              Math.abs(layout.padding[i] - inch * 96 * layout.zoom) < 0.1,
              JSON.stringify(layout),
            ),
          );
          const free = layout.boxHeight - layout.padding[2] - layout.padding[3] - layout.height;
          const expected = layout.padding[2] + free * { top: 0, center: 0.5, bottom: 1 }[anchor];
          assert.ok(Math.abs(layout.top - expected) < 1, JSON.stringify({ layout, expected }));
          await page.screenshot({
            path: `/tmp/pptx-body-${table ? 'table' : 'shape'}-${anchor}.png`,
          });
          await input.fill('編集済み Edited');
          await page.keyboard.press('ControlOrMeta+Enter');
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
          await editor.locator('.lang select').selectOption('ja');
          await page.reload();
          await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
          await editor.locator('.hit').dblclick();
          assert.equal(await input.innerText(), '編集済み Edited');
          const saved = await input.evaluate((node) => ({
            padding: parseFloat(getComputedStyle(node).paddingLeft),
            zoom: Number(getComputedStyle(node).getPropertyValue('--text-zoom')),
            anchor: getComputedStyle(node).alignContent,
          }));
          assert.equal(
            saved.anchor,
            { top: 'start', center: 'safe center', bottom: 'safe end' }[anchor],
          );
          assert.ok(Math.abs(saved.padding - 0.3 * 96 * saved.zoom) < 0.1);
        } finally {
          await browser?.close();
          await preview?.close();
          await rm(dir, { recursive: true, force: true });
        }
      },
    );
  }
