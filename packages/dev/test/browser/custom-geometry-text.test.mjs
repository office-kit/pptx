// A custom-geometry shape states its own text rectangle in `<a:custGeom>
// <a:rect>`, and the editor's caret has to sit where the renderer puts the
// glyphs — two independent implementations of the same rule (the editor lays
// out HTML padding, the renderer emits a `<foreignObject>`).
//
// There is no public API to author custGeom, so the geometry is injected into
// the saved deck, the same way the inherited-geometry test edits slide XML.

import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { chromium } from 'playwright';
import {
  createPresentation,
  addBlankSlide,
  addSlideShape,
  addSlideTextBox,
  groupShapes,
  setShapeBounds,
  setShapeTextMargins,
  savePresentation,
  inches,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// A right triangle, with the text pinned to the box's top-right quadrant
// through built-in guides. Neither the quadrant nor the triangle matches the
// preview's preset table, so a text rect taken from anywhere but the file
// itself lands somewhere else.
const CUST_GEOM =
  '<a:custGeom><a:avLst/><a:gdLst/><a:rect l="hc" t="t" r="r" b="vc"/>' +
  '<a:pathLst><a:path w="100" h="100">' +
  '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
  '<a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
  '<a:lnTo><a:pt x="50" y="100"/></a:lnTo>' +
  '<a:close/></a:path></a:pathLst></a:custGeom>';

for (const grouped of [false, true]) {
  test(
    `inline text matches the custGeom text rect (grouped=${grouped})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-custgeom-text-'));
      let preview, browser;
      try {
        const pres = createPresentation(),
          slide = addBlankSlide(pres);
        const shape = addSlideShape(slide, {
          preset: 'rect',
          x: inches(2),
          y: inches(1),
          w: inches(5),
          h: inches(3),
          text: '日本語 English',
        });
        setShapeTextMargins(shape, {
          left: inches(0.15),
          right: inches(0.25),
          top: inches(0.1),
          bottom: inches(0.2),
        });
        if (grouped) {
          const sibling = addSlideTextBox(slide, {
            x: inches(8),
            y: inches(1),
            w: inches(1),
            h: inches(1),
            text: 'Sibling',
          });
          const group = groupShapes([shape, sibling]);
          // Stretches 7×3 in to 8×4 in: the group scale reaches the text
          // bounds but not the `<a:ext>` the rect is written against.
          setShapeBounds(group, { x: inches(1), y: inches(0.5), w: inches(8), h: inches(4) });
        }
        const parts = unzipSync(await savePresentation(pres));
        const path = 'ppt/slides/slide1.xml';
        // The shape above is the first in document order, so this is its own
        // geometry even once the group wraps it.
        parts[path] = strToU8(
          strFromU8(parts[path]).replace(
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
            CUST_GEOM,
          ),
        );
        const source = join(dir, 'source.pptx'),
          file = join(dir, 'deck.tsx');
        await writeFile(source, zipSync(parts));
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
        if (grouped) await editor.locator('.hit').dblclick();
        await editor.locator('.hit').first().dblclick();
        const input = editor.locator('.inline-edit');
        await input.waitFor();
        const layout = await input.evaluate((node) => {
          const box = node.getBoundingClientRect(),
            style = getComputedStyle(node);
          const region = document.querySelector('.paint foreignObject');
          const rendered = region.getBoundingClientRect();
          const shapeBox = document.querySelector('.paint path').getBoundingClientRect();
          return {
            actual: [
              box.x + parseFloat(style.paddingLeft),
              box.y + parseFloat(style.paddingTop),
              box.width - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight),
              box.height - parseFloat(style.paddingTop) - parseFloat(style.paddingBottom),
            ],
            expected: [rendered.x, rendered.y, rendered.width, rendered.height],
            shape: [shapeBox.x, shapeBox.y, shapeBox.width, shapeBox.height],
          };
        });
        layout.actual.forEach((value, i) =>
          assert.ok(Math.abs(value - layout.expected[i]) < 1, JSON.stringify(layout)),
        );
        // Guard against both sides agreeing on the whole box: the stated rect
        // is the top-right quarter, so the text starts past the shape's
        // horizontal middle and ends above its vertical middle.
        const [shapeX, shapeY, shapeW, shapeH] = layout.shape;
        assert.ok(layout.actual[0] > shapeX + shapeW / 2, JSON.stringify(layout));
        assert.ok(
          layout.actual[1] + layout.actual[3] < shapeY + shapeH / 2,
          JSON.stringify(layout),
        );

        await input.fill('編集済み Edited');
        await page.keyboard.press('ControlOrMeta+Enter');
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        await editor.locator('.lang select').selectOption('ja');
        await page.reload();
        await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
        if (grouped) await editor.locator('.hit').dblclick();
        await editor.locator('.hit').first().dblclick();
        assert.equal(await input.innerText(), '編集済み Edited');
        // The saved deck still carries the custom geometry, so the rect is
        // still the one being honoured after a round trip.
        const reopened = await input.evaluate((node) => {
          const box = node.getBoundingClientRect(),
            style = getComputedStyle(node);
          const rendered = document.querySelector('.paint foreignObject').getBoundingClientRect();
          return {
            actual: [box.x + parseFloat(style.paddingLeft), box.y + parseFloat(style.paddingTop)],
            expected: [rendered.x, rendered.y],
          };
        });
        reopened.actual.forEach((value, i) =>
          assert.ok(Math.abs(value - reopened.expected[i]) < 1, JSON.stringify(reopened)),
        );
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}
