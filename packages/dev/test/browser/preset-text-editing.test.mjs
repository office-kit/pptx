import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
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

for (const { preset, grouped = false, collapsed = false } of [
  ...['rect', 'triangle', 'diamond', 'pentagon', 'star5', 'leftRightArrow'].map((preset) => ({
    preset,
  })),
  { preset: 'triangle', grouped: true },
  { preset: 'star5', collapsed: true },
]) {
  test(
    `inline text matches ${preset} preview text region (grouped=${grouped}, collapsed=${collapsed})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-preset-text-'));
      let preview, browser;
      try {
        const pres = createPresentation(),
          slide = addBlankSlide(pres);
        const shape = addSlideShape(slide, {
          preset,
          x: inches(2),
          y: inches(1),
          w: inches(5),
          h: inches(3),
          text: '日本語 English',
        });
        setShapeTextMargins(shape, {
          left: inches(0.15),
          right: inches(0.25),
          top: inches(collapsed ? 3 : 0.1),
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
          setShapeBounds(group, { x: inches(1), y: inches(0.5), w: inches(8), h: inches(4) });
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
        if (grouped) await editor.locator('.hit').dblclick();
        await editor.locator('.hit').first().dblclick();
        const input = editor.locator('.inline-edit');
        await input.waitFor();
        const layout = await input.evaluate((node) => {
          const box = node.getBoundingClientRect(),
            style = getComputedStyle(node);
          const paragraph = node.querySelector('[data-text-paragraph]');
          const region = document.querySelector('.paint foreignObject');
          const rendered = region.getBoundingClientRect();
          return {
            actual: [
              box.x + parseFloat(style.paddingLeft),
              box.y + parseFloat(style.paddingTop),
              box.width - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight),
              box.height - parseFloat(style.paddingTop) - parseFloat(style.paddingBottom),
            ],
            expected: [rendered.x, rendered.y, rendered.width, rendered.height],
            align: getComputedStyle(paragraph).textAlign,
            renderedAlign: getComputedStyle(region.querySelector('p')).textAlign,
          };
        });
        layout.actual.forEach((value, i) =>
          assert.ok(Math.abs(value - layout.expected[i]) < 1, JSON.stringify(layout)),
        );
        // The editor's default must track the renderer's, not just a literal.
        assert.equal(layout.align, layout.renderedAlign);
        assert.equal(layout.align, 'center');
        assert.equal(
          await editor
            .locator('.canvas-shell > .text-format-bar')
            .getByLabel('Paragraph alignment', { exact: true })
            .inputValue(),
          'center',
        );
        assert.equal(
          await editor
            .locator('.paragraphs')
            .getByLabel('Paragraph alignment', { exact: true })
            .inputValue(),
          'center',
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
        assert.equal(
          await editor
            .locator('.canvas-shell > .text-format-bar')
            .getByLabel('段落の配置', { exact: true })
            .inputValue(),
          'center',
        );
        assert.equal(
          await editor
            .locator('.paragraphs')
            .getByLabel('段落の配置', { exact: true })
            .inputValue(),
          'center',
        );
        if (preset === 'rect') {
          const alignment = editor
            .locator('.canvas-shell > .text-format-bar')
            .getByLabel('段落の配置', { exact: true });
          await alignment.selectOption('right');
          await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
          assert.equal(
            await editor
              .locator('.paragraphs')
              .getByLabel('段落の配置', { exact: true })
              .inputValue(),
            'right',
          );
          assert.equal(
            await input
              .locator('[data-text-paragraph]')
              .evaluate((node) => getComputedStyle(node).textAlign),
            'right',
          );
          await input.press('ControlOrMeta+z');
          await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
          assert.equal(await alignment.inputValue(), 'center');
          await input.press('ControlOrMeta+Shift+z');
          await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
          assert.equal(await alignment.inputValue(), 'right');
        }
        await page.screenshot({ path: `/tmp/pptx-preset-${preset}.png` });
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}
