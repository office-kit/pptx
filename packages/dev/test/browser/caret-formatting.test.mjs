import { installRichTextSelection } from '../helpers/rich-text.mjs';
import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getShapeParagraphElements,
  getTableCells,
  getTableCellParagraphs,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

for (const kind of ['shape', 'cell'])
  test(
    `caret typing formats preserve surrounding text in ${kind}`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-caret-format-'));
      let preview, browser;
      try {
        const file = join(dir, 'deck.tsx');
        const element =
          kind === 'shape'
            ? '<Text x={1} y={1} width={7} height={3}>aaa</Text>'
            : '<Table x={1} y={1} width={6} height={3} rows={[["aaa","Other"]]} />';
        await writeFile(
          file,
          `import {Presentation,Slide,Text,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide>${element}</Slide></Presentation>`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        const errors = [];
        page.on('pageerror', (e) => errors.push(e.message));
        await installRichTextSelection(page);
        await page.goto(preview.url);
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        let ja = false;
        const saved = () =>
          editor
            .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
            .waitFor();
        const runs = async () => {
          const pres = await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          );
          const shape = getSlideShapes(getSlides(pres)[0])[0];
          return kind === 'shape'
            ? getShapeParagraphElements(shape, 0)
            : getTableCellParagraphs(getTableCells(shape)[0][0])[0].elements;
        };
        await saved();
        await editor
          .locator('.hit')
          .first()
          .dblclick({ position: { x: 30, y: 20 } });
        const input = editor.locator('.inline-edit');
        const bar = editor.getByRole('group', { name: 'Selected text formatting', exact: true });
        const select = async (start, end = start) => {
          await input.focus();
          await input.evaluate(
            (node, range) => {
              window.selectEditorText(node, ...range);
              node.dispatchEvent(new Event('select', { bubbles: true }));
            },
            [start, end],
          );
        };
        await select(1);
        await input.press('Control+b');
        assert.equal(
          await bar.getByRole('button', { name: 'Bold', exact: true }).getAttribute('aria-pressed'),
          'true',
        );
        assert.equal((await runs()).map((r) => r.text).join(''), 'aaa');
        await input.press('a');
        await bar.getByLabel('Font size', { exact: true }).fill('32');
        await bar.getByLabel('Font size', { exact: true }).press('Tab');
        await bar.getByRole('button', { name: 'Text color', exact: true }).click();
        await editor
          .getByRole('menu', { name: 'Text color', exact: true })
          .getByRole('menuitemradio', { name: 'Accent 2', exact: true })
          .click();
        await bar.getByRole('button', { name: 'Text color', exact: true }).click();
        const palette = editor.getByRole('menu', { name: 'Text color', exact: true });
        const accent = palette.getByRole('menuitemradio', { name: 'Accent 2', exact: true });
        assert.equal(await accent.getAttribute('aria-checked'), 'true');
        assert.equal(
          await bar
            .getByRole('button', { name: 'Text color', exact: true })
            .locator('.swatch')
            .evaluate((node) => getComputedStyle(node).backgroundColor),
          await accent.evaluate((node) => getComputedStyle(node).backgroundColor),
        );
        await palette.press('Escape');
        await input.press('b');
        await input.press('Control+b');
        await input.press('c');
        await select(0);
        assert.equal(
          await bar.getByRole('button', { name: 'Bold', exact: true }).getAttribute('aria-pressed'),
          'false',
        );
        await input.press('z');
        await bar.getByRole('button', { name: 'Done', exact: true }).click();
        await saved();
        const formats = (await runs()).flatMap((r) =>
          Array.from(r.text, (text) => ({ text, ...r.format })),
        );
        assert.equal(formats.map((r) => r.text).join(''), 'zaabcaa');
        assert.notEqual(formats[0].bold, true);
        assert.notEqual(formats[1].bold, true);
        assert.equal(formats[2].bold, true);
        assert.equal(formats[3].bold, true);
        assert.equal(formats[3].size, 32);
        assert.equal(formats[3].color, 'accent2');
        assert.equal(formats[4].bold, false);
        assert.equal(formats[4].size, 32);
        assert.equal(formats[4].color, 'accent2');
        assert.notEqual(formats[5].color, 'accent2');
        assert.notEqual(formats[5].bold, true);
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await saved();
        assert.equal((await runs()).map((r) => r.text).join(''), 'aaa');
        await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
        await saved();
        await editor.locator('.lang select').selectOption('ja');
        ja = true;
        await editor
          .locator('.hit')
          .first()
          .dblclick({ position: { x: 30, y: 20 } });
        await select(4);
        const jaBar = editor.getByRole('group', { name: '選択した文字の書式', exact: true });
        await jaBar.getByRole('button', { name: '文字の書式を解除', exact: true }).click();
        await input.focus();
        await page.keyboard.insertText('日本語');
        await jaBar.getByRole('button', { name: '完了', exact: true }).click();
        await saved();
        const reset = (await runs()).find((r) => r.text.includes('日本語'));
        assert.equal(reset.format?.bold, undefined);
        assert.equal(reset.format?.size, undefined);
        await page.reload();
        await saved();
        assert.equal((await runs()).map((r) => r.text).join(''), 'zaab日本語caa');
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
