import { installRichTextSelection } from '../helpers/rich-text.mjs';
import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getTableCells,
  getTableCellParagraphs,
  getSlides,
  getSlideShapes,
  getParagraphBullet,
  getParagraphPropertiesEffective,
  getParagraphAlignment,
  getShapeParagraphElements,
  getShapeParagraphCount,
  getShapeText,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// Exercise an incremental edit, not textarea.fill's select-all replacement.
async function fillPreservingText(input, value) {
  const before = await input.textContent();
  let start = 0;
  while (start < before.length && start < value.length && before[start] === value[start]) start++;
  let end = before.length;
  let newEnd = value.length;
  while (end > start && newEnd > start && before[end - 1] === value[newEnd - 1]) {
    end--;
    newEnd--;
  }
  await input.focus();
  await input.evaluate(
    (node, range) => {
      window.selectEditorText(node, ...range);
      node.dispatchEvent(new Event('select', { bubbles: true }));
    },
    [start, end],
  );
  if (start !== end && !value.slice(start, newEnd)) await input.press('Backspace');
  if (value.slice(start, newEnd))
    await input.page().keyboard.insertText(value.slice(start, newEnd));
}

test(
  'inline paragraph formatting follows caret and selection with pending text edits',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-paragraph-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={4} paragraphs={[{runs:[{text:'English',format:{bold:true}}]},{runs:[{text:'日本語',format:{italic:true}}]},{runs:[{text:'Third paragraph'}]}]} /></Slide><Slide><Text x={1} y={1} width={7} height={3} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let locale = 'en';
      const saved = () =>
        editor
          .getByText(locale === 'en' ? 'Saved to this project' : 'このプロジェクトに保存済み', {
            exact: true,
          })
          .waitFor();
      const shape = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        )[0];
      const levels = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const text = getSlideShapes(getSlides(pres)[0])[0];
        return Array.from(
          { length: getShapeParagraphCount(text) },
          (_, i) => getParagraphPropertiesEffective(pres, text, i).level,
        );
      };
      await saved();
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.inline-edit');
      const select = async (start, end = start) => {
        await input.focus();
        await input.evaluate(
          (node, range) => {
            window.selectEditorText(node, range.start, range.end);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          },
          { start, end },
        );
      };
      const bar = editor.locator('.text-format-bar');
      await select(9);
      await bar.getByLabel('Paragraph alignment', { exact: true }).selectOption('center');
      await saved();
      assert.equal(getParagraphAlignment(await shape(), 1), 'ctr');
      assert.deepEqual(
        await input
          .locator('[data-text-paragraph]')
          .evaluateAll((nodes) => nodes.map((n) => getComputedStyle(n).textAlign)),
        ['left', 'center', 'left'],
      );
      assert.notEqual(getParagraphAlignment(await shape(), 0), 'ctr');
      assert.notEqual(getParagraphAlignment(await shape(), 2), 'ctr');
      await bar.getByLabel('List level', { exact: true }).selectOption({ value: '8' });
      await saved();
      assert.deepEqual((await levels()).slice(0, 3), [0, 8, 0]);
      await select(0, 8);
      await bar.getByLabel('List style', { exact: true }).selectOption('bullet');
      await saved();
      assert.equal(getParagraphBullet(await shape(), 0), 'bullet');
      assert.notEqual(getParagraphBullet(await shape(), 1), 'bullet');
      await fillPreservingText(input, 'Prefix\nEnglish\n日本語\nThird paragraph');
      await select(16);
      assert.equal(
        await bar.getByLabel('Paragraph alignment', { exact: true }).inputValue(),
        'center',
      );
      assert.equal(await bar.getByLabel('List level', { exact: true }).inputValue(), '8');
      assert.equal(getShapeText(await shape()), 'English\n日本語\nThird paragraph');
      await fillPreservingText(input, 'English\n日本語\nThird paragraph');
      await fillPreservingText(input, 'English\n日本語\nThird paragraph\nNew paragraph');
      await select(30);
      await bar.getByLabel('Paragraph alignment', { exact: true }).selectOption('right');
      await saved();
      assert.equal(getParagraphAlignment(await shape(), 3), 'r');
      assert.notEqual(getParagraphAlignment(await shape(), 2), 'r');
      assert.equal(getShapeParagraphElements(await shape(), 0)[0].format.bold, true);
      assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.italic, true);
      await bar.getByRole('button', { name: 'Done', exact: true }).click();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(getShapeText(await shape()), 'English\n日本語\nThird paragraph');
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor.locator('.lang select').selectOption('ja');
      locale = 'ja';
      await editor.locator('.hit').first().dblclick();
      await fillPreservingText(input, '\n日本語\nThird paragraph\nNew paragraph');
      await select(2);
      assert.equal(await bar.getByLabel('段落の配置', { exact: true }).inputValue(), 'center');
      assert.equal(await bar.getByLabel('リストの階層', { exact: true }).inputValue(), '8');
      await fillPreservingText(input, 'English\n日本語\nThird paragraph\nNew paragraph');
      await select(8, 28);
      assert.equal(await bar.getByLabel('リストの階層', { exact: true }).inputValue(), '');
      await bar.getByLabel('リストの階層', { exact: true }).selectOption({ value: '2' });
      await saved();
      assert.deepEqual(await levels(), [0, 2, 2, 0]);
      await bar.getByLabel('段落の配置', { exact: true }).selectOption('justify');
      await saved();
      assert.equal(getParagraphAlignment(await shape(), 1), 'just');
      assert.equal(getParagraphAlignment(await shape(), 2), 'just');
      assert.equal(getParagraphAlignment(await shape(), 3), 'r');
      await bar.getByRole('button', { name: '完了', exact: true }).click();
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await levels(), [0, 8, 0, 0]);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.equal(getParagraphAlignment(await shape(), 1), 'just');
      assert.equal(getParagraphBullet(await shape(), 0), 'bullet');
      assert.deepEqual(await levels(), [0, 2, 2, 0]);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'inline table paragraph formatting applies only within the edited cell',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-paragraph-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Table x={1} y={1} width={6} height={3} rows={[["First","Other"]]} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let locale = 'en';
      const saved = () =>
        editor
          .getByText(locale === 'en' ? 'Saved to this project' : 'このプロジェクトに保存済み', {
            exact: true,
          })
          .waitFor();
      const shape = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        )[0];
      await saved();
      const hit = editor.locator('.hit').first();
      const bounds = await hit.boundingBox();
      await hit.dblclick({ position: { x: bounds.width / 4, y: bounds.height / 2 } });
      const input = editor.locator('.inline-edit');
      await fillPreservingText(input, 'First\nSecond');
      await input.evaluate((node) => {
        window.selectEditorText(node, 8, 8);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      const bar = editor.getByRole('group', { name: 'Selected text formatting', exact: true });
      await bar.getByLabel('Paragraph alignment', { exact: true }).selectOption('right');
      await bar.getByLabel('List style', { exact: true }).selectOption('number');
      await saved();
      await bar.getByLabel('List level', { exact: true }).selectOption({ value: '3' });
      await saved();
      await input.focus();
      await input.press('Control+BracketLeft');
      await saved();
      assert.equal(await bar.getByLabel('List level', { exact: true }).inputValue(), '2');
      assert.equal(await input.textContent(), 'First\nSecond');
      await input.press('Control+BracketRight');
      await saved();
      await input.press('Control+z');
      await bar
        .getByLabel('List level', { exact: true })
        .locator('option:checked[value="2"]')
        .waitFor({ state: 'attached' });
      assert.equal(await bar.getByLabel('List level', { exact: true }).inputValue(), '2');
      assert.equal(await input.textContent(), 'First\nSecond');
      await input.press('Control+y');
      await bar
        .getByLabel('List level', { exact: true })
        .locator('option:checked[value="3"]')
        .waitFor({ state: 'attached' });
      await bar.getByLabel('Line spacing mode', { exact: true }).selectOption('pct');
      await bar.getByLabel('Line spacing value', { exact: true }).fill('2');
      await bar.getByLabel('Line spacing value', { exact: true }).press('Tab');
      await saved();
      let cells = getTableCells(await shape());
      assert.equal(getParagraphAlignment(cells[0][0], 1), 'r');
      assert.equal(getParagraphBullet(cells[0][0], 1), 'number');
      const presForLevels = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      assert.equal(getParagraphPropertiesEffective(presForLevels, cells[0][0], 1).level, 3);
      assert.deepEqual(getParagraphPropertiesEffective(presForLevels, cells[0][0], 1).lineSpacing, {
        kind: 'pct',
        value: 2,
      });
      assert.notDeepEqual(
        getParagraphPropertiesEffective(presForLevels, cells[0][0], 0).lineSpacing,
        { kind: 'pct', value: 2 },
      );
      assert.notDeepEqual(
        getParagraphPropertiesEffective(presForLevels, cells[0][1], 0).lineSpacing,
        { kind: 'pct', value: 2 },
      );
      assert.equal(getParagraphPropertiesEffective(presForLevels, cells[0][0], 0).level, 0);
      assert.equal(getParagraphPropertiesEffective(presForLevels, cells[0][1], 0).level, 0);
      assert.notEqual(getParagraphAlignment(cells[0][0], 0), 'r');
      assert.notEqual(getParagraphAlignment(cells[0][1], 0), 'r');
      await fillPreservingText(input, 'Prefix\nFirst\nSecond');
      await input.evaluate((node) => {
        window.selectEditorText(node, 15, 15);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      assert.equal(
        await bar.getByLabel('Paragraph alignment', { exact: true }).inputValue(),
        'right',
      );
      assert.equal(await bar.getByLabel('List style', { exact: true }).inputValue(), 'number');
      assert.equal(await bar.getByLabel('List level', { exact: true }).inputValue(), '3');
      assert.equal(await bar.getByLabel('Line spacing value', { exact: true }).inputValue(), '2');
      await fillPreservingText(input, 'First\nSecond');
      await input.focus();
      await input.evaluate((node) => {
        window.selectEditorText(node, 6, 12);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      await bar.getByRole('button', { name: 'Strikethrough', exact: true }).click();
      await saved();
      await bar.getByRole('button', { name: 'Superscript', exact: true }).click();
      await saved();
      await bar.getByRole('button', { name: 'Done', exact: true }).click();
      await page.reload();
      await saved();
      cells = getTableCells(await shape());
      assert.equal(getParagraphAlignment(cells[0][0], 1), 'r');
      assert.equal(getParagraphBullet(cells[0][0], 1), 'number');
      assert.equal(getTableCellParagraphs(cells[0][0])[1].elements[0].format.strike, true);
      assert.equal(getTableCellParagraphs(cells[0][0])[1].elements[0].format.baseline, 0.3);
      assert.notEqual(getTableCellParagraphs(cells[0][0])[0].elements[0].format?.strike, true);
      assert.notEqual(getTableCellParagraphs(cells[0][1])[0].elements[0].format?.baseline, 0.3);
      await hit.dblclick({ position: { x: bounds.width / 4, y: bounds.height / 2 } });
      await input.focus();
      await input.evaluate((node) => {
        window.selectEditorText(node, 6, 9);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      await bar.getByRole('button', { name: 'Clear text formatting', exact: true }).click();
      await saved();
      cells = getTableCells(await shape());
      assert.deepEqual(getTableCellParagraphs(cells[0][0])[1].elements[0].format ?? {}, {});
      assert.equal(getTableCellParagraphs(cells[0][0])[1].elements[1].format.strike, true);
      await bar.getByRole('button', { name: 'Done', exact: true }).click();
      await editor.getByRole('button', { name: 'Cell 1, 1', exact: true }).click();
      await editor
        .getByRole('group', { name: 'Format selected cells', exact: true })
        .getByRole('button', { name: 'Clear text formatting', exact: true })
        .click();
      await saved();
      await page.reload();
      await saved();
      cells = getTableCells(await shape());
      for (const element of getTableCellParagraphs(cells[0][0])[1].elements)
        assert.deepEqual(element.format ?? {}, {});
      assert.equal(getParagraphAlignment(cells[0][0], 1), 'r');
      assert.equal(getParagraphBullet(cells[0][0], 1), 'number');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'inline paragraph spacing supports mixed selections, inheritance and history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-paragraph-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={4} paragraphs={[{runs:[{text:'English',format:{bold:true}}]},{runs:[{text:'日本語',format:{italic:true}}]},{runs:[{text:'Third paragraph'}]}]} /></Slide><Slide><Text x={1} y={1} width={7} height={3} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let locale = 'en';
      const saved = () =>
        editor
          .getByText(locale === 'en' ? 'Saved to this project' : 'このプロジェクトに保存済み', {
            exact: true,
          })
          .waitFor();
      const shape = async () =>
        getSlideShapes(
          getSlides(
            await loadPresentation(
              new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
            ),
          )[0],
        )[0];
      const properties = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const text = getSlideShapes(getSlides(pres)[0])[0];
        return Array.from({ length: getShapeParagraphCount(text) }, (_, i) =>
          getParagraphPropertiesEffective(pres, text, i),
        );
      };
      await saved();
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.inline-edit');
      const select = async (start, end = start) => {
        await input.focus();
        await input.evaluate(
          (node, range) => {
            window.selectEditorText(node, range.start, range.end);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          },
          { start, end },
        );
      };
      const bar = editor.locator('.text-format-bar');
      const change = async (name, value) => {
        const control = bar.getByLabel(name, { exact: true });
        await control.fill(value);
        await control.press('Tab');
        await saved();
      };
      await select(9);
      await bar.getByLabel('Line spacing mode', { exact: true }).selectOption('pct');
      await change('Line spacing value', '1.5');
      await change('Before paragraph (pt)', '6');
      await change('After paragraph (pt)', '12');
      let props = await properties();
      assert.deepEqual(props[1].lineSpacing, { kind: 'pct', value: 1.5 });
      assert.equal(props[1].spcBefPts, 6);
      assert.equal(props[1].spcAftPts, 12);
      const spacing = await input
        .locator('[data-text-paragraph]')
        .nth(1)
        .evaluate((n) => {
          const css = getComputedStyle(n);
          return {
            line: parseFloat(css.lineHeight) / parseFloat(css.fontSize),
            before: parseFloat(css.marginTop),
            after: parseFloat(css.marginBottom),
            zoom: Number(css.getPropertyValue('--text-zoom')),
          };
        });
      assert.ok(Math.abs(spacing.line - 1.5) < 0.01);
      assert.ok(Math.abs(spacing.before - 8 * spacing.zoom) < 0.01);
      assert.ok(Math.abs(spacing.after - 16 * spacing.zoom) < 0.01);
      assert.notDeepEqual(props[0].lineSpacing, props[1].lineSpacing);
      assert.notEqual(props[2].spcAftPts, 12);
      await select(0, 12);
      assert.equal(await bar.getByLabel('Line spacing mode', { exact: true }).inputValue(), '');
      assert.equal(await bar.getByLabel('Before paragraph (pt)', { exact: true }).inputValue(), '');
      await bar.getByLabel('Line spacing mode', { exact: true }).selectOption('pts');
      await change('Line spacing value', '24');
      props = await properties();
      assert.deepEqual(props[0].lineSpacing, { kind: 'pts', value: 24 });
      assert.deepEqual(props[1].lineSpacing, props[0].lineSpacing);
      assert.notDeepEqual(props[2].lineSpacing, props[0].lineSpacing);
      await bar.getByRole('button', { name: 'Done', exact: true }).click();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual((await properties())[0].lineSpacing, { kind: 'pts', value: 18 });
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor.locator('.lang select').selectOption('ja');
      locale = 'ja';
      await editor.locator('.hit').first().dblclick();
      await select(9);
      await bar.getByLabel('行間の指定方法', { exact: true }).selectOption('inherit');
      await change('段落前（pt）', '');
      await change('段落後（pt）', '0');
      await page.reload();
      await saved();
      props = await properties();
      assert.deepEqual(props[0].lineSpacing, { kind: 'pts', value: 24 });
      assert.notDeepEqual(props[1].lineSpacing, props[0].lineSpacing);
      assert.notEqual(props[1].spcBefPts, 6);
      assert.equal(props[1].spcAftPts, 0);
      assert.notEqual(props[2].spcAftPts, 12);
      assert.equal(getShapeParagraphElements(await shape(), 0)[0].format.bold, true);
      assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.italic, true);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

for (const control of ['keyboard', 'toolbar'])
  test(
    `inline ${control} formatting toggles current runs after pending edits`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-inline-paragraph-'));
      let preview, browser, page;
      try {
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={4} paragraphs={[{runs:[{text:'English',format:{bold:true}}]},{runs:[{text:'日本語',format:{italic:true}}]},{runs:[{text:'Third paragraph'}]}]} /></Slide><Slide><Text x={1} y={1} width={7} height={3} /></Slide></Presentation>`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        const errors = [];
        page.on('pageerror', (e) => errors.push(e.message));
        await installRichTextSelection(page);
        await page.goto(preview.url);
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        let locale = 'en';
        const saved = () =>
          editor
            .getByText(locale === 'en' ? 'Saved to this project' : 'このプロジェクトに保存済み', {
              exact: true,
            })
            .waitFor();
        const shape = async () =>
          getSlideShapes(
            getSlides(
              await loadPresentation(
                new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
              ),
            )[0],
          )[0];
        await saved();
        await editor.locator('.hit').first().dblclick();
        const input = editor.locator('.inline-edit');
        await fillPreservingText(input, 'Prefix\nEnglish\n日本語\nThird paragraph');
        await input.evaluate((node) => {
          window.selectEditorText(node, 7, 14);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
        if (control === 'toolbar') {
          assert.equal(
            await editor
              .locator('.canvas-shell > .text-format-bar')
              .getByRole('button', { name: 'Bold', exact: true })
              .count(),
            0,
          );
          await editor.getByRole('tab', { name: 'View', exact: true }).click();
          assert.equal(await input.isVisible(), true);
          assert.equal(await editor.locator('.ribbon .font-ribbon').count(), 0);
          await editor.getByRole('tab', { name: 'Home', exact: true }).click();
          await editor.locator('.ribbon .font-ribbon').waitFor();
          await page.screenshot({ path: '/tmp/pptx-home-font-ribbon.png' });
        }
        const toggle = async (key, label) => {
          if (control === 'keyboard') await input.press(key);
          else
            await editor
              .locator('.ribbon .text-format-bar, .canvas-shell > .text-format-bar')
              .getByRole('button', { name: label, exact: true })
              .click();
        };
        assert.equal(
          await editor
            .locator('.ribbon .text-format-bar, .canvas-shell > .text-format-bar')
            .getByRole('button', { name: 'Bold', exact: true })
            .getAttribute('aria-pressed'),
          'true',
        );
        assert.equal(getShapeText(await shape()), 'English\n日本語\nThird paragraph');
        await toggle('Control+b', 'Bold');
        await saved();
        let text = await shape();
        assert.equal(getShapeParagraphElements(text, 1)[0].format.bold, false);
        assert.equal(getShapeParagraphElements(text, 2)[0].format.italic, true);
        const bar = editor.locator('.ribbon .text-format-bar, .canvas-shell > .text-format-bar');
        assert.equal(
          await bar.getByRole('button', { name: 'Bold', exact: true }).getAttribute('aria-pressed'),
          'false',
        );
        await toggle('Meta+i', 'Italic');
        await saved();
        assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.italic, true);
        await toggle('Control+i', 'Italic');
        await saved();
        assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.italic, false);
        if (control === 'toolbar') {
          await bar.getByLabel('Highlight color options', { exact: true }).click();
          await bar.getByRole('button', { name: 'Apply highlight', exact: true }).click();
          await bar.getByLabel('Highlight color options', { exact: true }).click();
          await saved();
          assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.highlight, '#FFFF00');
          await bar.getByLabel('Highlight color', { exact: true }).evaluate((node) => {
            node.value = '#ffcc00';
            node.dispatchEvent(new Event('change', { bubbles: true }));
          });
          await saved();
          assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.highlight, '#FFCC00');
          await bar.getByRole('button', { name: 'Strikethrough', exact: true }).click();
          await saved();
          assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.strike, true);
          await bar.getByRole('button', { name: 'Superscript', exact: true }).click();
          await saved();
          assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.baseline, 0.3);
          await bar.getByRole('button', { name: 'Subscript', exact: true }).click();
          await saved();
          assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.baseline, -0.25);
          assert.equal(
            await bar
              .getByRole('button', { name: 'Superscript', exact: true })
              .getAttribute('aria-pressed'),
            'false',
          );
        }
        await toggle('Meta+u', 'Underline');
        await saved();
        assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.underline, true);
        await bar.getByRole('button', { name: 'Done', exact: true }).click();
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await saved();
        assert.notEqual(getShapeParagraphElements(await shape(), 1)[0].format.underline, true);
        await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
        await saved();
        await editor.locator('.lang select').selectOption('ja');
        locale = 'ja';
        await editor.locator('.hit').first().dblclick();
        await input.focus();
        await input.evaluate((node) => {
          window.selectEditorText(node, 7, 14);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
        await toggle('Control+u', '下線');
        await saved();
        assert.equal(
          await bar.getByRole('button', { name: '下線', exact: true }).getAttribute('aria-pressed'),
          'false',
        );
        if (control === 'toolbar') {
          await bar.getByLabel('蛍光ペンの色のオプション', { exact: true }).click();
          await bar.getByRole('button', { name: 'ハイライトを解除', exact: true }).click();
          await bar.getByLabel('蛍光ペンの色のオプション', { exact: true }).click();
          await saved();
          await bar.getByRole('button', { name: '取り消し線', exact: true }).click();
          await saved();
          await bar.getByRole('button', { name: '下付き', exact: true }).click();
          await saved();
        }
        await page.reload();
        await saved();
        text = await shape();
        if (control === 'toolbar') {
          assert.equal(getShapeParagraphElements(text, 1)[0].format.highlight, undefined);
          assert.equal(getShapeParagraphElements(text, 1)[0].format.strike, false);
          assert.equal(getShapeParagraphElements(text, 1)[0].format.baseline, 0);
          assert.equal(getShapeParagraphElements(text, 2)[0].format.baseline, undefined);
        }
        assert.equal(getShapeText(text), 'Prefix\nEnglish\n日本語\nThird paragraph');
        assert.equal(getShapeParagraphElements(text, 1)[0].format.underline, false);
        assert.equal(getShapeParagraphElements(text, 1)[0].format.bold, false);
        assert.equal(getShapeParagraphElements(text, 2)[0].format.italic, true);
        await editor.locator('.hit').first().dblclick();
        await input.focus();
        await input.evaluate((node) => {
          window.selectEditorText(node, 7, 14);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
        await toggle('Control+Backslash', '文字の書式を解除');
        await saved();
        assert.deepEqual(getShapeParagraphElements(await shape(), 1)[0].format ?? {}, {});
        await bar.getByRole('button', { name: '完了', exact: true }).click();
        await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
        await saved();
        assert.equal(getShapeParagraphElements(await shape(), 1)[0].format.underline, false);
        await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
        await saved();
        await editor.locator('.lang select').selectOption('en');
        locale = 'en';
        await editor.locator('.hit').first().dblclick();
        await input.focus();
        await fillPreservingText(input, 'Prefix\nEnglish\n日本語です\nThird paragraph');
        await input.evaluate((node) => {
          window.selectEditorText(node, 15, 20);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
        await toggle('Meta+Backslash', 'Clear text formatting');
        await saved();
        await page.reload();
        await saved();
        text = await shape();
        assert.equal(getShapeText(text), 'Prefix\nEnglish\n日本語です\nThird paragraph');
        assert.deepEqual(getShapeParagraphElements(text, 2)[0].format ?? {}, {});
        assert.equal(getShapeParagraphElements(text, 0)[0].format.bold, true);
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );

test(
  'inline list markers preserve text offsets and nested numbering in both languages',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-list-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={4} size={24}>{'English\\n日本語\\nThird'}</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.inline-edit');
      const bar = editor.locator('.ribbon .text-format-bar, .canvas-shell > .text-format-bar');
      const select = async (start, end = start) => {
        await input.focus();
        await input.evaluate(
          (node, range) => {
            window.selectEditorText(node, range.start, range.end);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          },
          { start, end },
        );
      };
      const labels = () =>
        input
          .locator('[data-text-paragraph]')
          .evaluateAll((nodes) => nodes.map((n) => n.getAttribute('data-list-marker')));
      const expectLabels = async (expected) => {
        for (const [index, label] of expected.entries()) {
          await input
            .locator(`[data-text-paragraph]:nth-of-type(${index + 1})[data-list-marker="${label}"]`)
            .waitFor();
        }
        assert.deepEqual(await labels(), expected);
      };
      await select(0, 17);
      await bar.getByLabel('List style', { exact: true }).selectOption('number');
      await expectLabels(['1.', '2.', '3.']);
      await select(8);
      await input.press('Tab');
      await expectLabels(['1.', '1.', '2.']);
      assert.equal(await input.evaluate((n) => n === document.activeElement), true);
      await bar.getByRole('button', { name: 'Done', exact: true }).click();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().dblclick();
      await expectLabels(['1.', '2.', '3.']);
      await bar.getByRole('button', { name: 'Done', exact: true }).click();
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().dblclick();
      await expectLabels(['1.', '1.', '2.']);
      await select(8, 17);
      await input.press('Control+BracketRight');
      await expectLabels(['1.', '1.', '1.']);
      await input.press('Shift+Tab');
      await expectLabels(['1.', '1.', '2.']);
      await select(8);
      await bar.getByLabel('List level', { exact: true }).selectOption({ value: '8' });
      await input.focus();
      await input.press('Tab');
      assert.equal(await bar.getByLabel('List level', { exact: true }).inputValue(), '8');
      await bar.getByLabel('List level', { exact: true }).selectOption({ value: '1' });
      await input.focus();
      await input.evaluate((n) =>
        n.dispatchEvent(
          new KeyboardEvent('keydown', {
            key: 'Tab',
            isComposing: true,
            bubbles: true,
            cancelable: true,
          }),
        ),
      );
      assert.equal(await bar.getByLabel('List level', { exact: true }).inputValue(), '1');
      assert.equal(await input.textContent(), 'English\n日本語\nThird');
      const marker = await input
        .locator('[data-text-paragraph]')
        .first()
        .evaluate((n) => {
          const css = getComputedStyle(n, '::before');
          return {
            content: css.content,
            size: css.fontSize,
            textSize: getComputedStyle(n.querySelector('span')).fontSize,
          };
        });
      assert.equal(marker.content, '"1."');
      assert.equal(marker.size, marker.textSize);
      await select(8, 11);
      const copied = await input.evaluate((n) => {
        const clipboardData = new DataTransfer();
        n.dispatchEvent(
          new ClipboardEvent('copy', { clipboardData, bubbles: true, cancelable: true }),
        );
        return clipboardData.getData('text/plain');
      });
      assert.equal(copied, '日本語');
      await input.press('ArrowRight');
      await input.press('Enter');
      await input.pressSequentially('Nested');
      await expectLabels(['1.', '1.', '2.', '2.']);
      assert.equal(await input.textContent(), 'English\n日本語\nNested\nThird');
      await bar.getByRole('button', { name: 'Done', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.lang select').selectOption('ja');
      await editor.locator('.hit').first().dblclick();
      await expectLabels(['1.', '1.', '2.', '2.']);
      await select(0, 7);
      await bar.getByLabel('リストの種類', { exact: true }).selectOption('bullet');
      await expectLabels(['•', '1.', '2.', '1.']);
      await input.focus();
      await input.press('Shift+Tab');
      assert.equal(await bar.getByLabel('リストの階層', { exact: true }).inputValue(), '0');
      assert.equal(await input.evaluate((n) => n === document.activeElement), true);
      await select(8);
      await input.press('Meta+BracketLeft');
      await expectLabels(['•', '1.', '1.', '2.']);
      await page.screenshot({ path: join(tmpdir(), 'pptx-pr287-inline-lists.png') });
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'inline undo crosses formatting boundaries without losing pending redo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inline-history-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={4}>English 日本語</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.inline-edit');
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
      const bold = async (expected) => {
        await input
          .locator('span[style*="font-weight: bold"]')
          .first()
          .waitFor({ state: expected ? 'visible' : 'hidden', timeout: 5000 });
      };
      await select(0, 11);
      await input.press('Control+b');
      await bold(true);
      await input.press('Control+z');
      await bold(false);
      assert.equal(await input.evaluate((node) => document.activeElement === node), true);
      await input.press('Control+Shift+z');
      await bold(true);
      await select(11);
      await page.keyboard.insertText('です');
      assert.equal(await input.textContent(), 'English 日本語です');
      await input.press('Control+z');
      assert.equal(await input.textContent(), 'English 日本語');
      await input.press('Control+z');
      await bold(false);
      await input.press('Control+y');
      await bold(true);
      await input.press('Control+y');
      assert.equal(await input.textContent(), 'English 日本語です');
      // A new edit after undo branches history and cannot replay old text offsets.
      await input.press('Control+z');
      await input.press('Control+z');
      await bold(false);
      await select(11);
      await page.keyboard.insertText('!');
      await input.press('Control+y');
      assert.equal(await input.textContent(), 'English 日本語!');
      await select(0, 12);
      await input.press('Control+b');
      await bold(true);
      await input.press('Meta+z');
      await bold(false);
      assert.equal(await input.textContent(), 'English 日本語');
      await input.press('Meta+Shift+z');
      await bold(true);
      assert.equal(await input.textContent(), 'English 日本語!');
      // Native history events and rapid keyboard requests share the same ordering.
      await input.evaluate((node) => {
        node.dispatchEvent(
          new InputEvent('beforeinput', {
            inputType: 'historyUndo',
            bubbles: true,
            cancelable: true,
          }),
        );
        node.dispatchEvent(
          new KeyboardEvent('keydown', {
            key: 'z',
            ctrlKey: true,
            shiftKey: true,
            bubbles: true,
            cancelable: true,
          }),
        );
        node.dispatchEvent(
          new KeyboardEvent('keydown', {
            key: 'z',
            ctrlKey: true,
            bubbles: true,
            cancelable: true,
          }),
        );
      });
      await bold(false);
      await editor.locator('.inline-edit[aria-busy="false"]').waitFor();
      await bold(false);
      assert.equal(await input.textContent(), 'English 日本語');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
