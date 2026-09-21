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
            node.setSelectionRange(range.start, range.end);
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
      await input.fill('English\n日本語\nThird paragraph\nNew paragraph');
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
      await input.fill('First\nSecond');
      await input.evaluate((node) => {
        node.setSelectionRange(8, 8);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      const bar = editor.getByRole('group', { name: 'Selected text formatting', exact: true });
      await bar.getByLabel('Paragraph alignment', { exact: true }).selectOption('right');
      await bar.getByLabel('List style', { exact: true }).selectOption('number');
      await saved();
      await bar.getByLabel('List level', { exact: true }).selectOption({ value: '3' });
      await saved();
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
      await input.focus();
      await input.evaluate((node) => {
        node.setSelectionRange(6, 12);
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
            node.setSelectionRange(range.start, range.end);
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
        await input.fill('Prefix\nEnglish\n日本語\nThird paragraph');
        await input.evaluate((node) => {
          node.setSelectionRange(7, 14);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
        const toggle = async (key, label) => {
          if (control === 'keyboard') await input.press(key);
          else
            await editor
              .locator('.text-format-bar')
              .getByRole('button', { name: label, exact: true })
              .click();
        };
        assert.equal(
          await editor
            .locator('.text-format-bar')
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
        const bar = editor.locator('.text-format-bar');
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
          node.setSelectionRange(7, 14);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
        await toggle('Control+u', '下線');
        await saved();
        assert.equal(
          await bar.getByRole('button', { name: '下線', exact: true }).getAttribute('aria-pressed'),
          'false',
        );
        if (control === 'toolbar') {
          await bar.getByRole('button', { name: '取り消し線', exact: true }).click();
          await saved();
          await bar.getByRole('button', { name: '下付き', exact: true }).click();
          await saved();
        }
        await page.reload();
        await saved();
        text = await shape();
        if (control === 'toolbar') {
          assert.equal(getShapeParagraphElements(text, 1)[0].format.strike, false);
          assert.equal(getShapeParagraphElements(text, 1)[0].format.baseline, 0);
          assert.equal(getShapeParagraphElements(text, 2)[0].format.baseline, undefined);
        }
        assert.equal(getShapeText(text), 'Prefix\nEnglish\n日本語\nThird paragraph');
        assert.equal(getShapeParagraphElements(text, 1)[0].format.underline, false);
        assert.equal(getShapeParagraphElements(text, 1)[0].format.bold, false);
        assert.equal(getShapeParagraphElements(text, 2)[0].format.italic, true);
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
