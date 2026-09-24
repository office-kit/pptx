import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  cm,
  addTitleSlide,
  getShapeBodyPrEffective,
  getShapeTextAutoFit,
  getShapeTextColumns,
  getShapeText,
  getSlides,
  getSlideShapes,
  loadPresentation,
  savePresentation,
  removeSlide,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'inherited text layout is displayed and a single-column override survives save and undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inherited-layout-'));
    let preview, browser;
    try {
      const parts = unzipSync(
        await readFile(new URL('../../../../test/fixtures/minimal/blank.pptx', import.meta.url)),
      );
      const layout = 'ppt/slideLayouts/slideLayout1.xml';
      parts[layout] = strToU8(
        strFromU8(parts[layout]).replaceAll(
          '<a:bodyPr/>',
          '<a:bodyPr numCol="3" spcCol="90000"><a:normAutofit fontScale="75000"/></a:bodyPr>',
        ),
      );
      const pres = await loadPresentation(zipSync(parts));
      const originals = [...getSlides(pres)];
      addTitleSlide(pres, 'Inherited layout');
      for (const slide of originals) removeSlide(pres, slide);
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(pres));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () => {
        const copy = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        const title = getSlideShapes(getSlides(copy)[0])[0];
        return { body: getShapeBodyPrEffective(copy, title), literal: getShapeTextColumns(title) };
      };
      await saved();
      await editor.locator('.hit').first().click();
      const box = editor.locator('.text-box');
      await box.locator('summary').click();
      assert.equal(
        await box.getByRole('radio', { name: 'Shrink text on overflow', exact: true }).isChecked(),
        true,
      );
      await box.getByRole('button', { name: 'Columns...', exact: true }).click();
      const dialog = editor.getByRole('dialog', { name: 'Columns', exact: true });
      assert.equal(
        await dialog.getByRole('spinbutton', { name: 'Number of columns:' }).inputValue(),
        '3',
      );
      assert.equal(
        await dialog.getByRole('spinbutton', { name: 'Spacing between columns:' }).inputValue(),
        '0.25',
      );
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      assert.equal((await read()).literal, null, 'unchanged dialog must retain inheritance');
      await box.getByRole('button', { name: 'Columns...', exact: true }).click();
      await dialog.getByRole('spinbutton', { name: 'Number of columns:' }).fill('1');
      await dialog.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      assert.deepEqual((await read()).body.columns, { count: 1, gapEmu: cm(0.25) });
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await read()).literal, null);
      assert.equal((await read()).body.columns.count, 3);
      await editor.locator('.hit').first().dblclick();
      const input = editor.getByRole('textbox', { name: 'Edit text', exact: true });
      assert.equal(await input.evaluate((node) => getComputedStyle(node).columnCount), '3');
      await input.press('Escape');
      await box.getByRole('radio', { name: 'Do not Autofit', exact: true }).check();
      await saved();
      assert.equal((await read()).body.autoFit, 'none');
      assert.equal((await read()).body.autoFitParams, null);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'text-box properties preserve other margins, apply to a mixed selection and undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-text-box-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={2} height={1}>日本語 English</Text><Text x={4} y={1} width={2} height={1}>Second</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const box = editor.locator('.text-box');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const read = async () => {
        const pres = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(pres)[0]).map((shape) => ({
          ...getShapeBodyPrEffective(pres, shape),
          text: getShapeText(shape),
          fit: getShapeTextAutoFit(shape),
          columns: getShapeTextColumns(shape),
        }));
      };
      await saved();
      const initial = await read();
      await editor.locator('.hit').first().click();
      await box.locator('summary').click();
      const verticalAlignment = box.getByRole('combobox', {
        name: 'Vertical alignment',
        exact: true,
      });
      assert.equal(await verticalAlignment.locator('option').count(), 7);
      await verticalAlignment.selectOption('top-centered');
      await saved();
      assert.equal((await read())[0].anchorCentered, true);
      const textBlock = editor
        .locator('.paint foreignObject > div')
        .filter({ hasText: '日本語 English' })
        .first();
      assert.ok(await textBlock.isVisible());
      const previewTranslation = await textBlock.evaluate((el) => parseFloat(el.style.translate));
      assert.ok(previewTranslation > 0);
      await editor.locator('.hit').first().dblclick();
      const textEditor = editor.getByRole('textbox', { name: 'Edit text', exact: true });
      await textEditor.waitFor();
      assert.match(await textEditor.evaluate((el) => el.style.transform), /translate\([1-9]/);
      await textEditor.press('ControlOrMeta+a');
      await textEditor.press('Backspace');
      await textEditor.pressSequentially('Centered');
      await textEditor.press('Enter');
      await textEditor.pressSequentially('A');
      await textEditor.press('ControlOrMeta+Enter');
      await saved();
      assert.equal((await read())[0].text, 'Centered\nA');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      // Restore the fixture in one selection edit so the following mixed tests
      // retain their original expected text, independent of typing-burst history.
      await editor.locator('.hit').first().dblclick();
      await textEditor.press('ControlOrMeta+a');
      await textEditor.fill('日本語 English');
      await textEditor.press('ControlOrMeta+Enter');
      await saved();
      await verticalAlignment.selectOption('top');
      await saved();
      assert.equal((await read())[0].anchorCentered, false);
      const left = () =>
        box.getByRole('spinbutton', { name: ja ? '左余白' : 'Left margin', exact: true });
      await left().fill('0.375');
      await left().press('Tab');
      await saved();
      const changed = await read();
      assert.deepEqual(changed[0].margins, { ...initial[0].margins, left: cm(0.375) });
      assert.deepEqual(changed[1], initial[1]);
      for (const invalid of ['-1', '55.89', '']) {
        const before = await left().inputValue();
        await left().fill(invalid);
        await left().press('Tab');
        assert.equal(await left().inputValue(), before);
        assert.deepEqual(await read(), changed);
      }
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      assert.equal(await left().inputValue(), '');
      await left().fill('0.5');
      await left().press('Tab');
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.margins.left),
        [cm(0.5), cm(0.5)],
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), changed);
      await box
        .getByRole('combobox', { name: 'Text direction', exact: true })
        .selectOption('vert270');
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.vert),
        ['vert270', 'vert270'],
      );
      await box.getByRole('radio', { name: 'Shrink text on overflow', exact: true }).check();
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.fit),
        ['normal', 'normal'],
      );
      await box.getByRole('checkbox', { name: 'Wrap text in shape', exact: true }).uncheck();
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.wrap),
        ['none', 'none'],
      );
      const beforeColumns = await read();
      await box.getByRole('button', { name: 'Columns...', exact: true }).click();
      const columns = editor.getByRole('dialog', { name: 'Columns', exact: true });
      await columns.getByRole('spinbutton', { name: 'Number of columns:' }).fill('3');
      await columns.getByRole('spinbutton', { name: 'Spacing between columns:' }).fill('0.25');
      assert.deepEqual(await read(), beforeColumns);
      await columns.getByRole('button', { name: 'Cancel', exact: true }).click();
      assert.deepEqual(await read(), beforeColumns);
      await box.getByRole('button', { name: 'Columns...', exact: true }).click();
      await columns.getByRole('spinbutton', { name: 'Number of columns:' }).fill('17');
      await columns.getByRole('button', { name: 'OK', exact: true }).click();
      assert.equal(await columns.isVisible(), true);
      await columns.getByRole('spinbutton', { name: 'Number of columns:' }).fill('3');
      await columns.getByRole('spinbutton', { name: 'Spacing between columns:' }).fill('0.25');
      await columns.getByRole('spinbutton', { name: 'Spacing between columns:' }).press('Enter');
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.columns),
        [
          { count: 3, gapEmu: cm(0.25) },
          { count: 3, gapEmu: cm(0.25) },
        ],
      );
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), beforeColumns);
      await box.getByRole('button', { name: 'Columns...', exact: true }).click();
      await columns.getByRole('spinbutton', { name: 'Number of columns:' }).press('Escape');
      await columns.waitFor({ state: 'hidden' });
      assert.deepEqual(await read(), beforeColumns);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await box
        .getByRole('combobox', { name: '垂直方向の配置', exact: true })
        .selectOption('bottom');
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.anchor),
        ['bottom', 'bottom'],
      );
      const top = box.getByRole('spinbutton', { name: '上余白', exact: true });
      await top.fill('0.2');
      await top.press('Tab');
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.margins.top),
        [cm(0.2), cm(0.2)],
      );
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(
        (await read()).map((item) => item.margins.top),
        initial.map((item) => item.margins.top),
      );
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
