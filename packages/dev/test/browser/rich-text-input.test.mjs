import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addTitleSlide,
  findSlidePlaceholder,
  getSlides,
  getShapeRunFormatEffective,
  getShapeParagraphElements,
  loadPresentation,
  removeSlide,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

test(
  'rich text input displays formatting and preserves native multiline edits and history',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-rich-input-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={3} size={24}>abc</Text></Slide></Presentation>`,
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
      await editor.getByTitle('Zoom...', { exact: true }).click();
      await editor
        .getByRole('dialog', { name: 'Zoom', exact: true })
        .getByRole('radio', { name: '100%', exact: true })
        .check();
      await editor
        .getByRole('dialog', { name: 'Zoom', exact: true })
        .getByRole('button', { name: 'OK', exact: true })
        .click();
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      const input = editor.locator('.inline-edit');
      const select = async (start, end = start) => {
        await input.focus();
        await input.evaluate(
          (node, range) => window.selectEditorText(node, ...range),
          [start, end],
        );
      };
      await select(1, 2);
      assert.equal(
        await input
          .locator('span')
          .first()
          .evaluate((n) => getComputedStyle(n).fontSize),
        '32px',
      );
      // A keyboard/programmatic zoom leaves the input focused, including its selection.
      await editor.getByTitle('Zoom in (Ctrl+=)', { exact: true }).evaluate((n) => n.click());
      await page.waitForFunction(() => {
        const root = document.querySelector('#editor-frame')?.contentDocument;
        const span = root?.querySelector('.inline-edit span');
        return span && Math.abs(parseFloat(getComputedStyle(span).fontSize) - 35.2) < 0.01;
      });
      assert.equal(await input.evaluate(() => window.getSelection().toString()), 'b');
      const copiedHtml = await input.evaluate((node) => {
        const data = new DataTransfer();
        node.dispatchEvent(new ClipboardEvent('copy', { bubbles: true, clipboardData: data }));
        return data.getData('text/html');
      });
      assert.match(copiedHtml, /font-size: 24pt/);
      assert.doesNotMatch(copiedHtml, /--text-zoom/);
      await editor.getByTitle('Zoom out (Ctrl+-)', { exact: true }).evaluate((n) => n.click());
      await input.press('Control+b');
      assert.equal(
        await input
          .locator('span')
          .filter({ hasText: /^b$/ })
          .evaluate((n) => getComputedStyle(n).fontWeight),
        '700',
      );
      await select(2);
      await input.press('Enter');
      await page.keyboard.insertText('日本語');
      assert.equal(await input.textContent(), 'ab\n日本語c');
      await input.press('Control+z');
      assert.equal(await input.textContent(), 'ab\nc');
      await input.press('Control+z');
      assert.equal(await input.textContent(), 'abc');
      await input.press('Control+y');
      await input.press('Control+y');
      assert.equal(await input.textContent(), 'ab\n日本語c');
      await select(0, 6);
      await input.press('Backspace');
      assert.equal(await input.textContent(), 'c');
      await input.press('Control+z');
      assert.equal(await input.textContent(), 'ab\n日本語c');
      await select(7);
      await input.press('Enter');
      await input.press('Enter');
      await page.keyboard.insertText('終わり');
      assert.equal(await input.textContent(), 'ab\n日本語c\n\n終わり');
      await input.press('Control+Enter');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      assert.equal(await input.textContent(), 'ab\n日本語c\n\n終わり');
      await page.screenshot({ path: '/tmp/pptx-pr287-rich-text-input.png', fullPage: true });
      await select(0, 12);
      await input.evaluate((node) => {
        node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
        node.dispatchEvent(
          new InputEvent('beforeinput', {
            bubbles: true,
            inputType: 'insertCompositionText',
            isComposing: true,
          }),
        );
        node.textContent = '候';
        window.selectEditorText(node, 1);
        node.dispatchEvent(
          new InputEvent('input', {
            bubbles: true,
            inputType: 'insertCompositionText',
            isComposing: true,
          }),
        );
      });
      assert.equal(
        await input.locator('span').count(),
        0,
        'composition must keep the browser-owned DOM',
      );
      await input.evaluate((node) => {
        node.dispatchEvent(
          new InputEvent('beforeinput', {
            bubbles: true,
            inputType: 'insertCompositionText',
            isComposing: true,
          }),
        );
        node.textContent = '候補';
        window.selectEditorText(node, 2);
        node.dispatchEvent(
          new InputEvent('input', {
            bubbles: true,
            inputType: 'insertCompositionText',
            isComposing: true,
          }),
        );
        node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '候補' }));
      });
      assert.equal(await input.textContent(), '候補');
      await input.press('Control+z');
      assert.equal(await input.textContent(), 'ab\n日本語c\n\n終わり');
      await select(12);
      await page.keyboard.insertText('確定');
      await input.press('Escape');
      await input.waitFor({ state: 'detached' });
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      assert.equal(await input.textContent(), 'ab\n日本語c\n\n終わり確定');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'inline editing preserves inherited title styles across pending edits and save',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-inherited-input-'));
    let preview, browser;
    try {
      const pres = await loadPresentation(
        await readFile(new URL('../../../../test/fixtures/minimal/blank.pptx', import.meta.url)),
      );
      const originals = [...getSlides(pres)];
      const slide = addTitleSlide(pres, 'Theme title');
      for (const original of originals) removeSlide(pres, original);
      const title = findSlidePlaceholder(slide, 'title') ?? findSlidePlaceholder(slide, 'ctrTitle');
      const expected = getShapeRunFormatEffective(pres, title, 0, 0);
      assert.ok(expected.size > 28);
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
      page.on('pageerror', (e) => errors.push(e.message));
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByTitle('Zoom...', { exact: true }).click();
      await editor
        .getByRole('dialog', { name: 'Zoom', exact: true })
        .getByRole('radio', { name: '100%', exact: true })
        .check();
      await editor
        .getByRole('dialog', { name: 'Zoom', exact: true })
        .getByRole('button', { name: 'OK', exact: true })
        .click();
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      const input = editor.locator('.inline-edit');
      const checkStyle = async () => {
        const bar = editor.getByRole('group', { name: 'Selected text formatting', exact: true });
        assert.equal(
          await bar.getByLabel('Font size', { exact: true }).inputValue(),
          String(expected.size),
        );
        const actual = await input
          .locator('span')
          .last()
          .evaluate((n) => ({
            size: parseFloat(getComputedStyle(n).fontSize),
            font: getComputedStyle(n).fontFamily,
          }));
        assert.ok(Math.abs(actual.size - (expected.size * 4) / 3) < 0.01);
        assert.ok(actual.font.includes(expected.font));
        assert.ok(actual.font.includes('sans-serif'));
        const alignment = await input
          .locator('[data-text-paragraph]')
          .last()
          .evaluate((n) => {
            const box = n.getBoundingClientRect();
            const text = n.querySelector('span').getBoundingClientRect();
            return {
              align: getComputedStyle(n).textAlign,
              centerOffset: Math.abs((text.left + text.right - box.left - box.right) / 2),
            };
          });
        assert.equal(alignment.align, 'center');
        assert.ok(alignment.centerOffset < 1);
      };
      await checkStyle();
      await input.press('Enter');
      await page.keyboard.insertText('日本語');
      assert.equal(await input.textContent(), 'Theme title\n日本語');
      await checkStyle();
      await input.press('Control+Enter');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const saved = await loadPresentation(
        new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
      );
      const savedTitle =
        findSlidePlaceholder(getSlides(saved)[0], 'title') ??
        findSlidePlaceholder(getSlides(saved)[0], 'ctrTitle');
      assert.equal(getShapeParagraphElements(savedTitle, 1)[0].format?.size, undefined);
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      await checkStyle();
      await page.screenshot({ path: '/tmp/pptx-pr287-inherited-text-input.png', fullPage: true });
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
