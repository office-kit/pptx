import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

test('custom tab alignment uses painted browser font widths', { timeout: 90000 }, async () => {
  const dir = await mkdtemp(join(tmpdir(), 'office-custom-tabs-'));
  let preview, browser;
  try {
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={5} height={3}>{'A\\t12.34'}</Text></Slide></Presentation>`,
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
    await saved();
    await editor.locator('.hit').first().click();
    for (const alignment of ['Left', 'Center', 'Right', 'Decimal']) {
      await editor
        .locator('.ribbon')
        .getByRole('button', { name: 'Line spacing', exact: true })
        .click();
      await editor.getByRole('menuitem', { name: 'Line Spacing Options...', exact: true }).click();
      const paragraph = editor.getByRole('dialog', { name: 'Paragraph', exact: true });
      await paragraph.getByRole('button', { name: 'Tabs...', exact: true }).click();
      const tabs = editor.getByRole('dialog', { name: 'Tabs', exact: true });
      await tabs.getByLabel('Tab stop position:', { exact: true }).fill('5.08');
      await tabs.getByRole('radio', { name: alignment, exact: true }).check();
      await tabs.getByRole('button', { name: 'Set', exact: true }).click();
      await tabs.getByRole('button', { name: 'OK', exact: true }).click();
      await paragraph.getByRole('button', { name: 'OK', exact: true }).click();
      await saved();
      const placement = await editor
        .locator('.paint [data-pptx-paragraph]')
        .first()
        .evaluate((element) => {
          const spans = [...element.querySelectorAll('tspan')];
          const start = spans.find((span) => span.textContent === 'A');
          const field = spans.find((span) => span.textContent === '12.34');
          if (!start || !field) throw new Error('Missing tab field glyphs');
          return {
            start: field.getStartPositionOfChar(0).x - start.getStartPositionOfChar(0).x,
            width: field.getComputedTextLength(),
            decimal: field.getSubStringLength(0, 2),
          };
        });
      const shift =
        alignment === 'Center'
          ? placement.width / 2
          : alignment === 'Right'
            ? placement.width
            : alignment === 'Decimal'
              ? placement.decimal
              : 0;
      assert.ok(
        Math.abs(placement.start + shift - 192) < 1,
        `${alignment}: ${JSON.stringify(placement)}`,
      );
      const textEditor = editor.getByRole('textbox', { name: 'Edit text', exact: true });
      if (!(await textEditor.isVisible())) await editor.locator('.hit').first().click();
      await textEditor.click();
      const input = editor.locator('[contenteditable="true"]');
      const editing = input.locator('[data-text-paragraph]').first();
      const measureEditing = () =>
        editing.evaluate((element) => {
          const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
          let node;
          while ((node = walker.nextNode())) {
            const index = node.textContent.indexOf('12.34');
            if (index < 0) continue;
            const range = document.createRange();
            range.setStart(node, index);
            range.setEnd(node, index + 5);
            const rect = range.getBoundingClientRect();
            range.setEnd(node, index + 2);
            const zoom = Number(getComputedStyle(element).getPropertyValue('--text-zoom'));
            return {
              start: (rect.left - element.getBoundingClientRect().left) / zoom,
              width: rect.width / zoom,
              decimal: range.getBoundingClientRect().width / zoom,
              text: element.textContent,
            };
          }
          throw new Error('Missing editable tab field');
        });
      const assertEditing = async (expectedText) => {
        const edited = await measureEditing();
        const editShift =
          alignment === 'Center'
            ? edited.width / 2
            : alignment === 'Right'
              ? edited.width
              : alignment === 'Decimal'
                ? edited.decimal
                : 0;
        assert.ok(
          Math.abs(edited.start + editShift - 192) < 1,
          `editing ${alignment}: ${JSON.stringify(edited)}`,
        );
        assert.equal(edited.text, expectedText);
      };
      await assertEditing('A\t12.34');
      await input.press('Home');
      await input.press('ArrowRight');
      await page.keyboard.insertText('B');
      await input.filter({ hasText: /^AB\t12\.34$/ }).waitFor();
      await assertEditing('AB\t12.34');
      await input.press('ControlOrMeta+z');
      await input.filter({ hasText: /^A\t12\.34$/ }).waitFor();
      await assertEditing('A\t12.34');
      await input.press('Escape');
    }
    assert.deepEqual(errors, []);
  } finally {
    await browser?.close();
    await preview?.close();
    await rm(dir, { recursive: true, force: true });
  }
});

test('editing tab widths follow capitalization and kerning', async () => {
  const { build } = await import('esbuild');
  const { fileURLToPath } = await import('node:url');
  const bundle = await build({
    stdin: {
      contents: `export {layoutEditingTabs} from './editing-tabs.ts';`,
      loader: 'ts',
      resolveDir: fileURLToPath(new URL('../../../../site/src/lib/editor/core/', import.meta.url)),
    },
    bundle: true,
    format: 'esm',
    platform: 'browser',
    write: false,
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    const results = await page.evaluate(async (code) => {
      const module = await import(
        URL.createObjectURL(new Blob([code], { type: 'text/javascript' }))
      );
      return [
        { text: 'av av av', transform: 'uppercase', kerning: 'normal', caps: 'normal' },
        { text: 'AVAVAVAV', transform: 'none', kerning: 'none', caps: 'normal' },
        { text: 'Small Capitals', transform: 'none', kerning: 'normal', caps: 'small-caps' },
        { text: 'Spaced', transform: 'none', kerning: 'normal', caps: 'normal', spacing: 4 },
        { text: 'Tight', transform: 'none', kerning: 'normal', caps: 'normal', spacing: -2 },
      ].map((item) => {
        const root = document.createElement('div');
        root.style.cssText = `font:80px Arial;white-space:pre;`;
        const paragraph = document.createElement('section');
        paragraph.dataset.tabStops = '600:right';
        paragraph.style.cssText = `text-transform:${item.transform};font-kerning:${item.kerning};font-variant-caps:${item.caps};letter-spacing:${item.spacing ?? 0}px;tab-size:96px`;
        paragraph.textContent = '\t' + item.text;
        root.append(paragraph);
        document.body.append(root);
        module.layoutEditingTabs(root, 1);
        const range = document.createRange();
        range.selectNodeContents(paragraph.lastChild);
        const right = range.getBoundingClientRect().right - paragraph.getBoundingClientRect().left;
        root.remove();
        return { ...item, right };
      });
    }, bundle.outputFiles[0].text);
    for (const result of results)
      assert.ok(Math.abs(result.right - 600) < 1, JSON.stringify(result));
  } finally {
    await browser.close();
  }
});

test('vertical editing tabs align along the inline axis', async () => {
  const { build } = await import('esbuild');
  const { fileURLToPath } = await import('node:url');
  const bundle = await build({
    stdin: {
      contents: `export {layoutEditingTabs} from './editing-tabs.ts';`,
      loader: 'ts',
      resolveDir: fileURLToPath(new URL('../../../../site/src/lib/editor/core/', import.meta.url)),
    },
    bundle: true,
    format: 'esm',
    platform: 'browser',
    write: false,
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    const results = await page.evaluate(async (code) => {
      const module = await import(
        URL.createObjectURL(new Blob([code], { type: 'text/javascript' }))
      );
      return [
        { mode: 'vertical-rl', orientation: 'mixed', text: 'AV 12.34' },
        { mode: 'vertical-lr', orientation: 'mixed', text: 'AV 12.34' },
        { mode: 'vertical-rl', orientation: 'upright', text: '日本語 12.34' },
      ].flatMap((item) =>
        ['left', 'center', 'right', 'decimal'].map((alignment) => {
          const root = document.createElement('div');
          root.style.cssText = `font:40px Arial;white-space:pre;writing-mode:${item.mode};text-orientation:${item.orientation};height:800px;`;
          const paragraph = document.createElement('section');
          paragraph.dataset.tabStops = `600:${alignment}`;
          paragraph.style.cssText = 'tab-size:96px;';
          paragraph.textContent = 'A\t' + item.text;
          root.append(paragraph);
          document.body.append(root);
          module.layoutEditingTabs(root, 1);
          const range = document.createRange();
          range.selectNodeContents(paragraph.lastChild);
          const bounds = range.getBoundingClientRect();
          const top = paragraph.getBoundingClientRect().top;
          range.setEnd(paragraph.lastChild, item.text.indexOf('.'));
          const decimal = range.getBoundingClientRect().height;
          const right =
            bounds.top -
            top +
            (alignment === 'left'
              ? 0
              : alignment === 'center'
                ? bounds.height / 2
                : alignment === 'decimal'
                  ? decimal
                  : bounds.height);
          root.remove();
          return { ...item, alignment, right };
        }),
      );
    }, bundle.outputFiles[0].text);
    for (const result of results)
      assert.ok(Math.abs(result.right - 600) < 1, JSON.stringify(result));
  } finally {
    await browser.close();
  }
});

test('browser preview measures tracked SVG text at its painted width', async () => {
  const { build } = await import('esbuild');
  const { fileURLToPath } = await import('node:url');
  const bundle = await build({
    entryPoints: [
      fileURLToPath(new URL('../../../preview/src/browser-measure.ts', import.meta.url)),
    ],
    bundle: true,
    format: 'esm',
    platform: 'browser',
    write: false,
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    const results = await page.evaluate(async (code) => {
      const module = await import(
        URL.createObjectURL(new Blob([code], { type: 'text/javascript' }))
      );
      const measure = module.browserTextMeasurer();
      const results = [];
      for (const text of ['Spacing', 'AVAV', 'e\u0301e\u0301']) {
        for (const spacing of [4, -2, 0]) {
          const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
          const run = document.createElementNS(svg.namespaceURI, 'text');
          run.style.cssText = `font:40px Arial;font-kerning:normal;letter-spacing:${spacing}px`;
          run.textContent = text;
          svg.append(run);
          document.body.append(svg);
          const painted = run.getComputedTextLength();
          const measured = measure(text, {
            family: 'Arial',
            sizePx: 40,
            bold: false,
            italic: false,
            letterSpacingPx: spacing,
          }).widthPx;
          results.push({ text, spacing, painted, measured });
          svg.remove();
        }
      }
      return results;
    }, bundle.outputFiles[0].text);
    for (const result of results)
      assert.ok(Math.abs(result.painted - result.measured) < 0.1, JSON.stringify(result));
  } finally {
    await browser.close();
  }
});
