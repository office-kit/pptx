import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile, copyFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  loadPresentation,
  getSlides,
  getSlideShapes,
  getShapeParagraphElements,
  getShapeText,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';

for (const locale of ['en', 'ja'])
  test(
    `outline clipboard retains pending rich text and supports cut undo (${locale})`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-outline-clipboard-'));
      let preview, browser;
      try {
        await copyFile(
          new URL('../../../../test/fixtures/minimal/one-text-slide.pptx', import.meta.url),
          join(dir, 'template.pptx'),
        );
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          `import {readFileSync} from 'node:fs'; import {Presentation,Slide,Fill} from '@office-kit/pptx-dsl'; const source = new Uint8Array(readFileSync(new URL('./template.pptx', import.meta.url))); export default <Presentation source={source} mode="compose"><Slide layout={{name:"Title and Content"}}><Fill target={{placeholder:{type:'title'}}} format={{bold:true}}>Heading</Fill><Fill target={{placeholder:{idx:1}}} format={{italic:true}}>Body</Fill></Slide></Presentation>;`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
        const errors = [];
        page.on('pageerror', (error) => errors.push(error.message));
        await page.addInitScript(
          (language) => localStorage.setItem('ok-editor-locale', language),
          locale,
        );
        await page.goto(preview.url);
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        const view = locale === 'en' ? 'View' : '表示';
        const outlineName = locale === 'en' ? 'Outline View' : 'アウトライン表示';
        await editor.getByRole('tab', { name: view, exact: true }).click();
        await editor
          .getByRole('tabpanel', { name: view, exact: true })
          .getByRole('button', { name: outlineName, exact: true })
          .click();
        const outline = editor.getByRole('navigation', { name: outlineName, exact: true });
        const title = outline.getByRole('textbox').nth(0);
        const body = outline.getByRole('textbox').nth(1);
        const change = async (action) => {
          const before = (await waitForState(preview.url, () => true)).revision;
          await action();
          await waitForState(preview.url, (state) => state.revision !== before);
        };
        const shape = async (index) =>
          getSlideShapes(
            getSlides(
              await loadPresentation(
                new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
              ),
            )[0],
          )[index];
        const select = async (node, start, end) => {
          await node.focus();
          await node.evaluate((input, range) => input.setSelectionRange(...range), [start, end]);
        };
        const paste = async (node, contents) =>
          node.evaluate((input, data) => {
            const clipboardData = new DataTransfer();
            for (const [type, text] of Object.entries(data)) clipboardData.setData(type, text);
            input.dispatchEvent(
              new ClipboardEvent('paste', { clipboardData, bubbles: true, cancelable: true }),
            );
          }, contents);
        // Copy before autosave: pending input must be present without creating a history entry.
        await title.focus();
        const copied = await title.evaluate((input) => {
          input.setSelectionRange(input.value.length, input.value.length);
          input.dispatchEvent(
            new InputEvent('beforeinput', {
              bubbles: true,
              inputType: 'insertText',
              data: '日本語',
            }),
          );
          input.value += '日本語';
          input.setSelectionRange(input.value.length, input.value.length);
          input.dispatchEvent(
            new InputEvent('input', { bubbles: true, inputType: 'insertText', data: '日本語' }),
          );
          input.setSelectionRange(0, input.value.length);
          const clipboardData = new DataTransfer();
          input.dispatchEvent(
            new ClipboardEvent('copy', { clipboardData, bubbles: true, cancelable: true }),
          );
          return Object.fromEntries(
            [...clipboardData.types].map((type) => [type, clipboardData.getData(type)]),
          );
        });
        assert.equal(copied['text/plain'], 'Heading日本語');
        assert.ok(copied['text/html'].includes('font-weight'));
        await change(() => title.press('Escape'));
        await select(body, 0, 4);
        await change(async () => {
          await paste(body, copied);
          await body.press('Escape');
        });
        assert.equal(getShapeText(await shape(1)), 'Heading日本語');
        assert.ok(getShapeParagraphElements(await shape(1), 0).every((run) => run.format?.bold));
        await select(body, 7, 10);
        await change(async () => {
          const cut = await body.evaluate((input) => {
            const clipboardData = new DataTransfer();
            input.dispatchEvent(
              new ClipboardEvent('cut', { clipboardData, bubbles: true, cancelable: true }),
            );
            return clipboardData.getData('text/plain');
          });
          assert.equal(cut, '日本語');
          await body.press('Escape');
        });
        assert.equal(getShapeText(await shape(1)), 'Heading');
        await change(() => body.press('Control+z'));
        assert.equal(getShapeText(await shape(1)), 'Heading日本語');
        await change(() => body.press('Control+z'));
        assert.equal(getShapeText(await shape(1)), 'Body');
        await select(body, 0, 4);
        await change(async () => {
          await paste(body, {
            'text/plain': '赤い文字',
            'text/html': '<b style="color:#ff0000">赤い</b><i>文字</i>',
          });
          await body.press('Escape');
        });
        const runs = getShapeParagraphElements(await shape(1), 0);
        assert.equal(runs[0].format.bold, true);
        assert.equal(runs[0].format.color?.toLowerCase(), '#ff0000');
        assert.equal(runs.at(-1).format.italic, true);
        const menuAction = async (node, name) => {
          await node.dispatchEvent('contextmenu', { clientX: 140, clientY: 260 });
          await editor.getByRole('menuitem', { name: new RegExp('^' + name) }).click();
        };
        const copyName = locale === 'en' ? 'Copy' : 'コピー';
        const cutName = locale === 'en' ? 'Cut' : '切り取り';
        const pasteName = locale === 'en' ? 'Paste' : '貼り付け';
        await select(body, 0, 4);
        await menuAction(body, copyName);
        await page.waitForFunction(
          async () => (await navigator.clipboard.readText()) === '赤い文字',
        );
        await change(() => menuAction(body, cutName));
        assert.equal(getShapeText(await shape(1)), '');
        await change(() => menuAction(body, pasteName));
        assert.equal(getShapeText(await shape(1)), '赤い文字');
        assert.equal(
          getShapeParagraphElements(await shape(1), 0)[0].format.color?.toLowerCase(),
          '#ff0000',
        );
        await change(() => body.press('Control+z'));
        assert.equal(getShapeText(await shape(1)), '');
        await change(() => body.press('Control+z'));
        assert.equal(getShapeText(await shape(1)), '赤い文字');
        // A delayed permission response must not paste into an editor the user has left.
        await page.evaluate(() => navigator.clipboard.writeText('must not be inserted'));
        await select(body, 0, 4);
        await body.evaluate((input) => {
          const win = input.ownerDocument.defaultView;
          const clipboard = win.navigator.clipboard;
          const read = clipboard.read.bind(clipboard);
          clipboard.read = () =>
            new Promise((resolve) => {
              win.finishOutlinePaste = async () => {
                clipboard.read = read;
                resolve(await read());
              };
            });
        });
        await menuAction(body, pasteName);
        await title.focus();
        await title.evaluate(async (input) => {
          const win = input.ownerDocument.defaultView;
          await win.finishOutlinePaste();
          delete win.finishOutlinePaste;
        });
        await page.waitForTimeout(700);
        assert.equal(await body.inputValue(), '赤い文字');
        assert.equal(getShapeText(await shape(1)), '赤い文字');
        await page.reload();
        await editor.locator('.slide-workspace').waitFor();
        assert.equal(getShapeText(await shape(1)), '赤い文字');
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
