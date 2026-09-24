import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  applySlideBackgroundToAll,
  getSlideBackgroundImageFillLayout,
  getSlideBackgroundImageOpacity,
  getShapeImageBytes,
  getShapeKind,
  getSlideShapes,
  setSlideBackgroundImage,
  getSlides,
  getSlideBackground,
  loadPresentation,
  savePresentation,
} from '../../../../dist/index.js';
import { startPreview } from '../helpers/server.mjs';

test(
  'background picture layout edits inherited fills, remembers modes, undoes and persists',
  { timeout: 120000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-background-picture-'));
    let preview, browser;
    try {
      const pres = await loadPresentation(
        await readFile(
          new URL('../../../../test/fixtures/minimal/one-image-slide.pptx', import.meta.url),
        ),
      );
      const slide = getSlides(pres)[0];
      const picture = getSlideShapes(slide).find((shape) => getShapeKind(shape) === 'picture');
      setSlideBackgroundImage(slide, getShapeImageBytes(picture));
      addBlankSlide(pres);
      applySlideBackgroundToAll(pres, slide);
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(pres));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const read = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlides(deck).map((slide) => ({
          kind: getSlideBackground(slide).kind,
          layout: getSlideBackgroundImageFillLayout(slide),
          opacity: getSlideBackgroundImageOpacity(slide),
        }));
      };
      const open = async () => {
        await editor.getByRole('tab', { name: 'Design', exact: true }).click();
        await editor.getByRole('button', { name: 'Background Styles', exact: true }).click();
        await editor.getByRole('menuitem', { name: 'Format Background...', exact: true }).click();
      };
      await saved();
      await open();
      assert.equal(
        await editor
          .getByRole('radio', { name: 'Picture or texture fill', exact: true })
          .isChecked(),
        true,
      );
      const transparency = editor.getByRole('spinbutton', {
        name: 'Picture transparency',
        exact: true,
      });
      await transparency.fill('65');
      await transparency.press('Tab');
      await saved();
      assert.equal((await read())[0].opacity, 0.35);
      assert.equal((await read())[1].opacity, null);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await read())[0].opacity, null);
      await transparency.fill('20');
      await transparency.press('Tab');
      await saved();
      const left = editor.getByRole('spinbutton', { name: 'Offset left', exact: true });
      await left.fill('25');
      await left.press('Tab');
      await saved();
      let state = await read();
      assert.equal(state[0].layout.left, 0.25);
      assert.equal(state[1].kind, 'inherit');
      assert.equal(state[1].layout.left, 0);
      const tile = editor.getByRole('checkbox', { name: 'Tile picture as texture', exact: true });
      await tile.check();
      await saved();
      const scale = editor.getByRole('spinbutton', { name: 'Scale X', exact: true });
      await scale.fill('50');
      await scale.press('Tab');
      await saved();
      await editor.getByLabel('Alignment', { exact: true }).selectOption('ctr');
      await saved();
      await editor.getByLabel('Mirror type', { exact: true }).selectOption('xy');
      await saved();
      await tile.uncheck();
      await saved();
      assert.equal((await read())[0].layout.left, 0.25);
      await tile.check();
      await saved();
      assert.equal((await read())[0].layout.scaleX, 0.5);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal((await read())[0].layout.mode, 'stretch');
      await tile.check();
      await saved();
      await page.reload();
      await saved();
      await open();
      assert.equal(await scale.inputValue(), '50');
      state = await read();
      assert.equal(state[0].layout.alignment, 'ctr');
      assert.equal(state[0].layout.flip, 'xy');
      assert.equal(state[1].kind, 'inherit');
      assert.equal(
        await editor.getByRole('checkbox', { name: 'Rotate with shape', exact: true }).isEnabled(),
        false,
      );
      const frame = page.frames().find((frame) => frame.url().includes('/editor'));
      assert.ok(frame);
      await frame.evaluate(() => {
        Object.defineProperty(navigator, 'clipboard', {
          configurable: true,
          value: { read: async () => [] },
        });
      });
      const clipboard = editor.getByRole('button', { name: 'Clipboard', exact: true });
      assert.equal(await transparency.inputValue(), '20');
      assert.equal((await read())[0].opacity, 0.8);
      const beforePaste = await read();
      await clipboard.click();
      await editor
        .getByRole('alert')
        .filter({ hasText: 'The clipboard does not contain a picture.' })
        .waitFor();
      assert.deepEqual(await read(), beforePaste);
      await frame.evaluate(() => {
        navigator.clipboard.read = async () => {
          throw new DOMException('Denied', 'NotAllowedError');
        };
      });
      await clipboard.click();
      await editor.getByRole('alert').filter({ hasText: 'Clipboard access was denied' }).waitFor();
      assert.deepEqual(await read(), beforePaste);
      const image = Array.from(getShapeImageBytes(picture));
      await frame.evaluate((bytes) => {
        navigator.clipboard.read = async () => [
          {
            types: ['image/png'],
            getType: async () => new Blob([new Uint8Array(bytes)], { type: 'image/png' }),
          },
        ];
      }, image);
      await clipboard.click();
      await left.waitFor({ state: 'visible' });
      await saved();
      assert.equal((await read())[0].layout.mode, 'stretch');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await read(), beforePaste);
      await frame.evaluate((bytes) => {
        navigator.clipboard.read = () =>
          new Promise((resolve) => {
            window.finishBackgroundClipboard = () =>
              resolve([
                {
                  types: ['image/png'],
                  getType: async () => new Blob([new Uint8Array(bytes)], { type: 'image/png' }),
                },
              ]);
          });
      }, image);
      await clipboard.click();
      await editor.locator('.thumb-row').nth(1).click();
      await frame.evaluate(() => window.finishBackgroundClipboard());
      await editor
        .getByRole('alert')
        .filter({ hasText: 'The slide changed. Choose the background image again.' })
        .waitFor();
      assert.deepEqual(await read(), beforePaste);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
