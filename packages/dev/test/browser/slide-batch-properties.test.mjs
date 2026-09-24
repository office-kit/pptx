import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideBackground,
  getSlideBackgroundImageBytes,
  getSlideLayout,
  getSlideLayoutName,
  getSlideTransition,
  isSlideHidden,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'selected slide properties apply in one transaction with mixed values and bilingual persistence',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-slide-batch-'));
    let preview, browser, page;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['A','B','C'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const thumbs = editor.locator('.thumb-row');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const slides = async () =>
        getSlides(
          await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          ),
        );
      const undo = async () => {
        await editor
          .getByTitle(ja ? '元に戻す (Ctrl+Z)' : 'Undo (Ctrl+Z)', { exact: true })
          .click();
        await saved();
      };
      await saved();
      await thumbs.nth(0).click();
      const pane = editor.getByRole('region', { name: 'Slide options', exact: true });
      await pane.getByLabel('Background color: More Colors...', { exact: true }).fill('#aabbcc');
      await saved();
      await pane.getByLabel('Slide layout', { exact: true }).selectOption({ label: 'Title Slide' });
      await saved();
      await pane.getByLabel('Skip during presentation', { exact: true }).check();
      await saved();
      await thumbs.nth(1).click({ modifiers: ['Shift'] });
      assert.equal(
        await pane
          .getByLabel('Skip during presentation', { exact: true })
          .evaluate((node) => node.indeterminate),
        true,
      );
      assert.equal(
        await pane.getByLabel('Slide layout', { exact: true }).inputValue(),
        '__mixed__',
      );
      await pane.getByText('Background color: Mixed', { exact: true }).waitFor();
      await pane.getByLabel('Skip during presentation', { exact: true }).check();
      await saved();
      assert.deepEqual((await slides()).map(isSlideHidden), [true, true, false]);
      await undo();
      assert.deepEqual((await slides()).map(isSlideHidden), [true, false, false]);
      const color = pane.getByRole('button', { name: 'Background color', exact: true });
      await color.click();
      const palette = editor.getByRole('menu', { name: 'Background color', exact: true });
      assert.equal(await palette.locator('[aria-checked="true"]').count(), 0);
      await palette.getByRole('menuitemradio', { name: 'Accent 3', exact: true }).click();
      await saved();
      assert.deepEqual((await slides()).map(getSlideBackground), [
        { kind: 'solid', color: 'scheme:accent3' },
        { kind: 'solid', color: 'scheme:accent3' },
        { kind: 'inherit' },
      ]);
      await color.click();
      assert.equal(
        await palette
          .getByRole('menuitemradio', { name: 'Accent 3', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await palette.press('Escape');
      await undo();
      assert.deepEqual((await slides()).map(getSlideBackground), [
        { kind: 'solid', color: '#AABBCC' },
        { kind: 'inherit' },
        { kind: 'inherit' },
      ]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual((await slides()).map(getSlideBackground), [
        { kind: 'solid', color: 'scheme:accent3' },
        { kind: 'solid', color: 'scheme:accent3' },
        { kind: 'inherit' },
      ]);
      await undo();
      await pane
        .getByLabel('Slide layout', { exact: true })
        .selectOption({ label: 'Title and Content' });
      await saved();
      assert.deepEqual(
        (await slides()).slice(0, 2).map((slide) => getSlideLayoutName(getSlideLayout(slide))),
        ['Title and Content', 'Title and Content'],
      );
      await undo();
      assert.equal(getSlideLayoutName(getSlideLayout((await slides())[0])), 'Title Slide');
      assert.notEqual(getSlideLayoutName(getSlideLayout((await slides())[1])), 'Title Slide');
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      const jp = editor.getByRole('region', { name: 'スライドの設定', exact: true });
      const bytes = Buffer.from(
        await page.evaluate(() => {
          const canvas = document.createElement('canvas');
          canvas.width = 20;
          canvas.height = 20;
          return canvas.toDataURL('image/png').split(',')[1];
        }),
        'base64',
      );
      await jp
        .getByLabel('背景画像', { exact: true })
        .setInputFiles({ name: 'background.png', mimeType: 'image/png', buffer: bytes });
      await saved();
      for (const slide of (await slides()).slice(0, 2))
        assert.deepEqual(Buffer.from(getSlideBackgroundImageBytes(slide)), bytes);
      assert.equal(getSlideBackground((await slides())[2]).kind, 'inherit');
      await jp.getByRole('button', { name: '背景をリセット', exact: true }).click();
      await saved();
      assert.deepEqual(
        (await slides()).map((slide) => getSlideBackground(slide).kind),
        ['inherit', 'inherit', 'inherit'],
      );
      await undo();
      for (const slide of (await slides()).slice(0, 2))
        assert.deepEqual(Buffer.from(getSlideBackgroundImageBytes(slide)), bytes);
      await jp.getByRole('button', { name: 'スライドの画面切り替え', exact: true }).click();
      const dialog = editor.getByRole('dialog');
      await dialog.getByText('選択したスライドに適用: 2', { exact: true }).waitFor();
      await dialog.locator('select').first().selectOption('fade');
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await saved();
      assert.deepEqual(
        (await slides()).map((slide) => getSlideTransition(slide)?.effect ?? null),
        ['fade', 'fade', null],
      );
      await undo();
      assert.deepEqual((await slides()).map(getSlideTransition), [null, null, null]);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.screenshot({ path: '/tmp/pptx-pr287-slide-batch-ja.png', fullPage: true });
      await page.reload();
      await saved();
      assert.deepEqual(
        (await slides()).map((slide) => getSlideTransition(slide)?.effect ?? null),
        ['fade', 'fade', null],
      );
      for (const slide of (await slides()).slice(0, 2))
        assert.deepEqual(Buffer.from(getSlideBackgroundImageBytes(slide)), bytes);
      assert.deepEqual(errors, []);
    } catch (error) {
      await page?.screenshot({ path: '/tmp/pptx-pr287-slide-batch-failure.png', fullPage: true });
      throw error;
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
