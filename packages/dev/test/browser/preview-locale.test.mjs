import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

test(
  'preview controls and accessible labels follow the editor language across reloads',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-preview-locale-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['First','Second'].map(text=><Slide><Text x={1} y={1} width={5} height={1}>{text}</Text></Slide>)}</Presentation>;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({
        viewport: { width: 1280, height: 800 },
        locale: 'en-US',
      });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.lang select').selectOption('ja');
      await page.getByRole('button', { name: 'プレビュー', exact: true }).click();
      await page.getByRole('button', { name: 'プレゼンテーション', exact: true }).waitFor();
      assert.equal(await page.locator('html').getAttribute('lang'), 'ja');
      assert.equal(await page.locator('#count').textContent(), 'スライド 1 / 2');
      assert.equal(await page.locator('#status').textContent(), '2 スライド・ライブ');
      assert.equal(
        await page.getByRole('link', { name: 'PPTXをダウンロード', exact: true }).count(),
        1,
      );
      assert.equal(
        await page.getByRole('navigation', { name: 'スライド', exact: true }).count(),
        1,
      );
      assert.equal(
        await page.locator('#editor-frame').getAttribute('title'),
        'プレゼンテーション編集',
      );
      await page.getByRole('button', { name: '次のスライド', exact: true }).click();
      assert.equal(await page.locator('#count').textContent(), 'スライド 2 / 2');
      await page.reload();
      await page.getByRole('button', { name: 'プレゼンテーション', exact: true }).waitFor();
      await page.waitForFunction(() => state.slides.length === 2);
      assert.equal(await page.locator('html').getAttribute('lang'), 'ja');
      assert.equal(
        await page.locator('#editor-frame').getAttribute('src'),
        null,
        'preview reload does not need to load the editor to discover its language',
      );
      await page.getByRole('button', { name: 'プレゼンテーション', exact: true }).click();
      await page.keyboard.press('ArrowRight');
      assert.equal(await page.locator('#present-count').textContent(), 'スライド 2 / 2');
      await page.getByRole('button', { name: '終了 · Esc', exact: true }).click();
      const failState = (route) => route.fulfill({ status: 503, body: 'Unavailable' });
      await page.route('**/state*', failState);
      await page.evaluate(() => refresh());
      assert.equal(await page.locator('#status').textContent(), '再接続中…');
      await page.unroute('**/state*', failState);
      await page.evaluate(() => refresh());
      assert.equal(await page.locator('#status').textContent(), '2 スライド・ライブ');
      await page.getByRole('button', { name: '編集', exact: true }).click();
      await editor.locator('.lang select').selectOption('en');
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      assert.equal(await page.locator('html').getAttribute('lang'), 'en');
      assert.equal(await page.getByRole('button', { name: 'Present', exact: true }).count(), 1);
      assert.equal(await page.locator('#status').textContent(), '2 slides · Live');
      assert.equal(await page.locator('.thumbnail').first().getAttribute('aria-label'), 'Slide 1');
      assert.equal(await page.locator('#prev').getAttribute('aria-label'), 'Previous slide');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
