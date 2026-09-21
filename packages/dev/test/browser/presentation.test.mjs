import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import { getSlides, savePresentation, setSlideNotes, setSlideTransition } from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'presentation playback honors saved click and automatic advance settings',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-presentation-'));
    let preview;
    let browser;
    try {
      const deck = await compile(
        Presentation({
          children: ['Timed', 'Manual', 'Long timer', 'End'].map((title) =>
            Slide({ children: Text({ x: 1, y: 1, width: 5, height: 1, children: title }) }),
          ),
        }),
      );
      const slides = getSlides(deck);
      setSlideTransition(slides[0], {
        effect: 'none',
        advanceOnClick: false,
        advanceAfterMs: 1500,
      });
      setSlideTransition(slides[2], {
        effect: 'none',
        advanceOnClick: false,
        advanceAfterMs: 4294967295,
      });
      setSlideTransition(slides[3], { effect: 'none', advanceAfterMs: 0 });
      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() => state.slides.length === 4);
      const persisted = await (await fetch(preview.url + '/state')).json();
      assert.equal(persisted.transitions[0].advanceAfterMs, 1500);
      assert.equal(persisted.transitions[0].advanceOnClick, false);
      const current = () => page.locator('#count').textContent();
      const wait = (ms) =>
        page.evaluate((ms) => new Promise((resolve) => setTimeout(resolve, ms)), ms);
      // Ordinary preview has no running playback timer.
      await wait(1700);
      assert.equal(await current(), 'Slide 1 of 4');
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.locator('#stage').click({ position: { x: 15, y: 15 } });
      assert.equal(await current(), 'Slide 1 of 4');
      const timer = await page.evaluate(() => advanceTimer);
      await page.evaluate(() => refresh());
      assert.equal(
        await page.evaluate(() => advanceTimer),
        timer,
        'unchanged server refresh must not restart the timer',
      );
      await page.waitForFunction(() => index === 1);
      // No timing on this slide: remain until a click.
      await wait(1700);
      assert.equal(await current(), 'Slide 2 of 4');
      await page.locator('#stage').click({ position: { x: 15, y: 15 } });
      assert.equal(await current(), 'Slide 3 of 4');
      await wait(100);
      assert.equal(
        await current(),
        'Slide 3 of 4',
        'large OOXML delays must not overflow to immediate timeouts',
      );
      // Explicit navigation still works when stage click advance is disabled.
      await page.keyboard.press('ArrowRight');
      await wait(100);
      assert.equal(await current(), 'Slide 4 of 4');
      assert.equal(await page.locator('#present-next').isDisabled(), true);
      await page.keyboard.press('Home');
      await page.getByRole('button', { name: 'Exit · Esc', exact: true }).click();
      await wait(1700);
      assert.equal(
        await current(),
        'Slide 1 of 4',
        'leaving presentation cancels automatic advance',
      );
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.waitForFunction(() => index === 1);
      await page.keyboard.press('Escape');
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'bilingual presenter window keeps notes separate and synchronizes slide navigation',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-presenter-'));
    let preview;
    let browser;
    try {
      const deck = await compile(
        Presentation({
          children: ['Opening slide', 'Closing slide'].map((title) =>
            Slide({ children: Text({ x: 1, y: 1, width: 5, height: 1, children: title }) }),
          ),
        }),
      );
      const notes =
        '日本語の発表者ノート 🎉\nEnglish notes\n<script>window.notesExecuted=true</script>';
      setSlideNotes(getSlides(deck)[0], notes);
      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.waitForFunction(() => state.slides.length === 2);
      const popupReady = page.waitForEvent('popup');
      await page.getByRole('button', { name: 'Presenter view', exact: true }).click();
      const presenter = await popupReady;
      presenter.on('pageerror', (error) => errors.push(error.message));
      await presenter.getByText('Slide 1 of 2', { exact: true }).waitFor();
      assert.equal(await presenter.locator('#notes').textContent(), notes);
      assert.equal(await presenter.evaluate(() => window.notesExecuted), undefined);
      assert.equal(await presenter.locator('#current').textContent(), ''); // SVG is isolated in a shadow root.
      assert.ok((await presenter.locator('#current svg').textContent()).includes('Opening slide'));
      assert.ok((await presenter.locator('#next svg').textContent()).includes('Closing slide'));
      assert.equal(await page.getByText(notes, { exact: true }).count(), 0);
      await presenter.getByRole('button', { name: 'Next', exact: true }).click();
      await page.waitForFunction(() => index === 1);
      await presenter.getByText('Slide 2 of 2', { exact: true }).waitFor();
      assert.equal(await presenter.locator('#notes').textContent(), 'No speaker notes.');
      assert.equal(
        await presenter.getByRole('button', { name: 'Next', exact: true }).isDisabled(),
        true,
      );
      await presenter.getByRole('button', { name: 'Previous', exact: true }).click();
      await page.waitForFunction(() => index === 0);
      await presenter.getByRole('button', { name: 'Exit presentation', exact: true }).click();
      await page.waitForFunction(() => !presenting);
      await page.getByRole('button', { name: 'Edit', exact: true }).click();
      await page.frameLocator('#editor-frame').locator('.lang select').selectOption('ja');
      await presenter.getByRole('heading', { name: '発表者ビュー', exact: true }).waitFor();
      const editor = page.frameLocator('#editor-frame');
      await editor.getByRole('button', { name: '発表者ノート', exact: true }).click();
      const dialog = editor.getByRole('dialog', { name: '発表者ノート', exact: true });
      await dialog
        .getByLabel('ノートの内容', { exact: true })
        .fill('保存後のノート / Updated notes');
      await dialog.getByRole('button', { name: '適用', exact: true }).click();
      await presenter.getByText('保存後のノート / Updated notes', { exact: true }).waitFor();
      await page.getByRole('button', { name: 'プレビュー', exact: true }).click();
      await presenter.getByText('スライド 1 / 2', { exact: true }).waitFor();
      assert.equal(
        await presenter.locator('#notes').textContent(),
        '保存後のノート / Updated notes',
      );
      await presenter.screenshot({ path: '/tmp/pptx-pr287-presenter-ja.png', fullPage: true });
      await page.getByRole('button', { name: '発表者ビュー', exact: true }).click();
      await page.waitForFunction(() => presenting);
      assert.equal(context.pages().length, 2, 'reopening uses the existing presenter window');
      await presenter.reload();
      await presenter.getByText('スライド 1 / 2', { exact: true }).waitFor();
      await presenter.getByRole('button', { name: '次へ', exact: true }).click();
      await presenter.getByText('発表者ノートはありません。', { exact: true }).waitFor();
      await presenter.getByRole('button', { name: 'タイマーをリセット', exact: true }).click();
      assert.equal(await presenter.locator('#elapsed').textContent(), '00:00');
      await page.close();
      await presenter
        .getByText('プレゼンテーションのウィンドウは閉じられています。', { exact: true })
        .waitFor();
      assert.equal(
        await presenter.getByRole('button', { name: '前へ', exact: true }).isDisabled(),
        true,
      );
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
