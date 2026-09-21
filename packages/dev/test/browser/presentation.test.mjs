import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import {
  getSlides,
  getSlideShapes,
  setShapeClickAction,
  savePresentation,
  setSlideNotes,
  setSlideHidden,
  setSlideTransition,
} from '@office-kit/pptx';
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

test(
  'skipped slides remain editable but are omitted from playback and presenter navigation',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-skipped-'));
    let preview;
    let browser;
    try {
      const deck = await compile(
        Presentation({
          children: [
            'Hidden first',
            'Visible opening',
            'Hidden middle',
            'Visible closing',
            'Hidden last',
          ].map((title) =>
            Slide({ children: Text({ x: 1, y: 1, width: 5, height: 1, children: title }) }),
          ),
        }),
      );
      const slides = getSlides(deck);
      for (const i of [0, 2, 4]) setSlideHidden(slides[i], true);
      setSlideTransition(slides[1], { effect: 'none', advanceAfterMs: 1200 });
      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const context = await browser.newContext({ viewport: { width: 1500, height: 1000 } });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      const skip = editor.getByRole('checkbox', { name: 'Skip during presentation', exact: true });
      await skip.waitFor();
      assert.equal(await skip.isChecked(), true);
      assert.equal(await editor.locator('.thumb-row.skipped').count(), 3);
      await skip.uncheck();
      await page.waitForFunction(() => state.hiddenSlides?.[0] === false);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await page.waitForFunction(() => state.hiddenSlides?.[0] === true);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await page.waitForFunction(() => state.hiddenSlides?.[0] === false);
      await skip.check();
      await page.waitForFunction(() => state.hiddenSlides?.[0] === true);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      assert.equal(await page.locator('#count').textContent(), 'Slide 1 of 5');
      assert.equal(await page.locator('.thumbnail[data-skipped="true"]').count(), 3);
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      assert.equal(await page.locator('#count').textContent(), 'Slide 2 of 5');
      assert.equal(await page.locator('#present-prev').isDisabled(), true);
      await page.waitForFunction(() => index === 3); // Automatic advance skips the middle slide.
      assert.equal(await page.locator('#present-next').isDisabled(), true);
      await page.keyboard.press('ArrowRight');
      assert.equal(await page.locator('#count').textContent(), 'Slide 4 of 5');
      await page.keyboard.press('Home');
      assert.equal(await page.locator('#count').textContent(), 'Slide 2 of 5');
      await page.keyboard.press('End');
      assert.equal(await page.locator('#count').textContent(), 'Slide 4 of 5');
      await page.keyboard.press('ArrowLeft');
      assert.equal(await page.locator('#count').textContent(), 'Slide 2 of 5');
      await page.locator('#stage').click({ position: { x: 15, y: 15 } });
      assert.equal(await page.locator('#count').textContent(), 'Slide 4 of 5');
      await page.keyboard.press('Escape');
      await page.getByRole('button', { name: 'Slide 1', exact: true }).click();
      const ready = page.waitForEvent('popup');
      await page.getByRole('button', { name: 'Presenter view', exact: true }).click();
      const presenter = await ready;
      presenter.on('pageerror', (error) => errors.push(error.message));
      await presenter.getByText('Slide 2 of 5', { exact: true }).waitFor();
      assert.ok((await presenter.locator('#next svg').textContent()).includes('Visible closing'));
      assert.equal(
        await presenter.getByRole('button', { name: 'Previous', exact: true }).isDisabled(),
        true,
      );
      await presenter.getByRole('button', { name: 'Next', exact: true }).click();
      await presenter.getByText('Slide 4 of 5', { exact: true }).waitFor();
      assert.equal(
        await presenter.getByRole('button', { name: 'Next', exact: true }).isDisabled(),
        true,
      );
      await presenter.getByRole('button', { name: 'Exit presentation', exact: true }).click();
      await page.getByRole('button', { name: 'Edit', exact: true }).click();
      await editor.locator('.lang select').selectOption('ja');
      for (let i = 0; i < 5; i++) {
        await editor.locator('.thumb-row').nth(i).click();
        await editor
          .getByRole('checkbox', { name: 'プレゼンテーションでスキップ', exact: true })
          .check();
      }
      await page.waitForFunction(
        () => state.hiddenSlides?.length === 5 && state.hiddenSlides.every(Boolean),
      );
      await page.screenshot({ path: '/tmp/pptx-pr287-skip-ja.png', fullPage: true });
      await page.getByRole('button', { name: 'プレビュー', exact: true }).click();
      assert.equal(await page.locator('#present').isDisabled(), true);
      assert.equal(await page.locator('#presenter').isDisabled(), true);
      await page.reload();
      await page.waitForFunction(
        () => state.hiddenSlides?.length === 5 && state.hiddenSlides.every(Boolean),
      );
      assert.equal(await page.locator('#present').isDisabled(), true);
      assert.equal(await page.locator('.thumbnail').count(), 5);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'slide links navigate in preview and presentation without also advancing',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-slide-links-'));
    let preview;
    let browser;
    try {
      const deck = await compile(
        Presentation({
          children: ['Jump to third', 'Second', 'Back to first', 'Last'].map((title) =>
            Slide({ children: Text({ x: 1, y: 1, width: 5, height: 1, children: title }) }),
          ),
        }),
      );
      const slides = getSlides(deck);
      setShapeClickAction(getSlideShapes(slides[0])[0], { kind: 'slide', slide: slides[2] });
      setShapeClickAction(getSlideShapes(slides[2])[0], { kind: 'slide', slide: slides[0] });
      setShapeClickAction(getSlideShapes(slides[1])[0], {
        kind: 'url',
        url: 'https://example.com/link-test',
      });
      setSlideNotes(slides[2], 'Linked destination notes / 移動先のノート');
      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
      await context.route('https://example.com/**', (route) =>
        route.fulfill({ body: 'External destination' }),
      );
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.locator('#slide a').click();
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 4');
      await page.locator('#slide a').focus();
      await page.keyboard.press('Enter');
      assert.equal(await page.locator('#count').textContent(), 'Slide 1 of 4');
      assert.equal(new URL(page.url()).hash, '');
      const presenterReady = page.waitForEvent('popup');
      await page.getByRole('button', { name: 'Presenter view', exact: true }).click();
      const presenter = await presenterReady;
      presenter.on('pageerror', (error) => errors.push(error.message));
      await presenter.locator('#current a').click();
      await presenter.getByText('Slide 3 of 4', { exact: true }).waitFor();
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 4');
      assert.equal(
        await presenter.locator('#notes').textContent(),
        'Linked destination notes / 移動先のノート',
      );
      await presenter.locator('#current a').focus();
      await presenter.keyboard.press('Enter');
      await presenter.getByText('Slide 1 of 4', { exact: true }).waitFor();
      assert.equal(await page.locator('#count').textContent(), 'Slide 1 of 4');
      assert.equal(new URL(presenter.url()).hash, '');
      // A stale or malformed jump cannot clamp to an unrelated slide.
      await presenter.evaluate(() => {
        send('jump', -1);
        send('jump', 99);
        send('jump', '2');
        send('jump', 0.5);
        send('ready');
      });
      await presenter.getByText('Slide 1 of 4', { exact: true }).waitFor();
      assert.equal(await page.locator('#count').textContent(), 'Slide 1 of 4');
      await presenter.getByRole('button', { name: 'Exit presentation', exact: true }).click();
      await page.waitForFunction(() => !presenting);
      await presenter.close();

      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.locator('#slide a').click();
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 4');
      await page.locator('#slide a').click();
      assert.equal(await page.locator('#count').textContent(), 'Slide 1 of 4');
      await page.locator('#stage').click({ position: { x: 15, y: 15 } });
      assert.equal(await page.locator('#count').textContent(), 'Slide 2 of 4');
      const opened = page.waitForEvent('popup');
      await page.locator('#slide a').click();
      const external = await opened;
      await external.waitForLoadState();
      assert.equal(external.url(), 'https://example.com/link-test');
      assert.equal(await page.locator('#count').textContent(), 'Slide 2 of 4');
      await external.close();
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
  'navigation presets skip hidden slides and work from presenter next-slide preview',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-navigation-'));
    let preview, browser;
    try {
      const deck = await compile(
        Presentation({
          children: ['Next', 'Hidden', 'Previous', 'First'].map((title) =>
            Slide({ children: Text({ x: 1, y: 1, width: 5, height: 1, children: title }) }),
          ),
        }),
      );
      const slides = getSlides(deck);
      setSlideHidden(slides[1], true);
      for (const [index, kind] of [
        [0, 'nextSlide'],
        [2, 'prevSlide'],
        [3, 'firstSlide'],
      ])
        setShapeClickAction(getSlideShapes(slides[index])[0], { kind });
      await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.locator('#slide a').click();
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 4');
      await page.locator('#slide a').focus();
      await page.keyboard.press('Enter');
      assert.equal(await page.locator('#count').textContent(), 'Slide 1 of 4');
      const ready = page.waitForEvent('popup');
      await page.getByRole('button', { name: 'Presenter view', exact: true }).click();
      const presenter = await ready;
      presenter.on('pageerror', (error) => errors.push(error.message));
      await presenter.locator('#next a').click();
      await presenter.getByText('Slide 1 of 4', { exact: true }).waitFor();
      await presenter.locator('#current a').click();
      await presenter.getByText('Slide 3 of 4', { exact: true }).waitFor();
      assert.equal(await page.locator('#count').textContent(), 'Slide 3 of 4');
      await presenter.locator('#current a').focus();
      await presenter.keyboard.press('Enter');
      await presenter.getByText('Slide 1 of 4', { exact: true }).waitFor();
      assert.equal(await page.locator('#count').textContent(), 'Slide 1 of 4');
      await presenter.getByRole('button', { name: 'Exit presentation', exact: true }).click();
      await page.waitForFunction(() => !presenting);
      await presenter.close();
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
