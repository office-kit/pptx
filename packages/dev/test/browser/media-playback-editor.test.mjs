import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addBlankSlide,
  addSlideMedia,
  createPresentation,
  getShapeId,
  getShapeMediaPlayback,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

for (const [kind, nested] of [
  ['audio', false],
  ['video', false],
  ['audio', true],
  ['video', true],
]) {
  test(
    `Playback ribbon edits and persists ${nested ? 'nested ' : ''}${kind} options with undo`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-media-editor-'));
      let browser, preview;
      try {
        const pres = createPresentation();
        const slide = addBlankSlide(pres);
        const ascii = (value) => Array.from(value, (char) => char.charCodeAt(0));
        // Container headers suffice for editing; decoding is covered by playback tests.
        const data =
          kind === 'audio'
            ? new Uint8Array([...ascii('ID3'), 3, 0, 0, 0, 0, 0, 0, 0xff, 0xfb])
            : new Uint8Array([
                0,
                0,
                0,
                0x18,
                ...ascii('ftypmp42'),
                0,
                0,
                0,
                0,
                ...ascii('mp42isom'),
              ]);
        addSlideMedia(slide, {
          kind,
          data,
          x: inches(1),
          y: inches(1),
          w: inches(4),
          h: inches(2),
        });
        let source = await savePresentation(pres);
        if (nested) {
          const parts = unzipSync(source);
          const slidePath = 'ppt/slides/slide1.xml';
          const xml = strFromU8(parts[slidePath]);
          const media = new RegExp(`<p:${kind}\\b[^>]*>[\\s\\S]*?</p:${kind}>`);
          assert.ok(media.test(xml));
          parts[slidePath] = strToU8(
            xml.replace(
              media,
              (node) =>
                `<p:par><p:cTn id="3" dur="indefinite"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${node}</p:childTnLst></p:cTn></p:par>`,
            ),
          );
          source = zipSync(parts);
        }
        await writeFile(join(dir, 'source.pptx'), source);
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
        page.setDefaultTimeout(5000);
        await page.goto(preview.url);
        const editor = page.frameLocator('#editor-frame');
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.equal(await editor.getByRole('tab', { name: 'Playback', exact: true }).count(), 0);
        await editor.locator('.hit').first().click();
        await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
        const panel = editor.locator('#ribbon-panel');
        await panel.getByLabel('Start', { exact: true }).selectOption('automatic');
        await panel.getByLabel('Loop until stopped', { exact: true }).check();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        const readPlayback = async () => {
          const saved = await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          );
          return getShapeMediaPlayback(getSlideShapes(getSlides(saved)[0])[0]);
        };
        assert.equal((await readPlayback()).autoplay, true);
        assert.equal((await readPlayback()).loop, true);
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.equal((await readPlayback()).loop, false);
        assert.equal((await readPlayback()).autoplay, true);
        assert.equal(
          await panel.getByLabel('Play Full Screen', { exact: true }).count(),
          kind === 'video' ? 1 : 0,
        );
        if (kind === 'video') {
          await panel.getByLabel('Play Full Screen', { exact: true }).check();
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
          assert.equal((await readPlayback()).fullScreen, true);
        }
        const rewind = panel.getByLabel('Rewind After Playing', { exact: true });
        assert.equal(await rewind.isChecked(), false);
        await rewind.check();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.equal((await readPlayback()).rewindAfterPlaying, true);
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.equal((await readPlayback()).rewindAfterPlaying, undefined);
        await rewind.check();
        const fadeIn = panel.getByLabel('Fade In', { exact: true });
        const fadeOut = panel.getByLabel('Fade Out', { exact: true });
        await fadeIn.fill('0.15');
        await fadeIn.press('Tab');
        await fadeOut.fill('0.25');
        await fadeOut.press('Tab');
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.deepEqual((await readPlayback()).fade, { inMs: 150, outMs: 250 });
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.deepEqual((await readPlayback()).fade, { inMs: 150, outMs: 0 });
        const delay = panel.getByLabel('Start delay (seconds)', { exact: true });
        await delay.fill('1.001');
        await delay.press('Tab');
        const volume = panel.getByLabel('Volume', { exact: true });
        const chooseVolume = async (name) => {
          await volume.click();
          await panel.getByRole('menuitemradio', { name, exact: true }).click();
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
        };
        await chooseVolume('Low');
        assert.equal((await readPlayback()).volume, 0.2);
        await chooseVolume('Medium');
        assert.equal((await readPlayback()).volume, 0.5);
        await chooseVolume('High');
        assert.equal((await readPlayback()).volume, 0.8);
        await volume.click();
        assert.equal(
          await panel
            .getByRole('menuitemradio', { name: 'High', exact: true })
            .getAttribute('aria-checked'),
          'true',
        );
        await panel.getByRole('menuitemradio', { name: 'High', exact: true }).press('Escape');
        assert.equal(await panel.getByRole('menu').count(), 0);
        assert.equal(
          await volume.evaluate((node) => node === node.ownerDocument.activeElement),
          true,
        );
        await volume.click();
        await panel.getByRole('menuitemradio', { name: 'Low', exact: true }).focus();
        await panel.getByRole('menuitemradio', { name: 'Low', exact: true }).press('ArrowDown');
        await panel.getByRole('menuitemradio', { name: 'Medium', exact: true }).press('Enter');
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.equal((await readPlayback()).volume, 0.5);
        await chooseVolume('High');
        await volume.click();
        await panel.getByRole('menuitemradio', { name: 'Mute', exact: true }).click();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        await panel.getByLabel('Hide when not playing', { exact: true }).check();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        const settings = await readPlayback();
        assert.equal(settings.delayMs, 1001);
        assert.equal(settings.volume, 0.8);
        assert.equal(settings.muted, true);
        assert.equal(settings.hideWhenStopped, true);
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.equal((await readPlayback()).muted, true);
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.equal((await readPlayback()).muted, false);
        await volume.click();
        await panel.getByRole('menuitemradio', { name: 'Mute', exact: true }).click();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        await volume.click();
        await panel.getByRole('menuitemradio', { name: 'Mute', exact: true }).click();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        assert.equal((await readPlayback()).muted, true);
        await chooseVolume('High');
        await page.reload();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        await editor.locator('.hit').first().click();
        await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
        assert.equal(await panel.getByLabel('Start', { exact: true }).inputValue(), 'automatic');
        assert.equal(await delay.inputValue(), '1.001');
        assert.equal(await rewind.isChecked(), true);
        assert.equal((await readPlayback()).volume, 0.8);
        await editor.locator('.lang select').selectOption('ja');
        await panel.getByRole('button', { name: '音量', exact: true }).click();
        await panel.getByRole('menuitemradio', { name: '小', exact: true }).click();
        await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
        assert.equal((await readPlayback()).volume, 0.2);
        await panel.getByLabel('再生が終了したら巻き戻す', { exact: true }).uncheck();
        await panel.getByLabel('開始', { exact: true }).selectOption('click');
        await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
        assert.equal((await readPlayback()).autoplay, false);
        assert.equal((await readPlayback()).rewindAfterPlaying, undefined);
        assert.equal(await panel.getByLabel('開始の遅延（秒）', { exact: true }).count(), 0);
        await page.screenshot({
          path: `/tmp/pptx-media-editor-${nested ? 'nested-' : ''}${kind}.png`,
        });
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}

test(
  'native background timing converts to editable start mode with save, reload, and undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-native-media-editor-'));
    let browser;
    let preview;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const ascii = (value) => Array.from(value, (char) => char.charCodeAt(0));
      addSlideMedia(slide, {
        kind: 'audio',
        data: new Uint8Array([...ascii('ID3'), 3, 0, 0, 0, 0, 0, 0, 0xff, 0xfb]),
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2),
      });
      const [audio] = getSlideShapes(getSlides(pres)[0]);
      const parts = unzipSync(await savePresentation(pres));
      const slidePath = 'ppt/slides/slide1.xml';
      const timing = (
        await readFile(
          new URL('../../../../test/fixtures/native-media-background-timing.xml', import.meta.url),
          'utf8',
        )
      ).replaceAll('spid="2"', `spid="${getShapeId(audio)}"`);
      const slideXml = strFromU8(parts[slidePath]);
      parts[slidePath] = strToU8(
        /<p:timing\b/.test(slideXml)
          ? slideXml.replace(/<p:timing\b[\s\S]*?<\/p:timing>/, timing)
          : slideXml.replace('</p:sld>', `${timing}</p:sld>`),
      );
      const source = join(dir, 'source.pptx');
      await writeFile(source, zipSync(parts));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      page.setDefaultTimeout(5000);
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
      const saveChange = async (action, autoplay) => {
        const saved = page.waitForResponse(
          (response) =>
            response.url().endsWith('/editor/document') && response.request().method() === 'PUT',
        );
        await action();
        assert.equal((await saved).ok(), true);
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        const document = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        assert.equal(
          getShapeMediaPlayback(getSlideShapes(getSlides(document)[0])[0]).autoplay,
          autoplay,
        );
      };
      const start = editor.getByLabel('Start', { exact: true });
      assert.equal(await start.locator('option[value="click"]').textContent(), 'When Clicked On');
      assert.equal(await start.inputValue(), 'automatic');
      await saveChange(() => start.selectOption('click'), false);
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
      assert.equal(await editor.getByLabel('Start', { exact: true }).inputValue(), 'click');
      await saveChange(
        () => editor.getByLabel('Start', { exact: true }).selectOption('automatic'),
        true,
      );
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
      assert.equal(await editor.getByLabel('Start', { exact: true }).inputValue(), 'automatic');
      await saveChange(
        () => editor.getByLabel('Start', { exact: true }).selectOption('click'),
        false,
      );
      await saveChange(() => editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click(), true);
      assert.equal(await editor.getByLabel('Start', { exact: true }).inputValue(), 'automatic');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'audio exposes Play Across Slides, persists the range, and hides it for video',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-media-cross-slide-editor-'));
    let browser, preview;
    try {
      const pres = createPresentation();
      const audioSlide = addBlankSlide(pres);
      const videoSlide = addBlankSlide(pres);
      const ascii = (value) => Array.from(value, (char) => char.charCodeAt(0));
      addSlideMedia(audioSlide, {
        kind: 'audio',
        data: new Uint8Array([...ascii('ID3'), 3, 0, 0, 0, 0, 0, 0, 0xff, 0xfb]),
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2),
      });
      addSlideMedia(videoSlide, {
        kind: 'video',
        data: new Uint8Array([
          0,
          0,
          0,
          0x18,
          ...ascii('ftypmp42'),
          0,
          0,
          0,
          0,
          ...ascii('mp42isom'),
        ]),
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(2),
      });
      const source = join(dir, 'source.pptx');
      await writeFile(source, await savePresentation(pres));
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
      page.setDefaultTimeout(5000);
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();

      const readPlayback = async (slideIndex) => {
        const saved = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getShapeMediaPlayback(getSlideShapes(getSlides(saved)[slideIndex])[0]);
      };

      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
      const background = editor.getByRole('button', { name: 'Play in Background', exact: true });
      assert.equal(await background.count(), 1);
      const volume = editor.getByLabel('Volume', { exact: true });
      await volume.click();
      await editor.getByRole('menuitemradio', { name: 'Low', exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.getByLabel('Rewind After Playing', { exact: true }).check();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await background.click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const backgroundPlayback = await readPlayback(0);
      assert.equal(backgroundPlayback.autoplay, true);
      assert.equal(backgroundPlayback.slideCount, 999);
      assert.equal(backgroundPlayback.loop, true);
      assert.equal(backgroundPlayback.hideWhenStopped, true);
      assert.equal(backgroundPlayback.rewindAfterPlaying, true);
      assert.equal(backgroundPlayback.volume, 0.2);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      const undoneBackground = await readPlayback(0);
      assert.equal(undoneBackground.autoplay, false);
      assert.equal(undoneBackground.slideCount, undefined);
      assert.equal(undoneBackground.loop, false);
      assert.equal(undoneBackground.hideWhenStopped, false);
      assert.equal(undoneBackground.rewindAfterPlaying, true);
      assert.equal(undoneBackground.volume, 0.2);
      await background.click();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
      assert.equal(
        await editor.getByRole('button', { name: 'Play in Background', exact: true }).count(),
        1,
      );
      const across = editor.getByLabel('Play Across Slides', { exact: true });
      assert.equal(await across.count(), 1);
      assert.equal(await across.isChecked(), true);
      assert.equal(await editor.getByLabel('Play Full Screen', { exact: true }).count(), 0);

      await across.uncheck();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal((await readPlayback(0)).slideCount, undefined);

      await across.check();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await page.reload();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
      const reloadedAcross = editor.getByLabel('Play Across Slides', { exact: true });
      assert.equal(await reloadedAcross.isChecked(), true);
      await reloadedAcross.uncheck();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      assert.equal((await readPlayback(0)).slideCount, undefined);
      await reloadedAcross.check();
      await editor.getByText('Saved to this project', { exact: true }).waitFor();

      await editor.locator('.lang select').selectOption('ja');
      const acrossJapanese = editor.getByLabel('スライド切り替え後も再生', { exact: true });
      assert.equal(await acrossJapanese.count(), 1);
      assert.equal(await acrossJapanese.isChecked(), true);
      assert.equal(
        await editor.getByRole('button', { name: 'バックグラウンドで再生', exact: true }).count(),
        1,
      );

      await editor.getByRole('button', { name: 'スライド 2', exact: true }).click();
      await editor.locator('.hit').first().click();
      await editor.getByRole('tab', { name: '再生', exact: true }).click();
      assert.equal(await editor.getByLabel('スライド切り替え後も再生', { exact: true }).count(), 0);
      assert.equal(
        await editor.getByRole('button', { name: 'バックグラウンドで再生', exact: true }).count(),
        0,
      );
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
