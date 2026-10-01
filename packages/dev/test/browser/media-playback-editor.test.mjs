import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  addBlankSlide,
  addSlideMedia,
  createPresentation,
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
        const delay = panel.getByLabel('Start delay (seconds)', { exact: true });
        await delay.fill('1.001');
        await delay.press('Tab');
        const volume = panel.getByLabel('Volume', { exact: true });
        await volume.fill('25');
        await volume.press('Tab');
        await panel.getByLabel('Mute', { exact: true }).check();
        await panel.getByLabel('Hide when not playing', { exact: true }).check();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        const settings = await readPlayback();
        assert.equal(settings.delayMs, 1001);
        assert.equal(settings.volume, 0.25);
        assert.equal(settings.muted, true);
        assert.equal(settings.hideWhenStopped, true);
        await page.reload();
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        await editor.locator('.hit').first().click();
        await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
        assert.equal(await panel.getByLabel('Start', { exact: true }).inputValue(), 'automatic');
        assert.equal(await delay.inputValue(), '1.001');
        await editor.locator('.lang select').selectOption('ja');
        await panel.getByLabel('開始', { exact: true }).selectOption('click');
        await editor.getByText('このプロジェクトに保存済み', { exact: true }).waitFor();
        assert.equal((await readPlayback()).autoplay, false);
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
