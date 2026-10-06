import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideMedia,
  createPresentation,
  inches,
  savePresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const mediaBytes = {
  audio: Uint8Array.from([0x49, 0x44, 0x33, 0x04, 0, 0, 0, 0, 0, 0, 0x11, 0x22, 0x33]),
  video: Uint8Array.from([
    0, 0, 0, 0x18, 0x66, 0x74, 0x79, 0x70, 0x6d, 0x70, 0x34, 0x32, 0, 0, 0, 0, 0x6d, 0x70, 0x34,
    0x32,
  ]),
};

for (const kind of ['audio', 'video']) {
  test(
    `Playback ribbon saves embedded ${kind} with original bytes`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-media-save-as-'));
      let browser;
      let preview;
      try {
        const presentation = createPresentation();
        const slide = addBlankSlide(presentation);
        addSlideMedia(slide, {
          kind,
          data: mediaBytes[kind],
          x: inches(1),
          y: inches(1),
          w: inches(4),
          h: inches(2),
        });
        const sourcePath = join(dir, 'source.pptx');
        await writeFile(sourcePath, await savePresentation(presentation));
        const entry = join(dir, 'deck.tsx');
        await writeFile(
          entry,
          `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(sourcePath)})} />;`,
        );
        preview = await startPreview(entry);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({
          acceptDownloads: true,
          viewport: { width: 1500, height: 900 },
        });
        await page.goto(preview.url);
        const editor = page.frameLocator('#editor-frame');
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        await editor.locator('.hit').first().click();
        await editor.getByRole('tab', { name: 'Playback', exact: true }).click();
        const saveButton = editor.getByRole('button', { name: 'Save Media As', exact: true });
        assert.equal(await saveButton.isEnabled(), true);
        const downloadPromise = page.waitForEvent('download', { timeout: 10000 });
        await saveButton.click();
        const download = await downloadPromise;
        assert.equal(download.suggestedFilename(), kind === 'audio' ? 'media1.mp3' : 'media1.mp4');
        assert.deepEqual(Uint8Array.from(await readFile(await download.path())), mediaBytes[kind]);
      } finally {
        // Close the browser before the preview it is connected to; each step
        // still runs when an earlier one fails.
        try {
          await browser?.close();
        } finally {
          try {
            await preview?.close();
          } finally {
            await rm(dir, { recursive: true, force: true });
          }
        }
      }
    },
  );
}
