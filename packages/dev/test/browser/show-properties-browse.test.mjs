import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { startPreview } from '../helpers/server.mjs';

async function makeDeck() {
  const dir = await mkdtemp(join(tmpdir(), 'office-show-browse-'));
  const file = join(dir, 'deck.tsx');
  await writeFile(
    file,
    `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={1}>Browse</Text></Slide></Presentation>`,
  );
  return { dir, file };
}

for (const mode of [
  { kind: 'browse', showScrollbar: false },
  { kind: 'browse', showScrollbar: true },
  { kind: 'present' },
]) {
  const windowed = mode.kind === 'browse';
  test(
    `slideshow fullscreen behavior respects ${JSON.stringify(mode)}`,
    { timeout: 60000 },
    async () => {
      const { dir, file } = await makeDeck();
      let preview, browser;
      try {
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
        await page.route('**/state*', async (route) => {
          const response = await route.fetch();
          const body = await response.json();
          body.showProperties = {
            mode,
            slides: { kind: 'range', start: 1, end: 1 },
            loop: false,
            showNarration: false,
            showAnimation: false,
            useTimings: false,
          };
          await route.fulfill({ response, body: JSON.stringify(body) });
        });
        await page.goto(preview.url);
        await page.evaluate(() => {
          window.__fullscreenRequests = 0;
          document.documentElement.requestFullscreen = async () => {
            window.__fullscreenRequests += 1;
          };
        });
        await page.getByRole('button', { name: 'Preview', exact: true }).click();
        await page.getByRole('button', { name: 'Present', exact: true }).click();
        await page.waitForFunction(() => document.body.classList.contains('presenting'));
        assert.equal(await page.evaluate(() => window.__fullscreenRequests), windowed ? 0 : 1);
        if (windowed) assert.equal(await page.evaluate(() => document.fullscreenElement), null);
        await page.keyboard.press('Escape');
        await page.waitForFunction(() => !document.body.classList.contains('presenting'));
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}
