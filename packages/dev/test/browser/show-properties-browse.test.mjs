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
    `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['A','B','C'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
  );
  return { dir, file };
}

for (const mode of [
  { kind: 'browse', showScrollbar: false },
  { kind: 'browse', showScrollbar: true },
  { kind: 'present' },
]) {
  const showScrollbar = mode.kind === 'browse' && mode.showScrollbar;
  test(
    `slideshow window and scrollbar behavior respects ${JSON.stringify(mode)}`,
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
            slides: showScrollbar
              ? { kind: 'customShow', id: 7 }
              : { kind: 'range', start: 1, end: 3 },
            loop: false,
            showNarration: false,
            showAnimation: false,
            useTimings: false,
          };
          if (showScrollbar) {
            body.customShows = [{ id: 7, name: 'Repeated', slideIndices: [2, 1, 0, 2] }];
            body.hiddenSlides = [false, true, false];
          }
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
        assert.equal(
          await page.evaluate(() => window.__fullscreenRequests),
          mode.kind === 'browse' ? 0 : 1,
        );
        assert.equal(await page.evaluate(() => document.fullscreenElement), null);
        const scrollbar = page.locator('#presentation-scrollbar');
        assert.equal(await scrollbar.isVisible(), showScrollbar);
        if (showScrollbar) {
          assert.equal(await scrollbar.getAttribute('aria-valuemax'), '2');
          await scrollbar.focus();
          await page.keyboard.press('ArrowDown');
          assert.equal(await scrollbar.getAttribute('aria-valuenow'), '1');
          assert.match(
            await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
            /A/,
          );
          await page.keyboard.press('End');
          assert.equal(await scrollbar.getAttribute('aria-valuenow'), '2');
          assert.match(
            await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
            /C/,
          );
          const track = await scrollbar.boundingBox();
          const thumb = await scrollbar.locator('span').boundingBox();
          assert.ok(
            thumb.y + thumb.height <= track.y + track.height + 1,
            'thumb must stay within the track',
          );
          await page.mouse.move(thumb.x + thumb.width / 2, thumb.y + thumb.height / 2);
          await page.mouse.down();
          await page.mouse.move(thumb.x + thumb.width / 2, track.y + 1, { steps: 5 });
          await page.mouse.up();
          assert.equal(await scrollbar.getAttribute('aria-valuenow'), '0');
          await page.keyboard.press('ArrowDown');
          assert.equal(await scrollbar.getAttribute('aria-valuenow'), '1');
          await page.locator('#present-next').click();
          assert.equal(await scrollbar.getAttribute('aria-valuenow'), '2');
        }
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
