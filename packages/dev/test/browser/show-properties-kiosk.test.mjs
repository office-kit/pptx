import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import { getSlideShapes, getSlides, savePresentation, setShapeClickAction } from '@office-kit/pptx';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import { startPreview } from '../helpers/server.mjs';

test(
  'kiosk playback loops timed slides until Escape even without the loop flag',
  { timeout: 30000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-kiosk-'));
    const file = join(dir, 'deck.tsx');
    await writeFile(
      file,
      `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation>{['First','Last'].map(text => <Slide><Text x={1} y={1} width={7} height={1}>{text}</Text></Slide>)}</Presentation>`,
    );
    let preview, browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage();
      await page.route('**/state*', async (route) => {
        const response = await route.fetch();
        const body = await response.json();
        body.showProperties = {
          mode: { kind: 'kiosk', restart: 300000 },
          slides: { kind: 'all' },
          loop: false,
          showNarration: false,
          showAnimation: false,
          useTimings: true,
        };
        body.transitions = [{ advanceAfterMs: 400 }, { advanceAfterMs: 400 }];
        await route.fulfill({ response, body: JSON.stringify(body) });
      });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('Last'),
      );
      await page.waitForFunction(
        () => document.querySelector('#slide')?.shadowRoot?.textContent?.includes('First'),
        null,
        { timeout: 3000 },
      );
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => !document.body.classList.contains('presenting'));
      await page.waitForFunction(() => document.fullscreenElement === null);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'kiosk presentation ignores manual navigation and blank clicks but follows slide hyperlinks',
  { timeout: 30000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-kiosk-input-'));
    const file = join(dir, 'deck.tsx');
    const deck = await compile(
      Presentation({
        children: [
          Slide({ children: Text({ x: 1, y: 1, width: 7, height: 1, children: 'First' }) }),
          Slide({ children: Text({ x: 1, y: 1, width: 7, height: 1, children: 'Second' }) }),
        ],
      }),
    );
    const slides = getSlides(deck);
    const firstSlide = slides[0];
    const targetSlide = slides[1];
    assert.ok(firstSlide && targetSlide);
    const firstShape = getSlideShapes(firstSlide)[0];
    assert.ok(firstShape);
    setShapeClickAction(firstShape, { kind: 'slide', slide: targetSlide });
    await writeFile(join(dir, 'source.pptx'), await savePresentation(deck));
    await writeFile(
      file,
      `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(join(dir, 'source.pptx'))})} />;`,
    );
    let preview, browser;
    try {
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
      await page.route('**/state*', async (route) => {
        const response = await route.fetch();
        const body = await response.json();
        body.showProperties = {
          mode: { kind: 'kiosk', restart: 300000 },
          slides: { kind: 'all' },
          loop: true,
          showNarration: false,
          showAnimation: false,
          useTimings: false,
        };
        await route.fulfill({ response, body: JSON.stringify(body) });
      });
      await page.goto(preview.url);
      await page.getByRole('button', { name: 'Preview', exact: true }).click();
      await page.getByRole('button', { name: 'Present', exact: true }).click();
      const slideText = () =>
        page.locator('#slide').evaluate((node) => node.shadowRoot?.textContent ?? '');
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('First'),
      );
      const slide = page.locator('#slide');
      const box = await slide.boundingBox();
      assert.ok(box);
      await page.mouse.click(box.x + 20, box.y + 20);
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('PageDown');
      await page.keyboard.press('Home');
      await page.keyboard.press('End');
      await page.keyboard.press(' ');
      await page.waitForTimeout(100);
      assert.match(await slideText(), /First/);
      await page.locator('#slide a').click();
      await page.waitForFunction(() =>
        document.querySelector('#slide')?.shadowRoot?.textContent?.includes('Second'),
      );
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => !document.body.classList.contains('presenting'));
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
