import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getSlideShapes,
  getSlides,
  savePresentation,
  setShapeClickAction,
  setCustomShows,
  setSlideShowProperties,
  getSlideShowProperties,
} from '@office-kit/pptx';
import { compile, Presentation, Slide, Text } from '@office-kit/pptx-dsl';
import { startPreview } from '../helpers/server.mjs';

for (const repeated of [false, true])
  test(
    `show action links follow visit history, exit, and reset history (repeated: ${repeated})`,
    { timeout: 30000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-show-links-'));
      let preview, browser;
      try {
        const deck = await compile(
          Presentation({
            children: ['First', 'Middle', 'Last'].map((name) =>
              Slide({
                children: [
                  Text({ x: 1, y: 1, width: 5, height: 1, children: name }),
                  Text({ x: 1, y: 3, width: 5, height: 1, children: 'Return to visited slide' }),
                  Text({ x: 1, y: 5, width: 5, height: 1, children: 'End show' }),
                ],
              }),
            ),
          }),
        );
        const slides = getSlides(deck);
        for (const slide of slides) {
          const shapes = getSlideShapes(slide);
          setShapeClickAction(shapes[1], { kind: 'lastSlideViewed' });
          setShapeClickAction(shapes[2], { kind: 'endShow' }, { range: { start: 0, end: 8 } });
        }
        setShapeClickAction(getSlideShapes(slides[0])[0], { kind: 'slide', slide: slides[2] });
        if (repeated) {
          setCustomShows(deck, [
            { id: 7, name: 'Repeated', slides: [slides[0], slides[0], slides[2], slides[1]] },
          ]);
          setSlideShowProperties(deck, {
            ...getSlideShowProperties(deck),
            slides: { kind: 'customShow', id: 7 },
          });
        }
        const source = join(dir, 'source.pptx');
        await writeFile(source, await savePresentation(deck));
        const file = join(dir, 'deck.tsx');
        await writeFile(
          file,
          `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
        await page.goto(preview.url);
        await page.getByRole('button', { name: 'Preview', exact: true }).click();
        await page.getByRole('button', { name: 'Present', exact: true }).click();
        const expectSlide = async (name) => {
          await page.waitForFunction(
            (name) => document.querySelector('#slide')?.shadowRoot?.textContent?.includes(name),
            name,
          );
        };
        const back = () => page.locator('#slide a[href="#pptx-last-slide-viewed"]').click();
        await expectSlide('First');
        assert.equal(await page.locator('#slide a[href="#pptx-last-slide-viewed"]').count(), 1);
        await back();
        await expectSlide('First');
        await page.locator('#slide a[href="#slide-3"]').click();
        await expectSlide('Last');
        await back();
        await expectSlide('First');
        await back();
        await expectSlide('Last');
        await page.keyboard.press('ArrowLeft');
        await expectSlide(repeated ? 'First' : 'Middle');
        await back();
        await expectSlide('Last');
        if (repeated) {
          await page.keyboard.press('ArrowLeft');
          await expectSlide('First');
          await page.keyboard.press('ArrowLeft');
          await back();
          await page.keyboard.press('ArrowRight');
          await expectSlide('Last');
        }
        await page.locator('#slide a[href="#pptx-end-show"]').click();
        await page.waitForFunction(
          () => !document.body.classList.contains('presenting') && !document.fullscreenElement,
        );
        await page.getByRole('button', { name: 'Present', exact: true }).click();
        await expectSlide('First');
        await back();
        assert.match(
          await page.locator('#slide').evaluate((node) => node.shadowRoot.textContent),
          /First/,
        );
        await page.keyboard.press('Escape');
        await page.waitForFunction(() => !document.fullscreenElement);
        const popup = page.waitForEvent('popup');
        await page.getByRole('button', { name: 'Presenter view', exact: true }).click();
        const presenter = await popup;
        await presenter.locator('#current a[href="#slide-3"]').click();
        await expectSlide('Last');
        await presenter.locator('#current a[href="#pptx-last-slide-viewed"]').click();
        await expectSlide('First');
        await presenter.locator('#current a[href="#pptx-end-show"]').click();
        await page.waitForFunction(() => !document.body.classList.contains('presenting'));
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
