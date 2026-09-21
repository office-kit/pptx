import { mkdir, writeFile } from 'node:fs/promises';
import { join, dirname } from 'node:path';
import { randomUUID } from 'node:crypto';
import { chromium, type Browser } from 'playwright';

export interface VisualReview {
  prompt: string;
  images: string[];
}
/** Each session compares against its own turn's preview, including offscreen slides. */
export function createVisualReviewer(entry: string, slides: () => string[]) {
  let before: string[] = [];
  let rounds = 0;
  return {
    begin() {
      before = slides().slice();
      rounds = 0;
    },
    async next(): Promise<VisualReview | undefined> {
      const current = slides().slice();
      const changed = current.flatMap((svg, index) =>
        svg !== before[index] ? [{ svg, index }] : [],
      );
      if (!changed.length) return;
      if (rounds >= 3)
        throw new Error(
          'Visual review reached its correction limit. Review the latest slide manually before continuing.',
        );
      const round = rounds + 1;
      let browser: Browser;
      try {
        browser = await chromium.launch({
          channel: process.env.PLAYWRIGHT_CHANNEL || 'chrome',
          headless: true,
        });
      } catch {
        try {
          browser = await chromium.launch({ headless: true });
        } catch {
          throw new Error(
            'Visual review needs Chrome or a Playwright Chromium installation. Install Chrome and retry. The source edit was saved; visual review was not completed.',
          );
        }
      }
      try {
        const directory = join(dirname(entry), '.office-kit', 'reviews', randomUUID());
        await mkdir(directory, { recursive: true });
        const page = await browser.newPage({
          viewport: { width: 1600, height: 1000 },
          deviceScaleFactor: 1,
        });
        // Decks may reference remote assets. Review must not make network requests or execute SVG scripts.
        await page.route('**/*', (route) => route.abort());
        const images: string[] = [];
        for (const { svg, index } of changed) {
          await page.setContent(
            '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; img-src data: blob:; font-src data:"><style>html,body{margin:0}svg{display:block;width:1600px;height:auto}</style>' +
              svg,
          );
          await page.evaluate(() => document.fonts.ready.then(() => undefined));
          const file = join(directory, `slide-${index + 1}.png`);
          await page.locator('svg').first().screenshot({ path: file, timeout: 15000 });
          images.push(file);
        }
        const prompt = `Visual review after the edit (round ${round}/3). Open every attached screenshot (or read each local PNG path below) before finishing. Check clipping, overlap, alignment, spacing, legibility and foreground/background contrast. ${round === 3 ? 'This is the final review: do not edit further; report any remaining defects to the user.' : 'Fix concrete design defects only on these affected slides;'} preserve the user's exact text and unrelated work. Do not claim visual verification unless you inspected the images. If there are no defects, finish without edits. Screenshots: ${JSON.stringify(images)}`;
        await writeFile(join(directory, 'review.txt'), prompt);
        before = current;
        rounds = round;
        return { prompt, images };
      } finally {
        await browser.close();
      }
    },
  };
}
