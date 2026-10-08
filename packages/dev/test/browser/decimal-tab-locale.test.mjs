import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  inches,
  savePresentation,
  setParagraphTabs,
  setShapeTextLanguage,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

// The lines of the reference desktop app's (Mac) capture (2026-10-07): one decimal tab, and
// numbers with ",", ".", neither, and both. Set Proofing Language to German
// moved the alignment from "." to ","; a number without the run language's
// separator ends at the tab.
const LINES = ['A\t12,50', 'B\t3.25', 'C\t1234', 'D\t1.234,5'];
const ALIGNED = {
  'de-DE': { separator: ',', aligned: [0, 3], ending: [1, 2] },
  'en-US': { separator: '.', aligned: [1, 3], ending: [0, 2] },
};

test(
  'decimal tabs align on the run language separator in the preview and while editing',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-decimal-locale-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const languages = Object.keys(ALIGNED);
      for (const [index, lang] of languages.entries()) {
        const shape = addSlideTextBox(slide, {
          x: inches(1),
          y: inches(0.5 + index * 3.25),
          w: inches(5),
          h: inches(3),
          text: LINES.join('\n'),
        });
        setShapeTextLanguage(shape, lang);
        for (const paragraph of LINES.keys())
          setParagraphTabs(shape, paragraph, {
            tabStops: [{ positionEmu: inches(2), alignment: 'decimal' }],
          });
      }
      const source = join(dir, 'source.pptx');
      const file = join(dir, 'deck.tsx');
      await writeFile(source, await savePresentation(pres));
      await writeFile(
        file,
        `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await page.goto(preview.url);
      const editor = page.frameLocator('#editor-frame');
      await editor.getByText('Saved to this project', { exact: true }).waitFor();
      // Per number, in line order: the viewport x of its separator (null when
      // it has none) and of its last digit's right edge.
      const painted = (shape, separator) =>
        editor.locator('.paint').evaluate(
          (paint, [index, mark]) => {
            const group = paint.querySelectorAll('[data-pptx-shape-id]')[index];
            return [...group.querySelectorAll('tspan')]
              .filter((span) => /^[\d.,]+$/.test(span.textContent))
              .map((span) => {
                const at = span.textContent.indexOf(mark);
                const point = (position) =>
                  new DOMPoint(position.x, position.y).matrixTransform(span.getScreenCTM()).x;
                return {
                  separator: at < 0 ? null : point(span.getStartPositionOfChar(at)),
                  end: point(span.getEndPositionOfChar(span.textContent.length - 1)),
                };
              });
          },
          [shape, separator],
        );
      const edited = (input, separator) =>
        input.evaluate((node, mark) => {
          const result = [];
          const walker = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT);
          for (let text = walker.nextNode(); text; text = walker.nextNode()) {
            const number = /[\d.,]+$/.exec(text.data);
            if (!number || !/\d/.test(number[0])) continue;
            const range = node.ownerDocument.createRange();
            const at = text.data.indexOf(mark, number.index);
            let position = null;
            if (at >= 0) {
              range.setStart(text, at);
              range.setEnd(text, at + 1);
              position = range.getBoundingClientRect().left;
            }
            range.setStart(text, text.data.length - 1);
            range.setEnd(text, text.data.length);
            result.push({ separator: position, end: range.getBoundingClientRect().right });
          }
          return result;
        }, separator);
      const near = (a, b, tolerance) => Math.abs(a - b) < tolerance;
      for (const [index, lang] of languages.entries()) {
        const { separator, aligned, ending } = ALIGNED[lang];
        const numbers = await painted(index, separator);
        const detail = `${lang}: ${JSON.stringify(numbers)}`;
        assert.equal(numbers.length, LINES.length, detail);
        const [first, second] = aligned.map((line) => numbers[line].separator);
        assert.ok(near(first, second, 0.5), detail);
        // Numbers without the separator end where it would have been.
        for (const line of ending) assert.ok(near(numbers[line].end, first, 0.5), detail);

        await page.keyboard.press('Escape');
        await editor.locator('.hit').nth(index).dblclick();
        const input = editor.locator('.canvas-shell .inline-edit');
        await input.waitFor();
        const live = await edited(input, separator);
        assert.equal(live.length, LINES.length, JSON.stringify(live));
        for (const line of aligned)
          // Same tolerance as the other editing-versus-painted tab checks.
          assert.ok(
            near(live[line].separator, numbers[line].separator, 1.5),
            `${lang}: ${JSON.stringify(live)} vs ${detail}`,
          );
      }
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
