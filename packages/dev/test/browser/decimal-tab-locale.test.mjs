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

test(
  'decimal tabs align on the run language separator in the preview and while editing',
  { timeout: 90000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-decimal-locale-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      for (const [index, lang] of ['de-DE', 'en-US'].entries()) {
        const shape = addSlideTextBox(slide, {
          x: inches(1),
          y: inches(1 + index * 2.5),
          w: inches(5),
          h: inches(2),
          text: 'A\t12,5\nB\t3,25',
        });
        setShapeTextLanguage(shape, lang);
        for (const paragraph of [0, 1])
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
      // Viewport x of each comma, and the field end (last digit's right edge).
      const painted = (shape) =>
        editor.locator('.paint').evaluate((paint, index) => {
          const group = paint.querySelectorAll('[data-pptx-shape-id]')[index];
          return [...group.querySelectorAll('tspan')]
            .filter((span) => span.textContent.includes(','))
            .map((span) => {
              const at = span.textContent.indexOf(',');
              const point = (position) =>
                new DOMPoint(position.x, position.y).matrixTransform(span.getScreenCTM());
              return {
                comma: point(span.getStartPositionOfChar(at)).x,
                end: point(span.getEndPositionOfChar(span.textContent.length - 1)).x,
              };
            });
        }, shape);
      const edited = (input) =>
        input.evaluate((node) => {
          const result = [];
          const walker = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT);
          for (let text = walker.nextNode(); text; text = walker.nextNode()) {
            const at = text.data.indexOf(',');
            if (at < 0) continue;
            const range = node.ownerDocument.createRange();
            range.setStart(text, at);
            range.setEnd(text, at + 1);
            const comma = range.getBoundingClientRect().left;
            range.setStart(text, text.data.length - 1);
            range.setEnd(text, text.data.length);
            result.push({ comma, end: range.getBoundingClientRect().right });
          }
          return result;
        });
      const german = await painted(0);
      const english = await painted(1);
      assert.equal(german.length, 2);
      assert.ok(Math.abs(german[0].comma - german[1].comma) < 0.5, JSON.stringify(german));
      assert.ok(Math.abs(english[0].end - english[1].end) < 0.5, JSON.stringify(english));
      assert.ok(Math.abs(english[0].comma - english[1].comma) > 3);

      for (const [index, expected] of [german, english].entries()) {
        await page.keyboard.press('Escape');
        await editor.locator('.hit').nth(index).dblclick();
        const input = editor.locator('.canvas-shell .inline-edit');
        await input.waitFor();
        const live = await edited(input);
        assert.equal(live.length, 2);
        for (const [line, value] of live.entries()) {
          // Same tolerance as the other editing-versus-painted tab checks.
          assert.ok(
            Math.abs(value.comma - expected[line].comma) < 1.5,
            `${index}: ${JSON.stringify(live)} vs ${JSON.stringify(expected)}`,
          );
        }
      }
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
