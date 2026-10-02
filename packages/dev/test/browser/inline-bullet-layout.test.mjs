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
  setParagraphBullet,
  setParagraphLineSpacing,
  setShapeTextFormat,
} from '@office-kit/pptx';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { startPreview } from '../helpers/server.mjs';

// Measure every non-whitespace glyph, including runs with different font sizes.
// Per-character coordinates also catch changed wrapping without assuming that
// differently sized glyphs share the same top edge.
function paragraphGeometry(nodes) {
  return nodes.slice(-6).map((node) => {
    const rect = node.getBoundingClientRect();
    const glyphs = [];
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode;
      if (text.textContent === '•') continue;
      for (let offset = 0; offset < text.textContent.length; offset += 1) {
        if (/\s/.test(text.textContent[offset])) continue;
        const range = document.createRange();
        range.setStart(text, offset);
        range.setEnd(text, offset + 1);
        const box = range.getBoundingClientRect();
        glyphs.push({ x: box.x, y: box.y, width: box.width, height: box.height });
      }
    }
    return { top: rect.top, height: rect.height, glyphs };
  });
}

for (const width of [1100, 1500, 1900]) {
  test(
    `inline bullet editing keeps authored paragraph positions at viewport ${width}`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-inline-bullet-layout-'));
      let preview;
      let browser;
      try {
        const pres = createPresentation();
        const slide = addBlankSlide(pres);
        const text =
          'Large bullet\nPoint-size bullet\nRegular bullet\nA regular bullet paragraph with enough words to wrap onto a second line in the text box\nA fixed leading paragraph with enough words to wrap onto a second line in the text box\nTrailing regular bullet';
        const shape = addSlideTextBox(slide, {
          x: inches(1),
          y: inches(1),
          w: inches(6),
          h: inches(5),
          text,
        });
        for (let index = 0; index < 4; index += 1) setParagraphBullet(shape, index, 'bullet');
        const fixedStart = text.indexOf('A fixed leading');
        setShapeTextFormat(
          shape,
          { size: 24 },
          { range: { start: fixedStart, end: fixedStart + 'A fixed'.length } },
        );
        setParagraphLineSpacing(shape, 3, { kind: 'pct', value: 1.5 });
        setParagraphLineSpacing(shape, 4, { kind: 'pts', value: 24 });
        const parts = unzipSync(await savePresentation(pres));
        const name = 'ppt/slides/slide1.xml';
        const xml = strFromU8(parts[name]);
        let index = 0;
        parts[name] = strToU8(
          xml.replace(/<a:pPr([^>]*)><a:buChar char="•"\/><\/a:pPr>/g, (match, attrs) => {
            index += 1;
            if (index === 1)
              return `<a:pPr${attrs}><a:buSzPct val="400000"/><a:buChar char="•"/></a:pPr>`;
            if (index === 2)
              return `<a:pPr${attrs}><a:buSzPts val="2000"/><a:buChar char="•"/></a:pPr>`;
            return match;
          }),
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
        const page = await browser.newPage({ viewport: { width, height: 1000 } });
        await page.goto(preview.url);
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        await editor.getByText('Saved to this project', { exact: true }).waitFor();
        const before = await editor.locator('[data-pptx-paragraph]').evaluateAll(paragraphGeometry);
        await editor.locator('.hit').first().dblclick();
        const input = editor.locator('.inline-edit');
        await input.waitFor();
        const after = await input.locator('[data-text-paragraph]').evaluateAll(paragraphGeometry);
        if (after.length !== before.length)
          throw new Error(`paragraph count changed: before=${before.length} after=${after.length}`);
        for (let i = 0; i < before.length; i += 1) {
          const shift = Math.abs(after[i].top - before[i].top);
          if (shift > 1) throw new Error(`paragraph ${i + 1} shifted ${shift.toFixed(2)}px`);
          const heightShift = Math.abs(after[i].height - before[i].height);
          if (heightShift > 1)
            throw new Error(`paragraph ${i + 1} height shifted ${heightShift.toFixed(2)}px`);
          if (after[i].glyphs.length !== before[i].glyphs.length)
            throw new Error(
              `paragraph ${i + 1} contains ${after[i].glyphs.length} glyphs, expected ${before[i].glyphs.length}`,
            );
          for (let lineIndex = 0; lineIndex < before[i].glyphs.length; lineIndex += 1) {
            const expected = before[i].glyphs[lineIndex];
            const actual = after[i].glyphs[lineIndex];
            for (const key of ['x', 'y', 'width', 'height']) {
              const shift = Math.abs(actual[key] - expected[key]);
              if (shift > 1)
                throw new Error(
                  `paragraph ${i + 1} glyph ${lineIndex + 1} ${key} shifted ${shift.toFixed(2)}px`,
                );
            }
          }
        }
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );
}
