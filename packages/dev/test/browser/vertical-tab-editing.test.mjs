import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  setShapeTextDirection,
  setParagraphTabs,
  savePresentation,
  inches,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

for (const direction of ['vert', 'vert270']) {
  for (const alignment of ['left', 'center', 'right', 'decimal']) {
    test(
      `${direction} ${alignment} custom tab retains its position when entering the editor`,
      { timeout: 60000 },
      async () => {
        const dir = await mkdtemp(join(tmpdir(), 'office-vertical-tab-'));
        let preview, browser;
        try {
          const pres = createPresentation();
          const shape = addSlideTextBox(addBlankSlide(pres), {
            x: inches(1),
            y: inches(1),
            w: inches(3),
            h: inches(4),
            text: 'A\t12.34',
          });
          setShapeTextDirection(shape, direction);
          setParagraphTabs(shape, 0, { tabStops: [{ positionEmu: inches(2), alignment }] });
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
          await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
          const editor = page.frameLocator('#editor-frame');
          await editor.getByText('Saved to this project', { exact: true }).waitFor();
          const painted = await editor
            .locator('.paint [data-pptx-paragraph]')
            .evaluate((element) => {
              const spans = [...element.querySelectorAll('tspan')];
              const point = (text) => {
                const span = spans.find((span) => span.textContent === text);
                const position = span.getStartPositionOfChar(0);
                return new DOMPoint(position.x, position.y).matrixTransform(span.getScreenCTM());
              };
              return Math.abs(point('12.34').y - point('A').y);
            });
          await editor.locator('.hit').dblclick();
          const input = editor.locator('.inline-edit');
          await input.waitFor();
          const measureEditing = () =>
            input.evaluate(
              (element, { direction, alignment }) => {
                const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
                const positions = {};
                let shift = 0;
                let node;
                while ((node = walker.nextNode())) {
                  if (!['A', 'AB', '12.34'].includes(node.textContent)) continue;
                  const range = document.createRange();
                  range.setStart(node, 0);
                  range.setEnd(node, 1);
                  positions[node.textContent.startsWith('A') ? 'A' : node.textContent] =
                    range.getBoundingClientRect()[direction === 'vert270' ? 'bottom' : 'top'];
                  if (node.textContent === '12.34') {
                    range.setEnd(node, alignment === 'decimal' ? 2 : node.textContent.length);
                    shift =
                      alignment === 'left'
                        ? 0
                        : range.getBoundingClientRect().height / (alignment === 'center' ? 2 : 1);
                  }
                }
                return {
                  distance: Math.abs(positions['12.34'] - positions.A),
                  shift,
                  zoom: Number(getComputedStyle(element).getPropertyValue('--text-zoom')),
                };
              },
              { direction, alignment },
            );
          const editing = await measureEditing();
          assert.ok(
            Math.abs((editing.distance + editing.shift) / editing.zoom - 192) < 1,
            JSON.stringify(editing),
          );
          assert.ok(Math.abs(editing.distance - painted) < 2, JSON.stringify({ editing, painted }));
          await input.evaluate((element) => {
            const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
            let node;
            while ((node = walker.nextNode())) {
              if (node.textContent !== 'A') continue;
              const range = document.createRange();
              range.setStart(node, 1);
              range.collapse(true);
              const selection = window.getSelection();
              selection.removeAllRanges();
              selection.addRange(range);
              return;
            }
            throw new Error('Missing leading text');
          });
          await page.keyboard.insertText('B');
          await input.filter({ hasText: /^AB\t12\.34$/ }).waitFor();
          const reflowed = await measureEditing();
          assert.ok(
            Math.abs((reflowed.distance + reflowed.shift) / reflowed.zoom - 192) < 1,
            JSON.stringify(reflowed),
          );
          await input.press('ControlOrMeta+z');
          await input.filter({ hasText: /^A\t12\.34$/ }).waitFor();
          const undone = await measureEditing();
          assert.ok(Math.abs(undone.distance - painted) < 2, JSON.stringify(undone));
        } finally {
          await browser?.close();
          await preview?.close();
          await rm(dir, { recursive: true, force: true });
        }
      },
    );
  }
}
