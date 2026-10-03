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
  setShapeTextWrap,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

const text = 'A very long title that must stay on one line while editing';

async function openTextBox(t, wrap, value = text) {
  const dir = await mkdtemp(join(tmpdir(), 'office-text-wrap-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(1.35),
    h: inches(1.1),
    text: value,
  });
  setShapeTextWrap(shape, wrap);
  const source = join(dir, 'source.pptx');
  const file = join(dir, 'deck.tsx');
  await writeFile(source, await savePresentation(pres));
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
  );
  const preview = await startPreview(file);
  t.after(() => preview.close());
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
  await page.goto(preview.url);
  await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
  const editor = page.frameLocator('#editor-frame');
  await editor.getByText('Saved to this project', { exact: true }).waitFor();
  const editorFrame = page.frames().find((frame) => frame.url().endsWith('/editor'));
  assert.ok(editorFrame, 'editor frame did not load');
  const previewRects = await editorFrame.locator('.paint span').evaluateAll((nodes, needle) => {
    const candidates = nodes.filter((node) => node.textContent?.includes(needle));
    const slide = candidates.sort(
      (a, b) => b.getBoundingClientRect().width - a.getBoundingClientRect().width,
    )[0];
    if (!slide) return [];
    const range = document.createRange();
    range.selectNodeContents(slide);
    return [...range.getClientRects()].map((rect) => [rect.x, rect.y, rect.width, rect.height]);
  }, value);
  await editor.locator('.hit').first().dblclick();
  const input = editor.locator('.inline-edit');
  await input.waitFor();
  return {
    dir,
    preview,
    browser,
    page,
    input,
    previewRects,
  };
}

async function rangeRects(input) {
  return input.evaluate((node) => {
    const rects = [];
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode;
      if (!text.textContent) continue;
      const range = document.createRange();
      range.selectNodeContents(text);
      for (const rect of range.getClientRects())
        rects.push([rect.x, rect.y, rect.width, rect.height]);
    }
    return rects;
  });
}

function union(rects) {
  assert.ok(rects.length, 'text has no glyph rectangles');
  const x = Math.min(...rects.map((rect) => rect[0]));
  const y = Math.min(...rects.map((rect) => rect[1]));
  const right = Math.max(...rects.map((rect) => rect[0] + rect[2]));
  const bottom = Math.max(...rects.map((rect) => rect[1] + rect[3]));
  return { x, y, width: right - x, height: bottom - y };
}

function lineCount(rects) {
  return new Set(rects.filter((rect) => rect[2] > 0).map((rect) => Math.round(rect[1] * 10) / 10))
    .size;
}

test(
  'wrap=none keeps long text on one line while editing and typing',
  { timeout: 60000 },
  async (t) => {
    const state = await openTextBox(t, 'none');
    const initial = await rangeRects(state.input);
    assert.equal(lineCount(initial), 1, JSON.stringify(initial));
    const preview = union(state.previewRects);
    const editing = union(initial);
    assert.ok(Math.abs(editing.x - preview.x) < 2, JSON.stringify({ preview, editing }));
    assert.ok(Math.abs(editing.y - preview.y) < 2, JSON.stringify({ preview, editing }));
    assert.ok(Math.abs(editing.height - preview.height) < 2, JSON.stringify({ preview, editing }));
    await state.input.press('End');
    await state.input.type(' X');
    await state.page.waitForTimeout(100);
    const after = await rangeRects(state.input);
    assert.equal(lineCount(after), 1, JSON.stringify(after));
    const typed = union(after);
    assert.ok(Math.abs(typed.x - preview.x) < 2, JSON.stringify({ preview, typed }));
    assert.ok(Math.abs(typed.y - preview.y) < 2, JSON.stringify({ preview, typed }));
    assert.ok(Math.abs(typed.height - preview.height) < 2, JSON.stringify({ preview, typed }));
  },
);

test('wrap=none preserves explicit paragraph breaks', { timeout: 60000 }, async (t) => {
  const state = await openTextBox(t, 'none', 'First line\nSecond line');
  const initial = await rangeRects(state.input);
  assert.equal(lineCount(initial), 2, JSON.stringify(initial));
});

test(
  'square wrapping still produces multiple line fragments in the same narrow box',
  { timeout: 60000 },
  async (t) => {
    const state = await openTextBox(t, 'square');
    const initial = await rangeRects(state.input);
    assert.ok(initial.length > 1, JSON.stringify(initial));
  },
);
