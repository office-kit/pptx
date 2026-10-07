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
  getParagraphIndent,
  getParagraphPropertiesEffective,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphIndent,
  setParagraphTabs,
  setShapeParagraphs,
  setShapeRotation,
  setShapeTextDirection,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

async function open(build) {
  const dir = await mkdtemp(join(tmpdir(), 'office-ruler-transformed-'));
  const pres = createPresentation();
  const shape = addSlideTextBox(addBlankSlide(pres), {
    x: inches(2),
    y: inches(1.5),
    w: inches(4),
    h: inches(3),
    text: '',
  });
  setShapeParagraphs(shape, [
    { runs: [{ text: 'First paragraph' }] },
    { runs: [{ text: 'Second paragraph' }] },
  ]);
  build(shape);
  const source = join(dir, 'source.pptx');
  const file = join(dir, 'deck.tsx');
  await writeFile(source, await savePresentation(pres));
  await writeFile(
    file,
    `import {readFile} from 'node:fs/promises';import {Presentation} from '@office-kit/pptx-dsl';export default <Presentation source={await readFile(${JSON.stringify(source)})} />;`,
  );
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
  await installRichTextSelection(page);
  await page.goto(preview.url);
  const editor = page.frameLocator('#editor-frame');
  const saved = () =>
    editor.getByText(/^(Saved to this project|このプロジェクトに保存済み)$/).waitFor();
  await saved();
  const read = async () => {
    const loaded = await loadPresentation(
      new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
    );
    const target = getSlideShapes(getSlides(loaded)[0])[0];
    return [0, 1].map((index) => ({
      ...getParagraphIndent(target, index),
      tabs: getParagraphPropertiesEffective(loaded, target, index).tabStops ?? [],
    }));
  };
  const revision = async () => (await waitForState(preview.url, () => true)).revision;
  return {
    page,
    editor,
    saved,
    read,
    revision,
    preview,
    async edit(start, end) {
      await editor.getByRole('tab', { name: 'View', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: 'View', exact: true })
        .getByRole('checkbox', { name: 'Ruler', exact: true })
        .check();
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.canvas-shell .inline-edit');
      await input.evaluate(
        (node, [start, end]) => {
          window.selectEditorText(node, start, end);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        },
        [start, end],
      );
      return input;
    },
    async close() {
      await browser.close();
      await preview.close();
      await rm(dir, { recursive: true, force: true });
    },
  };
}

/** Viewport position of a paragraph's first character, inside the editor frame. */
const firstCharacter = (input, paragraph) =>
  input.evaluate((node, index) => {
    const section = node.querySelectorAll('[data-text-paragraph]')[index];
    const walker = node.ownerDocument.createTreeWalker(section, NodeFilter.SHOW_TEXT);
    const text = walker.nextNode();
    const range = node.ownerDocument.createRange();
    range.setStart(text, 0);
    range.setEnd(text, 1);
    const rect = range.getBoundingClientRect();
    return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
  }, paragraph);

/** Screen point of the text's indent zero, from the inline edit's real transform chain. */
const textOrigin = (input) =>
  input.evaluate((node) => {
    const rulers = node.ownerDocument.querySelector('.rulers').getBoundingClientRect();
    const style = getComputedStyle(node);
    const zero = node.ownerDocument.querySelector('.horizontal line[data-value="0"]');
    const verticalZero = node.ownerDocument.querySelector('.vertical line[data-value="0"]');
    return {
      vertical: style.writingMode.startsWith('vertical'),
      rulerX: rulers.left + Number(zero.getAttribute('x1')),
      rulerY: rulers.top + Number(verticalZero.getAttribute('y1')),
    };
  });

async function dragHandle(context, name, dx, dy, { hold } = {}) {
  const { page, editor } = context;
  const box = await editor.getByRole('button', { name, exact: true }).boundingBox();
  const start = { x: box.x + box.width / 2, y: box.y + box.height / 2 };
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(start.x + dx, start.y + dy, { steps: 6 });
  if (hold) await hold();
  await page.mouse.up();
}

test(
  'rotated text: ruler measures the unrotated text axis and reflows live while dragging',
  { timeout: 90000 },
  async () => {
    const context = await open((shape) => setShapeRotation(shape, 30));
    try {
      const { editor, read, revision, saved } = context;
      const input = await context.edit(17, 22);
      assert.equal(await editor.locator('.text-axis').getAttribute('data-axis'), 'x');
      const ruler = await editor.locator('.rulers svg.horizontal').boundingBox();
      const formatToggle = await editor
        .locator('.floating-text-format-bar > summary')
        .boundingBox();
      assert.ok(formatToggle.y >= ruler.y + ruler.height - 1, 'format toggle sits below the ruler');
      // The origin is the unrotated text-box start: the rotated box's centre
      // minus half its layout width plus its left inset.
      const geometry = await input.evaluate((node) => {
        const shell = node.closest('.inline-edit-shell').getBoundingClientRect();
        const matrix = new DOMMatrix(
          getComputedStyle(node.closest('.inline-edit-shell')).transform,
        );
        const pixel = Math.hypot(matrix.a, matrix.b);
        const style = getComputedStyle(node);
        return {
          expected:
            shell.left +
            shell.width / 2 -
            (node.offsetWidth / 2) * pixel +
            parseFloat(style.paddingLeft) * pixel,
        };
      });
      const origin = await textOrigin(input);
      assert.ok(
        Math.abs(origin.rulerX - geometry.expected) < 1.5,
        `${origin.rulerX} vs ${geometry.expected}`,
      );
      const firstLine = await editor
        .getByRole('button', { name: 'First line indent', exact: true })
        .boundingBox();
      assert.ok(Math.abs(firstLine.x + firstLine.width / 2 - origin.rulerX) < 1.5);

      // Live reflow: the second paragraph moves along the rotated text axis
      // before release, without saving; release commits one undo step.
      const before = await revision();
      const initial = await read();
      const startPoint = await firstCharacter(input, 1);
      let during;
      await dragHandle(context, 'First line indent', 40, 0, {
        hold: async () => {
          during = await firstCharacter(input, 1);
          assert.equal(await revision(), before);
        },
      });
      const angle = (30 * Math.PI) / 180;
      const moved =
        (during.x - startPoint.x) * Math.cos(angle) + (during.y - startPoint.y) * Math.sin(angle);
      assert.ok(Math.abs(moved - 40) < 4, `reflowed ${moved}px along the text axis`);
      await waitForState(context.preview.url, (state) => state.revision > before);
      await saved();
      const after = await read();
      assert.deepEqual(after[0], initial[0]);
      assert.ok(after[1].firstLineEmu > 0);
      const committed = await firstCharacter(input, 1);
      assert.ok(Math.hypot(committed.x - during.x, committed.y - during.y) < 2);
      assert.equal(
        await input.evaluate((node) => node.ownerDocument.getSelection().toString()),
        'econd',
      );

      // Escape restores the unchanged layout and saves nothing.
      const cancelRevision = await revision();
      const resting = await firstCharacter(input, 1);
      await dragHandle(context, 'Left indent', 30, 0, {
        hold: async () => {
          const mid = await firstCharacter(input, 1);
          assert.ok(Math.hypot(mid.x - resting.x, mid.y - resting.y) > 10);
          await input.press('Escape');
        },
      });
      assert.equal(await revision(), cancelRevision);
      const restored = await firstCharacter(input, 1);
      assert.ok(Math.hypot(restored.x - resting.x, restored.y - resting.y) < 1);

      const beforeUndo = await revision();
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await waitForState(context.preview.url, (state) => state.revision > beforeUndo);
      await saved();
      assert.deepEqual((await read())[1], initial[1]);
    } finally {
      await context.close();
    }
  },
);

for (const direction of ['vert', 'vert270']) {
  test(
    `${direction} text: indents and tabs measure along the vertical ruler`,
    { timeout: 90000 },
    async () => {
      const context = await open((shape) => setShapeTextDirection(shape, direction));
      try {
        const { editor, read, revision, saved } = context;
        const input = await context.edit(17, 22);
        assert.equal(await editor.locator('.text-axis').getAttribute('data-axis'), 'y');
        const sign = direction === 'vert' ? 1 : -1;
        const origin = await textOrigin(input);
        const handle = await editor
          .getByRole('button', { name: 'Left indent', exact: true })
          .boundingBox();
        assert.ok(Math.abs(handle.y + handle.height / 2 - origin.rulerY) < 1.5);
        // Indent zero is where the paragraph's first character starts.
        const start = await firstCharacter(input, 1);
        const glyph = await input.evaluate((node) => {
          const section = node.querySelectorAll('[data-text-paragraph]')[1];
          const range = node.ownerDocument.createRange();
          range.setStart(
            section.firstChild.nodeType === 3
              ? section.firstChild
              : section.querySelector('span').firstChild,
            0,
          );
          range.setEnd(range.startContainer, 1);
          const rect = range.getBoundingClientRect();
          return { top: rect.top, bottom: rect.bottom };
        });
        assert.ok(
          Math.abs((sign > 0 ? glyph.top : glyph.bottom) - origin.rulerY) < 3,
          `${JSON.stringify(glyph)} vs ${origin.rulerY}`,
        );

        const before = await revision();
        let during;
        await dragHandle(context, 'Left indent', 0, 30 * sign, {
          hold: async () => {
            during = await firstCharacter(input, 1);
            assert.equal(await revision(), before);
          },
        });
        assert.ok(Math.abs((during.y - start.y) * sign - 30) < 4, `${during.y - start.y}`);
        assert.ok(Math.abs(during.x - start.x) < 1);
        await waitForState(context.preview.url, (state) => state.revision > before);
        await saved();
        const indents = await read();
        assert.equal(indents[0].leftEmu, null);
        assert.ok(indents[1].leftEmu > 0);
        // The slide's painted (hidden) glyphs indent along the same axis.
        const painted = await editor.locator('.paint').evaluate((paint) => {
          const walker = paint.ownerDocument.createTreeWalker(paint, NodeFilter.SHOW_TEXT);
          for (let text = walker.nextNode(); text; text = walker.nextNode()) {
            if (!text.data.startsWith('Second')) continue;
            const range = paint.ownerDocument.createRange();
            range.setStart(text, 0);
            range.setEnd(text, 1);
            const rect = range.getBoundingClientRect();
            return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
          }
          throw new Error('painted paragraph not found');
        });
        const edited = await firstCharacter(input, 1);
        assert.ok(
          Math.hypot(painted.x - edited.x, painted.y - edited.y) < 2,
          `${JSON.stringify(painted)} vs ${JSON.stringify(edited)}`,
        );

        // A click on the vertical track adds a stop measured from the text start.
        const track = await editor
          .getByRole('button', { name: 'Add tab stop', exact: true })
          .boundingBox();
        const clickY = origin.rulerY + sign * 100;
        assert.ok(clickY > track.y && clickY < track.y + track.height);
        const beforeTab = await revision();
        await context.page.mouse.click(track.x + 5, clickY);
        await waitForState(context.preview.url, (state) => state.revision > beforeTab);
        await saved();
        const [stop] = (await read())[1].tabs;
        const pixelsPerEmu = await editor.locator('.rulers').evaluate(() => {
          const shell = document.querySelector('.inline-edit-shell');
          const matrix = new DOMMatrix(getComputedStyle(shell).transform);
          return Math.hypot(matrix.a, matrix.b) / 9525;
        });
        assert.ok(Math.abs(stop.positionEmu * pixelsPerEmu - 100) < 2, `${stop.positionEmu}`);
        const marker = await editor.locator('.tab-stop').boundingBox();
        assert.ok(Math.abs(marker.y + marker.height / 2 - clickY) < 1.5);

        // Keyboard steps follow the marker's screen direction.
        await editor.locator('.tab-stop').focus();
        const beforeKey = await revision();
        await editor.locator('.tab-stop').press('ArrowDown');
        await waitForState(context.preview.url, (state) => state.revision > beforeKey);
        await saved();
        assert.equal((await read())[1].tabs[0].positionEmu, stop.positionEmu + sign * 36000);
      } finally {
        await context.close();
      }
    },
  );
}

test(
  'mixed paragraphs show the last paragraph’s markers and move each relative to itself (Japanese labels)',
  { timeout: 90000 },
  async () => {
    const context = await open((shape) => {
      setParagraphIndent(shape, 0, { leftEmu: inches(0.5), firstLineEmu: -inches(0.25) });
      setParagraphIndent(shape, 1, { leftEmu: inches(1), firstLineEmu: 0 });
      setParagraphTabs(shape, 0, { tabStops: [{ positionEmu: inches(2), alignment: 'left' }] });
      setParagraphTabs(shape, 1, { tabStops: [{ positionEmu: inches(2.5), alignment: 'right' }] });
    });
    try {
      const { editor, read, revision, saved } = context;
      await editor.locator('.lang select').selectOption('ja');
      await editor.getByRole('tab', { name: '表示', exact: true }).click();
      await editor
        .getByRole('tabpanel', { name: '表示', exact: true })
        .getByRole('checkbox', { name: 'ルーラー', exact: true })
        .check();
      await editor.locator('.hit').first().dblclick();
      const input = editor.locator('.canvas-shell .inline-edit');
      const select = (start, end) =>
        input.evaluate(
          (node, [from, to]) => {
            window.selectEditorText(node, from, to);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          },
          [start, end],
        );
      const origin = await textOrigin(input);
      const pixelsPerEmu = await editor.locator('.rulers').evaluate(() => {
        const shell = document.querySelector('.inline-edit-shell');
        const matrix = new DOMMatrix(getComputedStyle(shell).transform);
        return Math.hypot(matrix.a, matrix.b) / 9525;
      });
      const center = async (name) => {
        const box = await editor.getByRole('button', { name, exact: true }).boundingBox();
        return box.x + box.width / 2;
      };
      const assertMarkers = async ({ left, first, tab, label }) => {
        await editor.getByRole('button', { name: label, exact: true }).waitFor();
        assert.ok(
          Math.abs((await center('左インデント')) - origin.rulerX - left * pixelsPerEmu) < 1.5,
        );
        assert.ok(
          Math.abs((await center('最初の行のインデント')) - origin.rulerX - first * pixelsPerEmu) <
            1.5,
        );
        assert.equal(await editor.locator('.tab-stop').count(), 1);
        const stop = await editor.locator('.tab-stop').boundingBox();
        assert.ok(Math.abs(stop.x + stop.width / 2 - origin.rulerX - tab * pixelsPerEmu) < 1.5);
      };
      // A selection inside one paragraph shows that paragraph's markers.
      await select(0, 3);
      await assertMarkers({
        left: inches(0.5),
        first: inches(0.25),
        tab: inches(2),
        label: 'タブ位置 5.08 cm',
      });
      // Across both paragraphs, Mac PowerPoint shows the last paragraph's.
      await select(3, 20);
      await assertMarkers({
        left: inches(1),
        first: inches(1),
        tab: inches(2.5),
        label: 'タブ位置 6.35 cm',
      });

      // Both paragraphs reflow during the drag.
      const initial = await read();
      const before = await revision();
      const starts = [await firstCharacter(input, 0), await firstCharacter(input, 1)];
      await dragHandle(context, '左インデント', 24, 0, {
        hold: async () => {
          const during = [await firstCharacter(input, 0), await firstCharacter(input, 1)];
          for (const index of [0, 1])
            assert.ok(Math.abs(during[index].x - starts[index].x - 24) < 3);
        },
      });
      await waitForState(context.preview.url, (state) => state.revision > before);
      await saved();
      const after = await read();
      const delta = after[0].leftEmu - initial[0].leftEmu;
      assert.ok(delta > 0);
      assert.equal(after[1].leftEmu - initial[1].leftEmu, delta);
      assert.equal(after[0].firstLineEmu, initial[0].firstLineEmu);

      // Moving the shown stop leaves the other paragraph's stop alone.
      const beforeTab = await revision();
      const box = await editor.locator('.tab-stop').boundingBox();
      await context.page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await context.page.mouse.down();
      await context.page.mouse.move(box.x + box.width / 2 + 30, box.y + box.height / 2, {
        steps: 5,
      });
      await context.page.mouse.up();
      await waitForState(context.preview.url, (state) => state.revision > beforeTab);
      await saved();
      const tabs = await read();
      assert.ok(tabs[1].tabs[0].positionEmu > inches(2.5));
      assert.equal(tabs[1].tabs[0].alignment, 'right');
      assert.deepEqual(tabs[0].tabs, initial[0].tabs);
    } finally {
      await context.close();
    }
  },
);

test('dragging a tab stop reflows the tabbed text before release', { timeout: 90000 }, async () => {
  const context = await open((shape) => {
    setShapeParagraphs(shape, [{ runs: [{ text: 'A\tB' }] }, { runs: [{ text: 'Plain' }] }]);
    setParagraphTabs(shape, 0, { tabStops: [{ positionEmu: inches(1.5), alignment: 'left' }] });
  });
  try {
    const { editor, page, read, revision, saved } = context;
    const input = await context.edit(1, 1);
    const glyphB = () =>
      input.evaluate((node) => {
        const walker = node.ownerDocument.createTreeWalker(node, NodeFilter.SHOW_TEXT);
        for (let text = walker.nextNode(); text; text = walker.nextNode()) {
          const index = text.data.indexOf('B');
          if (index < 0) continue;
          const range = node.ownerDocument.createRange();
          range.setStart(text, index);
          range.setEnd(text, index + 1);
          return range.getBoundingClientRect().left;
        }
        throw new Error('B not found');
      });
    const start = await glyphB();
    const before = await revision();
    const box = await editor.locator('.tab-stop').boundingBox();
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2 + 45, box.y + box.height / 2, { steps: 6 });
    assert.ok(Math.abs((await glyphB()) - start - 45) < 3);
    assert.equal(await revision(), before);
    // Dragging off the ruler previews removal: B falls back to the default grid.
    await page.mouse.move(box.x + box.width / 2 + 45, box.y + 80, { steps: 3 });
    assert.ok(Math.abs((await glyphB()) - start) > 3);
    await input.press('Escape');
    await page.mouse.up();
    assert.ok(Math.abs((await glyphB()) - start) < 1);
    assert.equal(await revision(), before);
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2 + 45, box.y + box.height / 2, { steps: 6 });
    await page.mouse.up();
    await waitForState(context.preview.url, (state) => state.revision > before);
    await saved();
    assert.ok((await read())[0].tabs[0].positionEmu > inches(1.5));
    assert.ok(Math.abs((await glyphB()) - start - 45) < 3);
  } finally {
    await context.close();
  }
});
