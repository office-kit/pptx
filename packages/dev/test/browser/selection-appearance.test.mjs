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
  addSlideLine,
  getShapeStrokeArrow,
  setShapeStrokeArrow,
  groupShapes,
  getGroupChildren,
  setShapeFill,
  setShapeStroke,
  setShapeStrokeDash,
  setShapeStrokeCap,
  setShapeStrokeJoin,
  setShapeStrokeCompound,
  getShapeStrokeCap,
  getShapeStrokeJoin,
  getShapeStrokeCompound,
  savePresentation,
  inches,
  getShapeFlip,
  getShapeFillColor,
  getShapeFill,
  getShapeStroke,
  getShapeStrokeWidth,
  getShapeStrokeDash,
  getShapeStrokeColor,
  getSlides,
  getSlideShapes,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test(
  'group child paint controls display saved values and track edits and undo',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-group-paint-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const first = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text: 'First',
      });
      const second = addSlideTextBox(slide, {
        x: inches(4),
        y: inches(1),
        w: inches(2),
        h: inches(1),
        text: 'Second',
      });
      setShapeFill(first, '123456');
      setShapeStroke(first, { color: 'ABCDEF', widthEmu: 25400 });
      setShapeStrokeDash(first, 'dash');
      setShapeStrokeCap(first, 'rnd');
      setShapeStrokeJoin(first, 'bevel');
      setShapeStrokeCompound(first, 'dbl');
      setShapeFill(second, '654321');
      groupShapes([first, second]);
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
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const colors = async () => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getGroupChildren(getSlideShapes(getSlides(deck)[0])[0]).map(getShapeFillColor);
      };
      await saved();
      await editor.locator('.hit').dblclick();
      await editor.locator('.hit').first().click();
      assert.equal(await editor.getByLabel('Fill', { exact: true }).inputValue(), '#123456');
      assert.equal(await editor.getByLabel('Outline', { exact: true }).inputValue(), '#abcdef');
      assert.equal(
        await editor.getByLabel('Outline width (points)', { exact: true }).inputValue(),
        '2',
      );
      assert.equal(
        await editor.getByRole('combobox', { name: /^Outline style/ }).inputValue(),
        'dash',
      );
      await editor.getByLabel('Fill', { exact: true }).evaluate((node) => {
        node.value = '#112233';
        node.dispatchEvent(new Event('change', { bubbles: true }));
      });
      await saved();
      assert.deepEqual(await colors(), ['#112233', '#654321']);
      assert.equal(await editor.getByLabel('Fill', { exact: true }).inputValue(), '#112233');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(), ['#123456', '#654321']);
      assert.equal(await editor.getByLabel('Fill', { exact: true }).inputValue(), '#123456');
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      assert.equal(await editor.locator('[data-paint-state=fill]').textContent(), 'Mixed');
      const lineProperties = async (read) => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getGroupChildren(getSlideShapes(getSlides(deck)[0])[0]).map(read);
      };
      for (const [label, read, previous, value] of [
        ['Compound type', getShapeStrokeCompound, 'dbl', 'tri'],
        ['Cap type', getShapeStrokeCap, 'rnd', 'sq'],
        ['Join type', getShapeStrokeJoin, 'bevel', 'miter'],
      ]) {
        const input = editor.getByRole('combobox', { name: label, exact: true });
        assert.equal(await input.inputValue(), 'mixed');
        await input.selectOption(value);
        await saved();
        assert.deepEqual(await lineProperties(read), [value, value]);
        assert.equal(await input.inputValue(), value);
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await saved();
        assert.deepEqual(await lineProperties(read), [previous, null]);
        assert.equal(await input.inputValue(), 'mixed');
      }
      const fillSection = editor
        .locator('.paint-section')
        .filter({ has: editor.locator('summary', { hasText: /^Fill$/ }) });
      await fillSection.locator('summary').click();
      assert.equal(await fillSection.getByLabel('Fill', { exact: true }).isVisible(), false);
      await fillSection.locator('summary').click();
      assert.equal(await fillSection.getByLabel('Fill', { exact: true }).isVisible(), true);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'fill and outline apply to all selected shapes in English and Japanese',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-selection-appearance-'));
    let preview, browser;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={2} height={1}>日本語</Text><Text x={4} y={1} width={2} height={1}>English</Text><Text x={1} y={3} width={2} height={1}>Untouched</Text></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      let ja = false;
      const saved = () =>
        editor
          .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
          .waitFor();
      const colors = async (reader) => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(deck)[0]).map(reader);
      };
      await saved();
      const initialFlips = await colors(getShapeFlip);
      await editor.locator('.hit').nth(0).click();
      await editor.getByRole('checkbox', { name: 'Flip horizontally', exact: true }).check();
      await saved();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      const mixedFlip = editor.getByRole('checkbox', {
        name: 'Flip horizontally (Mixed)',
        exact: true,
      });
      assert.equal(await mixedFlip.evaluate((node) => node.indeterminate), true);
      await mixedFlip.check();
      await saved();
      assert.deepEqual((await colors(getShapeFlip)).slice(0, 2), [
        { horizontal: true, vertical: false },
        { horizontal: true, vertical: false },
      ]);
      assert.deepEqual((await colors(getShapeFlip))[2], initialFlips[2]);
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.equal(await mixedFlip.evaluate((node) => node.indeterminate), true);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor.getByRole('checkbox', { name: '上下反転', exact: true }).check();
      await saved();
      assert.deepEqual((await colors(getShapeFlip)).slice(0, 2), [
        { horizontal: true, vertical: true },
        { horizontal: true, vertical: true },
      ]);
      await page.reload();
      await saved();
      await editor.locator('.hit').nth(0).click();
      assert.equal(
        await editor.getByRole('checkbox', { name: '左右反転', exact: true }).isChecked(),
        true,
      );
      assert.equal(
        await editor.getByRole('checkbox', { name: '上下反転', exact: true }).isChecked(),
        true,
      );
      assert.deepEqual((await colors(getShapeFlip))[2], initialFlips[2]);
      await editor.locator('.lang select').selectOption('en');
      ja = false;
      const initialFill = await colors(getShapeFillColor);
      const initialStroke = await colors(getShapeStrokeColor);
      await editor.locator('.hit').nth(0).click();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      assert.equal(await editor.locator('.hit.selected').count(), 2);
      const changeColor = async (index, color) => {
        await editor
          .locator('.bespoke input[type=color]')
          .nth(index)
          .evaluate((node, value) => {
            node.value = value;
            node.dispatchEvent(new Event('input', { bubbles: true }));
            node.dispatchEvent(new Event('change', { bubbles: true }));
          }, color);
        await saved();
      };
      await changeColor(0, '#123456');
      assert.deepEqual(await colors(getShapeFillColor), ['#123456', '#123456', initialFill[2]]);
      assert.equal(
        await editor.locator('.bespoke input[type=color]').nth(0).inputValue(),
        '#123456',
      );
      await editor.locator('.hit').nth(2).click();
      assert.notEqual(
        await editor.locator('.bespoke input[type=color]').nth(0).inputValue(),
        '#123456',
      );
      await editor
        .locator('.hit')
        .nth(0)
        .click({ modifiers: ['Shift'] });
      assert.equal(await editor.locator('[data-paint-state=fill]').textContent(), 'Mixed');
      await editor
        .locator('.hit')
        .nth(2)
        .click({ modifiers: ['Shift'] });
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeFillColor), initialFill);
      assert.notEqual(
        await editor.locator('.bespoke input[type=color]').nth(0).inputValue(),
        '#123456',
      );
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeFillColor), ['#123456', '#123456', initialFill[2]]);
      await editor.locator('.lang select').selectOption('ja');
      ja = true;
      await editor.locator('.hit').nth(2).click();
      await editor
        .locator('.hit')
        .nth(0)
        .click({ modifiers: ['Shift'] });
      assert.equal(await editor.locator('[data-paint-state=fill]').textContent(), '混在');
      await editor
        .locator('.hit')
        .nth(2)
        .click({ modifiers: ['Shift'] });
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      await changeColor(1, '#abcdef');
      assert.deepEqual(await colors(getShapeStrokeColor), ['#ABCDEF', '#ABCDEF', initialStroke[2]]);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeStrokeColor), initialStroke);
      await changeColor(1, '#abcdef');
      await page.reload();
      await saved();
      await editor.locator('.hit').nth(0).click();
      assert.equal(
        await editor.locator('.bespoke input[type=color]').nth(0).inputValue(),
        '#123456',
      );
      assert.equal(
        await editor.locator('.bespoke input[type=color]').nth(1).inputValue(),
        '#abcdef',
      );
      assert.deepEqual(await colors(getShapeFillColor), ['#123456', '#123456', initialFill[2]]);
      assert.deepEqual(await colors(getShapeStrokeColor), ['#ABCDEF', '#ABCDEF', initialStroke[2]]);
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      const originalWidths = await colors(getShapeStrokeWidth);
      const originalDashes = await colors(getShapeStrokeDash);
      const widthInput = editor.getByRole('spinbutton', {
        name: '枠線の太さ（ポイント）',
        exact: true,
      });
      await widthInput.fill('3.5');
      await widthInput.press('Tab');
      await saved();
      assert.deepEqual(await colors(getShapeStrokeWidth), [44450, 44450, originalWidths[2]]);
      assert.deepEqual(await colors(getShapeStrokeColor), ['#ABCDEF', '#ABCDEF', initialStroke[2]]);
      await widthInput.fill('-1');
      await widthInput.press('Tab');
      assert.equal(await widthInput.inputValue(), '3.5');
      await editor.getByRole('combobox', { name: '枠線の種類', exact: true }).selectOption('dash');
      await saved();
      assert.deepEqual(await colors(getShapeStrokeDash), ['dash', 'dash', originalDashes[2]]);
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeStrokeDash), originalDashes);
      await editor.getByTitle('やり直し (Ctrl+Y)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeStrokeDash), ['dash', 'dash', originalDashes[2]]);
      for (const dash of [
        'solid',
        'dot',
        'lgDash',
        'dashDot',
        'lgDashDot',
        'lgDashDotDot',
        'sysDash',
        'sysDot',
        'sysDashDot',
        'sysDashDotDot',
        'dash',
      ]) {
        await editor.getByRole('combobox', { name: '枠線の種類', exact: true }).selectOption(dash);
        await saved();
        assert.deepEqual(await colors(getShapeStrokeDash), [dash, dash, originalDashes[2]]);
      }
      await page.reload();
      await saved();
      await editor.locator('.hit').nth(0).click();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      assert.equal(
        await editor
          .getByRole('spinbutton', { name: '枠線の太さ（ポイント）', exact: true })
          .inputValue(),
        '3.5',
      );
      assert.equal(
        await editor.getByRole('combobox', { name: '枠線の種類', exact: true }).inputValue(),
        'dash',
      );
      const originalFillKinds = await colors((shape) => getShapeFill(shape).kind);
      const originalStrokeKinds = await colors((shape) => getShapeStroke(shape).kind);
      await editor
        .locator('.bespoke')
        .getByRole('button', { name: '塗りつぶしなし', exact: true })
        .click();
      await saved();
      assert.deepEqual(await colors((shape) => getShapeFill(shape).kind), [
        'none',
        'none',
        originalFillKinds[2],
      ]);
      assert.equal(await editor.locator('[data-paint-state=fill]').textContent(), 'なし');
      await editor.getByTitle('元に戻す (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeFillColor), ['#123456', '#123456', initialFill[2]]);
      await editor.locator('.lang select').selectOption('en');
      ja = false;
      await editor
        .locator('.bespoke')
        .getByRole('button', { name: 'No fill', exact: true })
        .click();
      await saved();
      await editor
        .locator('.bespoke')
        .getByRole('button', { name: 'No outline', exact: true })
        .click();
      await saved();
      assert.deepEqual(await colors((shape) => getShapeStroke(shape).kind), [
        'none',
        'none',
        originalStrokeKinds[2],
      ]);
      assert.equal(await editor.locator('[data-paint-state=stroke]').textContent(), 'None');
      await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
      await saved();
      assert.deepEqual(await colors(getShapeStrokeColor), ['#ABCDEF', '#ABCDEF', initialStroke[2]]);
      await editor.getByTitle('Redo (Ctrl+Y)', { exact: true }).click();
      await saved();
      await page.reload();
      await saved();
      assert.deepEqual(await colors((shape) => getShapeFill(shape).kind), [
        'none',
        'none',
        originalFillKinds[2],
      ]);
      assert.deepEqual(await colors((shape) => getShapeStroke(shape).kind), [
        'none',
        'none',
        originalStrokeKinds[2],
      ]);
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'arrow galleries preserve endpoint dimensions across mixed edits, undo and reload',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-arrow-gallery-'));
    let preview, browser;
    try {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const lines = [1, 5].map((x) =>
        addSlideLine(slide, {
          from: { x: inches(x), y: inches(1) },
          to: { x: inches(x + 2), y: inches(2) },
          widthEmu: 25400,
        }),
      );
      setShapeStrokeArrow(lines[0], 'head', { type: 'diamond', width: 'sm', length: 'lg' });
      setShapeStrokeArrow(lines[1], 'head', { type: 'oval', width: 'lg', length: 'sm' });
      setShapeStrokeArrow(lines[0], 'tail', { type: 'arrow', width: 'lg', length: 'lg' });
      addSlideTextBox(slide, {
        x: inches(1),
        y: inches(4),
        w: inches(2),
        h: inches(1),
        text: 'Text',
      });
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
      const errors = [];
      page.on('pageerror', (error) => errors.push(error.message));
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = () => editor.getByText('Saved to this project', { exact: true }).waitFor();
      const arrows = async (end) => {
        const deck = await loadPresentation(
          new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
        );
        return getSlideShapes(getSlides(deck)[0])
          .slice(0, 2)
          .map((shape) => getShapeStrokeArrow(shape, end));
      };
      const choose = async (label, option) => {
        await editor.getByRole('button', { name: label, exact: true }).click();
        await editor.getByRole('menuitemradio', { name: option, exact: true }).click();
        await saved();
      };
      await saved();
      await editor.locator('.hit').nth(0).click();
      await editor
        .locator('.hit')
        .nth(1)
        .click({ modifiers: ['Shift'] });
      const initialHead = await arrows('head');
      const initialTail = await arrows('tail');
      for (const [label, type] of [
        ['No Arrow', 'none'],
        ['Arrow', 'triangle'],
        ['Open Arrow', 'arrow'],
        ['Stealth Arrow', 'stealth'],
        ['Diamond Arrow', 'diamond'],
        ['Oval Arrow', 'oval'],
      ]) {
        await choose('Begin Arrow type', label);
        assert.deepEqual(
          await arrows('head'),
          initialHead.map((value) => ({ ...value, type })),
        );
        assert.deepEqual(await arrows('tail'), initialTail);
        await editor.getByTitle('Undo (Ctrl+Z)', { exact: true }).click();
        await saved();
        assert.deepEqual(await arrows('head'), initialHead);
      }
      const dimensions = ['sm', 'med', 'lg'];
      for (let index = 0; index < 9; index++) {
        await choose('Begin Arrow size', `Arrow Size ${index + 1}`);
        assert.deepEqual(
          await arrows('head'),
          initialHead.map((value) => ({
            ...value,
            width: dimensions[Math.floor(index / 3)],
            length: dimensions[index % 3],
          })),
        );
      }
      await choose('End Arrow type', 'Stealth Arrow');
      assert.deepEqual(await arrows('tail'), [
        { ...initialTail[0], type: 'stealth' },
        { type: 'stealth' },
      ]);
      await choose('End Arrow size', 'Arrow Size 2');
      assert.deepEqual(await arrows('tail'), [
        { type: 'stealth', width: 'sm', length: 'med' },
        { type: 'stealth', width: 'sm', length: 'med' },
      ]);
      await editor.getByRole('button', { name: 'End Arrow type', exact: true }).click();
      await editor
        .getByRole('menuitemradio', { name: 'Stealth Arrow', exact: true })
        .press('Escape');
      assert.equal(
        await editor.getByRole('menu', { name: 'End Arrow type', exact: true }).count(),
        0,
      );
      await page.reload();
      await saved();
      await editor.locator('.hit').nth(0).click();
      await editor.getByRole('button', { name: 'End Arrow size', exact: true }).click();
      assert.equal(
        await editor
          .getByRole('menuitemradio', { name: 'Arrow Size 2', exact: true })
          .getAttribute('aria-checked'),
        'true',
      );
      await editor
        .getByRole('menuitemradio', { name: 'Arrow Size 2', exact: true })
        .press('Escape');
      await editor.locator('.hit').nth(2).click();
      assert.equal(
        await editor.getByRole('button', { name: 'Begin Arrow type', exact: true }).isDisabled(),
        true,
      );
      assert.deepEqual(errors, []);
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
