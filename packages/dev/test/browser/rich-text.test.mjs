import { chromium } from 'playwright';
import { build } from 'esbuild';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';

test('rich editing preserves autofit font and line scales, text and selection', async () => {
  const bundle = await build({
    entryPoints: [fileURLToPath(new URL('../../src/rich-text.ts', import.meta.url))],
    bundle: true,
    write: false,
    format: 'iife',
    globalName: 'richText',
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.setContent('<body></body>');
    await page.addScriptTag({ content: bundle.outputFiles[0].text });
    const result = await page.evaluate(() => {
      const field = richText.createRichTextField();
      document.body.append(field);
      field.value = 'Large\nSmall';
      field.focus();
      field.setSelectionRange(2, 8);
      const runs = [
        { start: 0, end: 5, format: { size: 40 } },
        { start: 6, end: 11, format: { size: 20 } },
      ];
      const paragraphs = [
        { start: 0, end: 5, properties: { level: 0, lineSpacing: { kind: 'pct', value: 1.2 } } },
        { start: 6, end: 11, properties: { level: 0, lineSpacing: { kind: 'pts', value: 30 } } },
      ];
      const snapshot = () => ({
        text: field.value,
        selection: [field.selectionStart, field.selectionEnd],
        sizes: [...field.querySelectorAll('span')].map((s) =>
          parseFloat(getComputedStyle(s).fontSize),
        ),
        lines: [...field.children].map((p) => parseFloat(getComputedStyle(p).lineHeight)),
      });
      field.renderRuns(runs, 1 / 12700, paragraphs);
      const before = snapshot();
      field.renderRuns(runs, 1 / 12700, paragraphs, { fontScale: 0.65, lnSpcReduction: 0.1 });
      const scaled = snapshot();
      field.renderRuns(runs, 1 / 12700, paragraphs, null);
      const restored = snapshot();
      Object.assign(field.style, {
        width: '140px',
        height: '60px',
        boxSizing: 'border-box',
        padding: '4px',
        whiteSpace: 'pre-wrap',
        overflow: 'auto',
        alignContent: 'center',
      });
      const fitted = richText.fitRichText(field, runs, 1 / 12700, paragraphs, 40);
      const fittedSnapshot = snapshot();
      const overflows =
        field.scrollHeight > field.clientHeight || field.scrollWidth > field.clientWidth;
      field.style.height = '200px';
      field.style.width = '400px';
      const expanded = richText.fitRichText(field, runs, 1 / 12700, paragraphs, 40);
      field.style.width = '140px';
      field.style.height = '10px';
      const natural = richText.measureRichTextBox(field);
      field.style.writingMode = 'vertical-rl';
      field.style.height = '140px';
      field.style.width = '10px';
      const verticalNatural = richText.measureRichTextBox(field);
      const selectionAfterMeasure = [field.selectionStart, field.selectionEnd];
      const inactive = richText.createRichTextField();
      inactive.value = 'Large\nSmall';
      document.body.append(inactive);
      inactive.renderRuns(runs, 1 / 12700, paragraphs);
      const selectionAfterInactiveRender = [field.selectionStart, field.selectionEnd];
      const selectionStillInField = field.contains(getSelection().anchorNode);
      inactive.dispose();
      inactive.remove();
      field.dispose();
      return {
        before,
        scaled,
        restored,
        fitted,
        fittedSnapshot,
        expanded,
        overflows,
        anchor: field.style.alignContent,
        natural,
        verticalNatural,
        selectionAfterMeasure,
        selectionAfterInactiveRender,
        selectionStillInField,
      };
    });
    assert.deepEqual(result.scaled.sizes, [26, 13]);
    assert.deepEqual(result.scaled.selection, [2, 8]);
    assert.equal(result.scaled.text, 'Large\nSmall');
    for (let i = 0; i < 2; i++)
      assert.ok(Math.abs(result.scaled.lines[i] / result.before.lines[i] - 0.585) < 0.001);
    assert.deepEqual(result.restored, result.before);
    assert.ok(result.fitted.fontScale < 1 && result.fitted.fontScale >= 0.25);
    assert.equal(result.overflows, false);
    assert.deepEqual(result.fittedSnapshot.selection, [2, 8]);
    assert.equal(result.fittedSnapshot.text, 'Large\nSmall');
    assert.equal(result.expanded.fontScale, 1);
    assert.equal(result.anchor, 'center');
    assert.equal(result.natural.width, 140);
    assert.ok(result.natural.height > 60);
    assert.equal(result.verticalNatural.height, 140);
    assert.ok(result.verticalNatural.width > 10);
    assert.deepEqual(result.selectionAfterMeasure, [2, 8]);
    assert.deepEqual(result.selectionAfterInactiveRender, [2, 8]);
    assert.equal(result.selectionStillInField, true);
  } finally {
    await browser.close();
  }
});

test('centered text bounds preserve paragraph alignment and editing selection', async () => {
  const bundle = await build({
    entryPoints: [fileURLToPath(new URL('../../src/rich-text.ts', import.meta.url))],
    bundle: true,
    write: false,
    format: 'iife',
    globalName: 'richText',
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.setContent('<body></body>');
    await page.addScriptTag({ content: bundle.outputFiles[0].text });
    const result = await page.evaluate(() => {
      const field = richText.createRichTextField();
      document.body.append(field);
      Object.assign(field.style, {
        width: '400px',
        height: '200px',
        padding: '10px',
        boxSizing: 'border-box',
        whiteSpace: 'pre-wrap',
      });
      field.value = 'Long text here\nShort';
      field.focus();
      field.setSelectionRange(2, 5);
      field.renderRuns([{ start: 0, end: field.value.length, format: { size: 20 } }], 1 / 12700, [
        { start: 0, end: 14, properties: { level: 0, align: 'left' } },
        { start: 15, end: 20, properties: { level: 0, align: 'left' } },
      ]);
      const xs = () =>
        [...field.querySelectorAll('span')].map((span) => span.getBoundingClientRect().left);
      const before = xs();
      richText.centerRichText(field, true);
      const after = xs();
      const selection = [field.selectionStart, field.selectionEnd];
      richText.centerRichText(field, false);
      const restored = xs();
      field.dispose();
      return { before, after, restored, selection };
    });
    assert.ok(result.after[0] > result.before[0] + 50);
    assert.ok(Math.abs(result.after[0] - result.after[1]) < 0.1);
    assert.deepEqual(result.selection, [2, 5]);
    assert.deepEqual(result.restored, result.before);
  } finally {
    await browser.close();
  }
});

test('menu deletion uses saved selection across formatted paragraphs', async () => {
  const bundle = await build({
    entryPoints: [fileURLToPath(new URL('../../src/rich-text.ts', import.meta.url))],
    bundle: true,
    write: false,
    format: 'iife',
    globalName: 'richText',
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.setContent('<button id="menu">Edit</button>');
    await page.addScriptTag({ content: bundle.outputFiles[0].text });
    const result = await page.evaluate(() => {
      const field = richText.createRichTextField();
      document.body.append(field);
      field.value = 'Alpha\nBeta';
      field.renderRuns(
        [
          { start: 0, end: 5, format: { bold: true } },
          { start: 6, end: 10, format: { italic: true } },
        ],
        1 / 12700,
        [
          { start: 0, end: 5, properties: { level: 0 } },
          { start: 6, end: 10, properties: { level: 0 } },
        ],
      );
      field.focus();
      field.setSelectionRange(2, 8);
      // Save the selection before the application menu takes focus.
      const selected = [field.selectionStart, field.selectionEnd];
      document.getElementById('menu').focus();
      const events = [];
      field.addEventListener('beforeinput', (event) =>
        events.push({
          type: event.inputType,
          value: field.value,
          range: [field.selectionStart, field.selectionEnd],
        }),
      );
      field.addEventListener('input', (event) =>
        events.push({
          type: event.inputType,
          value: field.value,
          range: [field.selectionStart, field.selectionEnd],
        }),
      );
      field.deleteSelection();
      field.deleteSelection(); // A caret-only deletion is a no-op.
      const result = {
        selected,
        value: field.value,
        events,
        focused: document.activeElement === field,
      };
      field.dispose();
      return result;
    });
    assert.deepEqual(result.selected, [2, 8]);
    assert.equal(result.value, 'Alta');
    assert.equal(result.focused, true);
    assert.deepEqual(result.events, [
      { type: 'deleteContentBackward', value: 'Alpha\nBeta', range: [2, 8] },
      { type: 'deleteContentBackward', value: 'Alta', range: [2, 2] },
    ]);
  } finally {
    await browser.close();
  }
});
