import { installRichTextSelection } from '../helpers/rich-text.mjs';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { build, transform } from 'esbuild';
import { chromium } from 'playwright';
import {
  getSlides,
  getSlideShapes,
  getShapeParagraphElements,
  getTableCells,
  getTableCellParagraphs,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview } from '../helpers/server.mjs';

test('HTML clipboard parsing preserves inline formats without executing markup or fetching resources', async () => {
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    const requests = [];
    page.on('request', (request) => requests.push(request.url()));
    const module = await transform(
      await readFile(
        new URL('../../../../site/src/lib/editor/core/html-text-clipboard.ts', import.meta.url),
        'utf8',
      ),
      { loader: 'ts', format: 'esm' },
    );
    const measurer = await build({
      stdin: {
        contents: await readFile(
          new URL('../../../../packages/preview/src/browser-measure.ts', import.meta.url),
          'utf8',
        ),
        loader: 'ts',
        resolveDir: fileURLToPath(new URL('../../../../packages/preview/src/', import.meta.url)),
      },
      bundle: true,
      format: 'esm',
      platform: 'browser',
      write: false,
    });
    const measurerCode = measurer.outputFiles[0]?.text;
    if (!measurerCode) throw new Error('browser text measurer bundle is empty');
    const result = await page.evaluate(
      async ({ clipboardSource, measurerSource }) => {
        const importSource = async (source) => {
          const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
          try {
            return await import(url);
          } finally {
            URL.revokeObjectURL(url);
          }
        };
        const { parseHtmlTextClipboard: parse, textClipboardHtml: serialize } =
          await importSource(clipboardSource);
        const { browserTextMeasurer } = await importSource(measurerSource);
        const measureText = browserTextMeasurer();
        if (!measureText) throw new Error('browser text measurer unavailable');
        const measureSpec = {
          family: 'Arial',
          sizePx: 48,
          bold: false,
          italic: false,
          letterSpacingPx: 0,
        };
        const browserMeasureKerning = [true, false, true].map(
          (kerning) => measureText('AV', { ...measureSpec, kerning }).widthPx,
        );
        const mixed = parse(
          '<b style="font-weight:normal"><span style="font-weight:700;font-size:24px;color:rgb(255,0,0);font-family:Arial">English</span><i style="background-color:yellow;text-decoration:underline line-through">日本語</i></b>',
          'English日本語',
        );
        const multiline = parse(
          '<p><strong>A</strong></p><p></p><p><sub>日本語</sub><br>End</p>',
          'A\n\n日本語\nEnd',
        );
        const spacing = parse('<span style="letter-spacing:4px">Wide</span>', 'Wide');
        const negativeSpacing = parse('<span style="letter-spacing:-2px">Tight</span>', 'Tight');
        const zeroSpacing = parse('<span style="letter-spacing:0px">Default</span>', 'Default');
        const kerningOffHtml = serialize({
          text: 'AV',
          formats: [{ start: 0, end: 2, format: { size: 10, kern: 1200 } }],
        });
        const kerningOnHtml = serialize({
          text: 'AV',
          formats: [{ start: 0, end: 2, format: { size: 24, kern: 1200 } }],
        });
        const resetSpacing = parse(
          '<span style="letter-spacing:4px">Wide<span style="letter-spacing:normal">Reset</span></span>',
          'WideReset',
        );
        const allCaps = parse(
          '<span style="text-transform:uppercase">Mixed Case</span>',
          'Mixed Case',
        );
        const allCapsWithNormalVariant = parse(
          '<span style="text-transform:uppercase;font-variant-caps:normal">Mixed Case</span>',
          'Mixed Case',
        );
        const smallCaps = parse(
          '<span style="font-variant-caps:small-caps">Mixed Case</span>',
          'Mixed Case',
        );
        const smallCapsWithNoneTransform = parse(
          '<span style="text-transform:none;font-variant-caps:small-caps">Mixed Case</span>',
          'Mixed Case',
        );
        const resetCaps = parse(
          '<span style="text-transform:uppercase">LOUD<span style="text-transform:none">Quiet</span></span>',
          'LOUDQuiet',
        );
        const inheritedAllCaps = parse(
          '<span style="text-transform:uppercase">LOUD<span style="font-variant-caps:normal">Still Loud</span></span>',
          'LOUDStill Loud',
        );
        const inheritedSmallCaps = parse(
          '<span style="font-variant-caps:small-caps">Small<span style="text-transform:none">Still Small</span></span>',
          'SmallStill Small',
        );
        const cssUnderlineStyles = ['double', 'dotted', 'dashed', 'wavy'];
        const parsedUnderlineStyles = Object.fromEntries(
          cssUnderlineStyles.map((style) => {
            const parsed = parse(
              `<span style="text-decoration-line:underline;text-decoration-style:${style}">A</span>`,
              'A',
            );
            return [style, parsed?.formats[0]?.format.underline];
          }),
        );
        const serializedUnderlineStyles = Object.fromEntries(
          cssUnderlineStyles.map((style) => {
            const value = { double: 'dbl', dotted: 'dotted', dashed: 'dash', wavy: 'wavy' }[style];
            return [
              style,
              serialize({
                text: 'A',
                formats: [{ start: 0, end: 1, format: { underline: value } }],
              }),
            ];
          }),
        );
        const combinedPatternedStrike = serialize({
          text: 'A',
          formats: [{ start: 0, end: 1, format: { underline: 'wavy', strike: true } }],
        });
        const combinedPatternedStrikeRoundtrip = parse(combinedPatternedStrike, 'A');
        const underlineNone = parse('<u style="text-decoration:none">A</u>', 'A');
        const strikeNone = parse('<s style="text-decoration:none">A</s>', 'A');
        const inheritedDecorationNone = parse(
          '<span style="text-decoration:underline">Parent<u style="text-decoration:none">Child</u></span>',
          'ParentChild',
        );
        const decorationNoneMount = document.createElement('div');
        decorationNoneMount.innerHTML =
          '<u style="text-decoration:none">A</u><s style="text-decoration:none">B</s>';
        document.body.append(decorationNoneMount);
        const computedDecorationNone = [...decorationNoneMount.children].map(
          (node) => getComputedStyle(node).textDecorationLine,
        );
        decorationNoneMount.remove();
        const hostile = parse(
          '<script>globalThis.clipboardExecuted=true</script><img src="https://clipboard.invalid/image" onerror="globalThis.clipboardExecuted=true"><iframe src="https://clipboard.invalid/frame"></iframe><style>@import "https://clipboard.invalid/style";</style><b>Safe</b>',
          'Safe',
        );
        const hostileFont = {
          text: '<script>&日本語',
          formats: [
            {
              start: 0,
              end: 12,
              format: { font: '";background:url(https://clipboard.invalid/font)', bold: true },
            },
          ],
        };
        const exported = serialize(hostileFont);
        const markup = document.createElement('template');
        markup.innerHTML = exported;
        const kerningMount = document.createElement('div');
        kerningMount.innerHTML = `${kerningOnHtml}${kerningOffHtml}${kerningOnHtml}`;
        document.body.append(kerningMount);
        const kerningSpans = [...kerningMount.querySelectorAll('span')];
        return {
          inherited: parse(
            '<b><i><span style="font-weight:inherit;font-style:inherit">A</span></i></b>',
            'A',
          ),
          mixed,
          multiline,
          spacing,
          negativeSpacing,
          zeroSpacing,
          resetSpacing,
          allCaps,
          allCapsWithNormalVariant,
          smallCaps,
          smallCapsWithNoneTransform,
          resetCaps,
          inheritedAllCaps,
          inheritedSmallCaps,
          parsedUnderlineStyles,
          serializedUnderlineStyles,
          combinedPatternedStrike,
          combinedPatternedStrikeRoundtrip,
          underlineNone,
          strikeNone,
          inheritedDecorationNone,
          computedDecorationNone,
          spacingRoundtrip: parse(serialize(spacing), spacing.text),
          allCapsRoundtrip: parse(serialize(allCaps), allCaps.text),
          smallCapsRoundtrip: parse(serialize(smallCaps), smallCaps.text),
          zeroSpacingRoundtrip: parse(serialize(zeroSpacing), zeroSpacing.text),
          kerningOffHtml,
          kerningOnHtml,
          inlineKerning: kerningSpans.map((span) => getComputedStyle(span).fontKerning),
          browserMeasureKerning,
          hostile,
          executed: !!globalThis.clipboardExecuted,
          mismatch: parse('<b>wrong</b>', 'right'),
          table: parse('<table><tr><td>A</td></tr></table>', 'A'),
          deep: parse('<span>'.repeat(150) + 'A' + '</span>'.repeat(150), 'A'),
          pre: parse('<div style="white-space:pre-wrap"><b>A  B\n日本語</b></div>', 'A  B\n日本語'),
          roundtrip: parse(serialize(mixed), mixed.text),
          exportedText: markup.content.textContent,
          exportedScript: !!markup.content.querySelector('script'),
        };
      },
      { clipboardSource: module.code, measurerSource: measurerCode },
    );
    assert.equal(result.inherited.formats[0].format.bold, true);
    assert.equal(result.inherited.formats[0].format.italic, true);
    assert.equal(result.mixed.formats[0].format.bold, true);
    assert.equal(result.mixed.formats[0].format.size, 18);
    assert.equal(result.mixed.formats[0].format.color, '#ff0000');
    assert.equal(result.mixed.formats[0].format.fontEastAsian, 'Arial');
    assert.equal(result.mixed.formats[1].format.bold, false);
    assert.equal(result.mixed.formats[1].format.italic, true);
    assert.equal(result.mixed.formats[1].format.underline, true);
    assert.equal(result.mixed.formats[1].format.strike, true);
    assert.equal(result.mixed.formats[1].format.highlight, '#ffff00');
    assert.equal(result.multiline.text, 'A\n\n日本語\nEnd');
    assert.equal(result.multiline.formats.find((s) => s.format.baseline)?.format.baseline, -0.25);
    assert.equal(result.spacing.formats[0].format.spc, 300);
    assert.equal(result.negativeSpacing.formats[0].format.spc, -150);
    assert.equal(result.zeroSpacing.formats[0].format.spc, 0);
    assert.match(result.kerningOffHtml, /font-kerning:\s*none/);
    assert.match(result.kerningOnHtml, /font-kerning:\s*normal/);
    assert.deepEqual(result.inlineKerning, ['normal', 'none', 'normal']);
    assert.ok(result.browserMeasureKerning[0] < result.browserMeasureKerning[1]);
    assert.ok(Math.abs(result.browserMeasureKerning[0] - result.browserMeasureKerning[2]) < 1e-9);
    assert.equal(result.resetSpacing.formats[1].format.spc, 0);
    assert.equal(result.allCaps.formats[0].format.cap, 'all');
    assert.equal(result.allCapsWithNormalVariant.formats[0].format.cap, 'all');
    assert.equal(result.smallCaps.formats[0].format.cap, 'small');
    assert.equal(result.smallCapsWithNoneTransform.formats[0].format.cap, 'small');
    assert.equal(result.resetCaps.formats[1].format.cap, 'none');
    assert.equal(result.inheritedAllCaps.formats[1].format.cap, 'all');
    assert.equal(result.inheritedSmallCaps.formats[1].format.cap, 'small');
    assert.deepEqual(result.parsedUnderlineStyles, {
      double: 'dbl',
      dotted: 'dotted',
      dashed: 'dash',
      wavy: 'wavy',
    });
    for (const style of ['double', 'dotted', 'dashed', 'wavy'])
      assert.match(
        result.serializedUnderlineStyles[style],
        new RegExp(`text-decoration-style: ${style}`),
      );
    assert.match(result.combinedPatternedStrike, /text-decoration-line: line-through/);
    assert.match(result.combinedPatternedStrike, /<u[^>]*text-decoration-line: underline/);
    assert.equal(result.combinedPatternedStrikeRoundtrip.formats[0].format.underline, 'wavy');
    assert.equal(result.combinedPatternedStrikeRoundtrip.formats[0].format.strike, true);
    assert.equal(result.underlineNone.formats[0].format.underline, undefined);
    assert.equal(result.strikeNone.formats[0].format.strike, undefined);
    assert.equal(result.inheritedDecorationNone.formats[1].format.underline, true);
    assert.deepEqual(result.computedDecorationNone, ['none', 'none']);
    assert.deepEqual(result.spacingRoundtrip, result.spacing);
    assert.deepEqual(result.allCapsRoundtrip, result.allCaps);
    assert.deepEqual(result.smallCapsRoundtrip, result.smallCaps);
    assert.deepEqual(result.zeroSpacingRoundtrip, result.zeroSpacing);
    assert.equal(result.hostile.text, 'Safe');
    assert.equal(result.executed, false);
    assert.equal(result.mismatch, null);
    assert.equal(result.table, null);
    assert.equal(result.deep, null);
    assert.equal(result.pre.text, 'A  B\n日本語');
    assert.deepEqual(result.roundtrip, result.mixed);
    assert.equal(result.exportedText, '<script>&日本語');
    assert.equal(result.exportedScript, false);
    assert.deepEqual(
      requests.filter((url) => /^https?:/.test(url)),
      [],
    );
  } finally {
    await browser.close();
  }
});

for (const kind of ['shape', 'cell'])
  test(
    `external HTML paste retains saved formatting and undo in ${kind}`,
    { timeout: 60000 },
    async () => {
      const dir = await mkdtemp(join(tmpdir(), 'office-html-clipboard-'));
      let browser, preview;
      try {
        const file = join(dir, 'deck.tsx');
        const element =
          kind === 'shape'
            ? '<Text x={1} y={1} width={6} height={2}>old</Text>'
            : '<Table x={1} y={1} width={6} height={2} rows={[["old","Other"]]} />';
        await writeFile(
          file,
          `import {Presentation,Slide,Text,Table} from '@office-kit/pptx-dsl';export default <Presentation><Slide>${element}</Slide></Presentation>`,
        );
        preview = await startPreview(file);
        browser = await chromium.launch({ headless: true });
        const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
        const errors = [];
        page.on('pageerror', (e) => errors.push(e.message));
        await installRichTextSelection(page);
        await page.goto(preview.url);
        await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
        const editor = page.frameLocator('#editor-frame');
        let ja = false;
        const saved = () =>
          editor
            .getByText(ja ? 'このプロジェクトに保存済み' : 'Saved to this project', { exact: true })
            .waitFor();
        await saved();
        if (kind === 'cell') {
          await editor.locator('.lang select').selectOption('ja');
          ja = true;
        }
        await editor
          .locator('.hit')
          .first()
          .dblclick({ position: { x: 30, y: 20 } });
        const input = editor.locator('.inline-edit');
        await input.evaluate((node) => {
          window.selectEditorText(node, 0, 3);
          node.dispatchEvent(new Event('select', { bubbles: true }));
        });
        await input.evaluate((node) => {
          const data = new DataTransfer();
          data.setData('text/plain', 'English日本語');
          data.setData(
            'text/html',
            '<span style="font-size:24pt;color:#13579b"><b>English</b><i>日本語</i></span>',
          );
          node.dispatchEvent(
            new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
          );
        });
        assert.equal(await input.textContent(), 'English日本語');
        await input.press('Control+Enter');
        await saved();
        const runs = async () => {
          const pres = await loadPresentation(
            new Uint8Array(await (await fetch(preview.url + '/deck.pptx')).arrayBuffer()),
          );
          const shape = getSlideShapes(getSlides(pres)[0])[0];
          return kind === 'shape'
            ? getShapeParagraphElements(shape, 0)
            : getTableCellParagraphs(getTableCells(shape)[0][0])[0].elements;
        };
        const result = await runs();
        assert.equal(result[0].format.bold, true);
        assert.equal(result[0].format.size, 24);
        assert.equal(result[0].format.color.toUpperCase(), '#13579B');
        assert.equal(result.at(-1).format.italic, true);
        await editor
          .getByTitle(ja ? '元に戻す (Ctrl+Z)' : 'Undo (Ctrl+Z)', { exact: true })
          .click();
        await saved();
        assert.equal((await runs()).map((r) => r.text).join(''), 'old');
        assert.deepEqual(errors, []);
      } finally {
        await browser?.close();
        await preview?.close();
        await rm(dir, { recursive: true, force: true });
      }
    },
  );

test(
  'inline editing preserves all caps for newly typed characters',
  { timeout: 60000 },
  async () => {
    const dir = await mkdtemp(join(tmpdir(), 'office-cap-inline-'));
    let browser;
    let preview;
    try {
      const file = join(dir, 'deck.tsx');
      await writeFile(
        file,
        `import {Presentation,Slide,Text} from '@office-kit/pptx-dsl';export default <Presentation><Slide><Text x={1} y={1} width={7} height={2} paragraphs={[{runs:[{text:'hello',format:{cap:'all',size:24}}]}]} /></Slide></Presentation>`,
      );
      preview = await startPreview(file);
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
      await installRichTextSelection(page);
      await page.goto(preview.url);
      await page.getByRole('button', { name: '✦ Agents', exact: true }).click();
      const editor = page.frameLocator('#editor-frame');
      const saved = editor.getByText('Saved to this project', { exact: true });
      await saved.waitFor();
      await editor
        .locator('.hit')
        .first()
        .dblclick({ position: { x: 30, y: 20 } });
      const input = editor.locator('.inline-edit');
      await input.waitFor();
      assert.equal(
        await input
          .locator('span')
          .first()
          .evaluate((node) => getComputedStyle(node).textTransform),
        'uppercase',
      );
      await input.focus();
      await input.press('End');
      await input.press('!');
      await input.press('Control+Enter');
      await saved.waitFor();
      const pres = await loadPresentation(
        new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()),
      );
      const shape = getSlideShapes(getSlides(pres)[0])[0];
      const runs = getShapeParagraphElements(shape, 0);
      assert.equal(runs.map((run) => run.text).join(''), 'hello!');
      assert.equal(runs.at(-1).format.cap, 'all');
    } finally {
      await browser?.close();
      await preview?.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
