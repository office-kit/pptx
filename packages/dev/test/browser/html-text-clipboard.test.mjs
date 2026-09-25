import { installRichTextSelection } from '../helpers/rich-text.mjs';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { transform } from 'esbuild';
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
    const result = await page.evaluate(async (source) => {
      const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
      const { parseHtmlTextClipboard: parse, textClipboardHtml: serialize } = await import(url);
      URL.revokeObjectURL(url);
      const mixed = parse(
        '<b style="font-weight:normal"><span style="font-weight:700;font-size:24px;color:rgb(255,0,0);font-family:Arial">English</span><i style="background-color:yellow;text-decoration:underline line-through">日本語</i></b>',
        'English日本語',
      );
      const multiline = parse(
        '<p><strong>A</strong></p><p></p><p><sub>日本語</sub><br>End</p>',
        'A\n\n日本語\nEnd',
      );
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
      return {
        inherited: parse(
          '<b><i><span style="font-weight:inherit;font-style:inherit">A</span></i></b>',
          'A',
        ),
        mixed,
        multiline,
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
    }, module.code);
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
