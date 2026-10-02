import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { build } from 'esbuild';
import { chromium } from 'playwright';

test('HTML paste ignores sub-point sizes that cannot be written to a presentation', async () => {
  const module = await build({
    entryPoints: [
      fileURLToPath(
        new URL('../../../../site/src/lib/editor/core/html-text-clipboard.ts', import.meta.url),
      ),
    ],
    bundle: true,
    platform: 'browser',
    format: 'esm',
    write: false,
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    const results = await page.evaluate(async (source) => {
      const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
      const { parseHtmlTextClipboard: parse } = await import(url);
      URL.revokeObjectURL(url);
      return ['0.1pt', '1px', '1pt', '4000pt', '4001pt'].map(
        (size) => parse(`<span style="font-size:${size}">A</span>`, 'A').formats[0].format,
      );
    }, module.outputFiles[0].text);
    assert.deepEqual(
      results.map((format) => format.size),
      [undefined, undefined, 1, 4000, undefined],
    );
  } finally {
    await browser.close();
  }
});
