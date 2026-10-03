import assert from 'node:assert/strict';
import test from 'node:test';
import { build } from 'esbuild';
import { chromium } from 'playwright';
import { fileURLToPath } from 'node:url';

const underlineStyles = [
  'heavy',
  'dottedHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dotDotDashHeavy',
  'wavyDbl',
  'words',
];

test('HTML clipboard preserves advanced underline styles on paste', async () => {
  const bundle = await build({
    stdin: {
      contents: `
        import { parseHtmlTextClipboard, textClipboardHtml } from './html-text-clipboard.ts';
        export function roundTrip(underline, strike, editing, underlineColor) {
          const copied = { text: 'Alpha  Beta\\nGamma', formats: [{ start: 0, end: 17, format: { underline, strike, underlineColor } }] };
          const html = textClipboardHtml(copied, { editing });
          const parsed = parseHtmlTextClipboard(html, copied.text);
          return { html, parsed };
        }
        export function invalidMetadata() {
          return parseHtmlTextClipboard(
            '<span data-office-kit-underline="not-a-style" style="text-decoration:underline;text-decoration-style:dotted">X</span>',
            'X',
          );
        }
      `,
      loader: 'ts',
      resolveDir: fileURLToPath(new URL('../../../../site/src/lib/editor/core/', import.meta.url)),
    },
    bundle: true,
    format: 'esm',
    platform: 'browser',
    write: false,
  });
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage();
    for (const underline of underlineStyles) {
      for (const strike of [undefined, true, 'dblStrike']) {
        for (const underlineColor of [undefined, '#123456']) {
          for (const editing of [false, true]) {
            const result = await page.evaluate(
              async ({ code, underline, strike, editing, underlineColor }) => {
                const module = await import(
                  URL.createObjectURL(new Blob([code], { type: 'text/javascript' }))
                );
                return module.roundTrip(underline, strike, editing, underlineColor);
              },
              { code: bundle.outputFiles[0].text, underline, strike, editing, underlineColor },
            );
            assert.ok(result.parsed);
            assert.equal(result.parsed.text, 'Alpha  Beta\nGamma');
            assert.ok(
              result.parsed.formats.every(
                ({ format }) => format.underline === underline && format.strike === strike,
              ),
            );
            if (underline === 'words') {
              assert.equal(result.parsed.formats.at(-1)?.end, 17);
              assert.ok(
                result.parsed.formats
                  .filter(({ start, end }) => /\S/.test(result.parsed.text.slice(start, end)))
                  .every(
                    ({ format }) =>
                      underlineColor === undefined || format.underlineColor === underlineColor,
                  ),
              );
            } else {
              const expectedFormat = { underline };
              if (strike !== undefined) expectedFormat.strike = strike;
              if (underlineColor !== undefined) expectedFormat.underlineColor = underlineColor;
              assert.deepEqual(
                result.parsed.formats,
                [{ start: 0, end: 17, format: expectedFormat }],
                `${underline}/${strike}/${underlineColor}/${editing}`,
              );
            }
          }
        }
      }
    }
    const invalid = await page.evaluate(async (code) => {
      const module = await import(
        URL.createObjectURL(new Blob([code], { type: 'text/javascript' }))
      );
      return module.invalidMetadata();
    }, bundle.outputFiles[0].text);
    assert.deepEqual(invalid?.formats, [{ start: 0, end: 1, format: { underline: 'dotted' } }]);
  } finally {
    await browser.close();
  }
});
