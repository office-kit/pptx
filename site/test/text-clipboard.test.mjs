import assert from 'node:assert/strict';
import test from 'node:test';
import { parseTextClipboard } from '../src/lib/editor/core/text-clipboard.ts';

const clipboard = (format) =>
  JSON.stringify({ version: 1, text: '日本語', formats: [{ start: 0, end: 3, format }] });

test('invalid character metadata falls back to plain text before paste', () => {
  for (const format of [
    { underline: 'invalid' },
    { strike: 'dotted' },
    { normalizeHeight: 'true' },
    { normalizeHeight: 1 },
    { size: 0.1 },
    { spc: 400001 },
    { kern: -1 },
    { kern: 400001 },
    { baseline: 30000 },
    { baseline: -30000 },
    { color: 'ACCENT1' },
    { underlineColor: 'ACCENT1' },
  ]) {
    assert.equal(parseTextClipboard(clipboard(format), '日本語'), null, JSON.stringify(format));
  }
});

test('supported character metadata preserves theme colors and detailed underline styles', () => {
  const format = {
    underline: 'wavyDbl',
    strike: 'dblStrike',
    size: 1,
    spc: -400000,
    kern: 400000,
    baseline: 0.3,
    color: 'scheme:accent1',
    underlineColor: '#123456',
  };
  assert.deepEqual(parseTextClipboard(clipboard(format), '日本語')?.formats[0].format, format);
});

test('underlineColor accepts null as follow-text metadata', () => {
  for (const underlineColor of [null, '#123456', 'scheme:accent1']) {
    const format = { underline: true, underlineColor };
    assert.deepEqual(parseTextClipboard(clipboard(format), '日本語')?.formats[0].format, format);
  }
});
