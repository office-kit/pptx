import assert from 'node:assert/strict';
import test from 'node:test';
import { parseTextClipboard } from '../src/lib/editor/core/text-clipboard.ts';

const clipboard = (format) =>
  JSON.stringify({ version: 1, text: '日本語', formats: [{ start: 0, end: 3, format }] });

test('invalid character metadata falls back to plain text before paste', () => {
  for (const format of [
    { underline: 'invalid' },
    { strike: 'dotted' },
    { size: 0.1 },
    { spc: 400001 },
    { color: 'ACCENT1' },
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
    color: 'scheme:accent1',
  };
  assert.deepEqual(parseTextClipboard(clipboard(format), '日本語')?.formats[0].format, format);
});
