import { test } from 'node:test';
import assert from 'node:assert/strict';
import { findText } from '../src/text-search.ts';
const search = (text, query, options) =>
  findText(
    [{ shapes: [{ id: 1, text, textable: true, bounds: { x: 0, y: 0, w: 1, h: 1 } }] }],
    query,
    options,
  ).map(({ start, end }) => [start, end]);
test('literal search supports independent case and whole-word options', () => {
  assert.deepEqual(search('Cat cat scatter cat.', 'cat'), [
    [0, 3],
    [4, 7],
    [9, 12],
    [16, 19],
  ]);
  assert.deepEqual(search('Cat cat scatter cat.', 'cat', { matchCase: true, wholeWords: true }), [
    [4, 7],
    [16, 19],
  ]);
  assert.deepEqual(search('[a] [A]', '[a]', { matchCase: true }), [[0, 3]]);
  assert.deepEqual(search('Cat cat', '', { wholeWords: true }), []);
});
test('word boundaries respect Unicode letters, combining marks and UTF-16 offsets', () => {
  assert.deepEqual(search('écat cat猫 cat_ cat2 cat́ 😀cat!', 'cat', { wholeWords: true }), [
    [27, 30],
  ]);
  assert.deepEqual(search('𐐀cat cat𐐀 cat', 'cat', { wholeWords: true }), [[12, 15]]);
});
