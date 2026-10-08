import assert from 'node:assert/strict';
import test from 'node:test';
import { mergeTextFormat } from '../src/core/merge-text-format.ts';

const textFill = {
  kind: 'pattern',
  preset: 'dkUpDiag',
  foreground: 'accent1',
  background: '#FFFFFF',
};

test('choosing a font color replaces an inherited Text Art fill while typing', () => {
  assert.deepEqual(mergeTextFormat({ textFill, bold: true }, { color: '#123456' }), {
    color: '#123456',
    bold: true,
  });
  assert.deepEqual(mergeTextFormat({ textFill }, { color: null }), { color: null });
});

test('choosing a Text Art fill replaces a pending solid color while unrelated formatting retains it', () => {
  assert.deepEqual(mergeTextFormat({ color: '#123456', italic: true }, { textFill }), {
    textFill,
    italic: true,
  });
  assert.deepEqual(mergeTextFormat({ textFill }, { bold: true }), { textFill, bold: true });
});
