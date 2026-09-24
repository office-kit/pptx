import { test } from 'node:test';
import assert from 'node:assert/strict';
import { textCaseEdits } from '../src/text-case.ts';

test('limits case edits to the selection and preserves contextual Unicode casing', () => {
  const apply = (text, mode, start, end) => {
    for (const edit of textCaseEdits(text, mode, { start, end }))
      text = text.slice(0, edit.start) + edit.text + text.slice(edit.end);
    return text;
  };
  assert.equal(apply('Before Straße after', 'upper', 7, 13), 'Before STRASSE after');
  assert.equal(apply('Before ΟΣ after', 'lower', 7, 9), 'Before ος after');
  assert.equal(apply('A😀B', 'lower', 0, 4), 'a😀b');
  assert.equal(apply('ABC', 'lower', 1, 1), 'ABC');
  assert.throws(() => textCaseEdits('text', 'lower', { start: -1, end: 2 }));
});

test('sentence, word and toggle modes retain punctuation and selection boundaries', () => {
  const apply = (text, mode, start = 0, end = text.length) => {
    for (const edit of textCaseEdits(text, mode, { start, end }))
      text = text.slice(0, edit.start) + edit.text + text.slice(edit.end);
    return text;
  };
  assert.equal(
    apply('“HELLO WORLD!” NEXT LINE.\nANOTHER LINE.', 'sentence'),
    '“Hello world!” Next line.\nAnother line.',
  );
  assert.equal(apply('élÈVE don’T STOP 123abc 日本語', 'title'), 'Élève Don’t Stop 123Abc 日本語');
  assert.equal(apply('HeLLo Straße 😀', 'toggle'), 'hEllO sTRASSE 😀');
  assert.equal(apply('LEAVE hELLO wORLD LEAVE', 'sentence', 6, 17), 'LEAVE Hello world LEAVE');
  assert.equal(apply('ΟΣ ΑΣ', 'title'), 'Ος Ας');
  assert.throws(() => textCaseEdits('text', 'invalid'));
});
