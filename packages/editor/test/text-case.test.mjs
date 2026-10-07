import assert from 'node:assert/strict';
import test from 'node:test';
import { textCaseRange, textCaseSelection } from '../src/core/text-case.ts';

test('case target uses selected text or the caret word', () => {
  assert.deepEqual(textCaseRange('Outline title', { start: 3, end: 3 }), { start: 0, end: 7 });
  assert.deepEqual(textCaseRange('Outline title', { start: 8, end: 8 }), { start: 8, end: 13 });
  assert.deepEqual(textCaseRange('one  two', { start: 4, end: 4 }), { start: 4, end: 4 });
  assert.deepEqual(textCaseRange('one two', { start: 1, end: 5 }), { start: 1, end: 5 });
});
test('case selection tracks expanding Unicode mappings', () => {
  assert.deepEqual(
    textCaseSelection('straße', 'STRASSE', { start: 5, end: 5 }, { start: 0, end: 6 }),
    { start: 6, end: 6 },
  );
  assert.deepEqual(
    textCaseSelection('İ x', 'i\u0307 x', { start: 0, end: 1 }, { start: 0, end: 1 }),
    { start: 0, end: 2 },
  );
});
