import assert from 'node:assert/strict';
import test from 'node:test';
import { COLOR_SETS, FONT_PAIRS, presetName, THEMES } from '../src/core/design-presets.ts';

test("a deck's built-in color-scheme names from the reference desktop app map to the gallery's", () => {
  const names = new Set(COLOR_SETS.map((set) => set.name));
  for (const [written, shown] of [
    ['Office', 'Standard'],
    ['Office 2013 - 2022', 'Classic'],
    ['Office 2007 - 2010', 'Legacy'],
  ]) {
    assert.equal(presetName(written), shown);
    assert.ok(names.has(shown), shown);
  }
  for (const name of ['Standard', 'Blue Warm', 'My Colors', ''])
    assert.equal(presetName(name), name);
});

test('the first three themes pair the same-named color set and font pair', () => {
  assert.deepEqual(
    THEMES.slice(0, 3).map((theme) => [theme.name, theme.colors.name, theme.fonts.name]),
    [
      ['Standard Theme', 'Standard', 'Standard'],
      ['Classic Theme', 'Classic', 'Classic'],
      ['Legacy Theme', 'Legacy', 'Legacy'],
    ],
  );
  assert.deepEqual(
    FONT_PAIRS.slice(0, 3).map((pair) => pair.name),
    ['Standard', 'Classic', 'Legacy'],
  );
});
