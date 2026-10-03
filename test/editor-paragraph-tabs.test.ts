import { expect, it } from 'vitest';
import { editTabStops } from '../site/src/lib/editor/core/paragraph-tabs.ts';

it('applies tab edits independently to paragraphs with different existing stops', () => {
  const first = [{ positionEmu: 100, alignment: 'left' as const }];
  const second = [{ positionEmu: 200, alignment: 'decimal' as const }];
  const edit = { kind: 'set' as const, stop: { positionEmu: 300, alignment: 'right' as const } };
  expect(editTabStops(first, [edit])).toEqual([...first, edit.stop]);
  expect(editTabStops(second, [edit])).toEqual([...second, edit.stop]);
  expect(editTabStops(first, [edit, { kind: 'clear', positionEmu: 100 }])).toEqual([edit.stop]);
  expect(editTabStops(second, [edit, { kind: 'clearAll' }, edit])).toEqual([edit.stop]);
  expect(first).toEqual([{ positionEmu: 100, alignment: 'left' }]);
});
