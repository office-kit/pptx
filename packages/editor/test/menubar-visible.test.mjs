import assert from 'node:assert/strict';
import test from 'node:test';
import { pruneMenus } from '../src/core/menubar-visible.ts';

const item = (id, children) => (children ? { id, label: id, children } : { id, label: id });
const labels = (entries) =>
  entries.map((entry) =>
    entry === '-' ? '-' : entry.children ? [entry.id, labels(entry.children)] : entry.id,
  );

test('hidden items leave no leading, trailing or doubled separators', () => {
  const menus = [
    {
      id: 'file',
      label: 'File',
      items: [item('a'), '-', item('b'), '-', item('c'), '-', item('d'), '-', item('e')],
    },
  ];
  const hidden = new Set(['a', 'c', 'e']);
  const [file] = pruneMenus(menus, (id) => hidden.has(id));
  assert.deepEqual(labels(file.items), ['b', '-', 'd']);
});

test('a submenu with every item hidden goes, and so does an emptied menu', () => {
  const menus = [
    {
      id: 'insert',
      label: 'Insert',
      items: [
        item('chart', [item('chart/pie'), item('chart/funnel')]),
        item('art', [item('art/x')]),
      ],
    },
    { id: 'window', label: 'Window', items: [item('window/min')] },
  ];
  const hidden = (id) => id === 'chart/funnel' || id.startsWith('art/') || id.startsWith('window/');
  assert.deepEqual(
    pruneMenus(menus, hidden).map((menu) => [menu.id, labels(menu.items)]),
    [['insert', [['chart', ['chart/pie']]]]],
  );
});

test('a hidden menu goes whole', () => {
  const menus = [
    { id: 'help', label: 'Help', items: [item('help/editor-help')] },
    { id: 'window', label: 'Window', items: [item('window/min')] },
  ];
  assert.deepEqual(
    pruneMenus(menus, (id) => id === 'window').map((menu) => menu.id),
    ['help'],
  );
});
