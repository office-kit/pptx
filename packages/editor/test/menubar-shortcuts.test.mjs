import assert from 'node:assert/strict';
import test from 'node:test';
import { NATIVE_MENUBAR } from '../src/core/menubar-native.ts';
import { menuItemForKey, parseShortcut } from '../src/core/menubar-shortcuts.ts';

const key = (code, modifiers = {}, value = code.replace(/^(Key|Digit)/, '').toLowerCase()) => ({
  code,
  key: value,
  metaKey: false,
  ctrlKey: false,
  altKey: false,
  shiftKey: false,
  ...modifiers,
});
const id = (locale, event) => menuItemForKey(NATIVE_MENUBAR[locale], event)?.id;

test('Mac glyph strings parse into modifiers and a key', () => {
  assert.deepEqual(parseShortcut('⌥⇧⌘G'), {
    cmd: true,
    ctrl: false,
    alt: true,
    shift: true,
    key: 'G',
  });
  assert.deepEqual(parseShortcut('⇧C'), {
    cmd: false,
    ctrl: false,
    alt: false,
    shift: true,
    key: 'C',
  });
  assert.equal(parseShortcut('⌃⇧⇥').key, '⇥');
  // fn/Globe and the dictation key never reach a web page.
  assert.equal(parseShortcut('🌐F'), null);
  assert.equal(parseShortcut('🌐⌃F'), null);
  assert.equal(parseShortcut('🎤'), null);
});

test('key events find PowerPoint’s menu item, by physical key', () => {
  assert.equal(id('en', key('KeyG', { metaKey: true, altKey: true }, '©')), 'arrange/group');
  assert.equal(
    id('en', key('KeyG', { metaKey: true, altKey: true, shiftKey: true }, '˝')),
    'arrange/ungroup',
  );
  assert.equal(id('en', key('KeyG', { metaKey: true })), 'edit/find/find-next');
  assert.equal(id('en', key('KeyD', { metaKey: true })), 'edit/duplicate');
  assert.equal(id('en', key('KeyD', { metaKey: true, shiftKey: true })), 'insert/duplicate-slide');
  assert.equal(id('en', key('Digit4', { metaKey: true })), 'view/outline-view');
  assert.equal(
    id('en', key('Digit1', { metaKey: true, altKey: true })),
    'view/master/slide-master',
  );
  assert.equal(id('en', key('Equal', { metaKey: true }, '=')), 'view/zoom/zoom-in');
  assert.equal(id('en', key('Minus', { metaKey: true }, '-')), 'view/zoom/zoom-out');
  assert.equal(
    id('en', key('Slash', { metaKey: true, shiftKey: true }, '?')),
    'help/powerpoint-help',
  );
  assert.equal(
    id('en', key('Enter', { metaKey: true }, 'Enter')),
    'slide-show/play-from-current-slide',
  );
  assert.equal(id('en', key('Enter', { altKey: true }, 'Enter')), 'view/presenter-view');
  assert.equal(id('en', key('KeyC', { shiftKey: true }, 'C')), 'format/crop');
  assert.equal(id('en', key('KeyC')), undefined);
});

test('Control stands in for Command unless PowerPoint gives the Control chord its own item', () => {
  assert.equal(id('en', key('KeyZ', { ctrlKey: true })), 'edit/undo');
  assert.equal(id('en', key('KeyT', { ctrlKey: true })), 'format/font');
  // ⌃F and ⌃H are Advanced Find and Replace, not ⌘F.
  assert.equal(id('en', key('KeyF', { ctrlKey: true })), 'edit/find/advanced-find');
  assert.equal(id('en', key('KeyH', { ctrlKey: true })), 'edit/find/replace');
  assert.equal(id('en', key('KeyV', { ctrlKey: true, metaKey: true })), 'edit/paste-special');
});

test('Pick Up Object Style is ⇧⌘C in English and ⌥⌘C in Japanese', () => {
  assert.equal(
    id('en', key('KeyC', { metaKey: true, shiftKey: true }, 'C')),
    'format/pick-up-object-style',
  );
  assert.equal(id('en', key('KeyC', { metaKey: true, altKey: true }, 'ç')), undefined);
  assert.equal(
    id('ja', key('KeyC', { metaKey: true, altKey: true }, 'ç')),
    'format/pick-up-object-style',
  );
  assert.equal(id('ja', key('KeyC', { metaKey: true, shiftKey: true }, 'C')), undefined);
  assert.equal(
    id('ja', key('KeyV', { metaKey: true, shiftKey: true }, 'V')),
    'format/apply-object-style',
  );
});

test('every chord names one command, except where PowerPoint repeats an item in two menus', () => {
  for (const locale of ['en', 'ja']) {
    const owners = new Map();
    const walk = (entries) => {
      for (const entry of entries) {
        if (entry === '-') continue;
        if (entry.shortcut && parseShortcut(entry.shortcut)) {
          owners.set(entry.shortcut, [...(owners.get(entry.shortcut) ?? []), entry.id]);
        }
        if (entry.children) walk(entry.children);
      }
    };
    for (const menu of NATIVE_MENUBAR[locale]) walk(menu.items);
    const shared = [...owners].filter(([, ids]) => ids.length > 1);
    assert.deepEqual(shared, [
      ['⌥↩', ['view/presenter-view', 'slide-show/presenter-view']],
      ['⇧⌘↩', ['view/slide-show', 'slide-show/play-from-start']],
    ]);
  }
});

test('the Japanese menus pair every item with an English one', () => {
  const ids = (menus) => {
    const out = [];
    const walk = (entries) => {
      for (const entry of entries) {
        if (entry === '-') continue;
        out.push(entry.id);
        if (entry.children) walk(entry.children);
      }
    };
    for (const menu of menus) walk(menu.items);
    return out.sort();
  };
  // Japanese PowerPoint has no Slide Show ▸ Rehearse with Coach.
  assert.deepEqual(
    ids(NATIVE_MENUBAR.ja),
    ids(NATIVE_MENUBAR.en).filter((item) => item !== 'slide-show/rehearse-with-coach'),
  );
});
