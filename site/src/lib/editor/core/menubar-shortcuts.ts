// Keyboard shortcuts of the menu bar, read from PowerPoint's own glyph strings
// (menubar-native.ts) so a menu's hint and its key can never disagree.

import type { NativeMenu, NativeMenuEntry, NativeMenuItem } from './menubar-native.ts';

export interface Chord {
  readonly cmd: boolean;
  readonly ctrl: boolean;
  readonly alt: boolean;
  readonly shift: boolean;
  /** The key after the modifier glyphs, e.g. `G`, `1`, `↩`. */
  readonly key: string;
}

type ChordKeyEvent = Pick<
  KeyboardEvent,
  'metaKey' | 'ctrlKey' | 'altKey' | 'shiftKey' | 'key' | 'code'
>;

const MODIFIERS: Readonly<Record<string, 'cmd' | 'ctrl' | 'alt' | 'shift'>> = {
  '⌘': 'cmd',
  '⌃': 'ctrl',
  '⌥': 'alt',
  '⇧': 'shift',
};

// PowerPoint names these by the character, which a US layout types with
// Shift; whether Shift is down is not part of the shortcut.
const SHIFTED_CHARACTERS = new Set(['?', '+', '^']);

/**
 * Parses a Mac glyph string such as `⌥⇧⌘G`. Returns null for keys a web page
 * never receives: the fn/Globe modifier (🌐) and the dictation key (🎤).
 */
export function parseShortcut(glyphs: string): Chord | null {
  if (glyphs.includes('🌐') || glyphs.includes('🎤')) return null;
  const chord = { cmd: false, ctrl: false, alt: false, shift: false };
  const characters = [...glyphs];
  let index = 0;
  for (; index < characters.length - 1; index++) {
    const modifier = MODIFIERS[characters[index]!];
    if (!modifier) break;
    chord[modifier] = true;
  }
  return { ...chord, key: characters.slice(index).join('') };
}

/** Physical codes for letters and digits: Option rewrites `key` on macOS. */
function keyMatches(key: string, event: ChordKeyEvent): boolean {
  if (/^[A-Z]$/.test(key)) return event.code === `Key${key}`;
  if (/^[0-9]$/.test(key)) return event.code === `Digit${key}`;
  switch (key) {
    case ',':
      return event.code === 'Comma';
    case '-':
      return event.code === 'Minus';
    // ⌘= is how ⌘+ is typed without Shift.
    case '+':
      return event.key === '+' || event.code === 'Equal';
    case '?':
      return event.key === '?' || (event.code === 'Slash' && event.shiftKey);
    case '^':
      return event.key === '^';
    case '↩':
      return event.key === 'Enter';
    case '⇥':
      return event.key === 'Tab';
    default:
      return false;
  }
}

function chordMatches(chord: Chord, event: ChordKeyEvent, controlAsCommand: boolean): boolean {
  const cmd = event.metaKey || (controlAsCommand && event.ctrlKey);
  const ctrl = !controlAsCommand && event.ctrlKey;
  return (
    chord.cmd === cmd &&
    chord.ctrl === ctrl &&
    chord.alt === event.altKey &&
    (SHIFTED_CHARACTERS.has(chord.key) || chord.shift === event.shiftKey) &&
    keyMatches(chord.key, event)
  );
}

function* shortcutItems(entries: readonly NativeMenuEntry[]): Generator<NativeMenuItem> {
  for (const entry of entries) {
    if (entry === '-') continue;
    if (entry.shortcut) yield entry;
    if (entry.children) yield* shortcutItems(entry.children);
  }
}

const chordCache = new WeakMap<readonly NativeMenu[], readonly [NativeMenuItem, Chord][]>();

function chords(menus: readonly NativeMenu[]): readonly [NativeMenuItem, Chord][] {
  let list = chordCache.get(menus);
  if (!list) {
    list = menus.flatMap((menu) =>
      [...shortcutItems(menu.items)].flatMap((item): [NativeMenuItem, Chord][] => {
        const chord = parseShortcut(item.shortcut!);
        return chord ? [[item, chord]] : [];
      }),
    );
    chordCache.set(menus, list);
  }
  return list;
}

/**
 * The menu item whose shortcut the key event types. Control stands in for
 * Command when no item takes the Control chord itself, which keeps Ctrl+Z,
 * Ctrl+T … working on keyboards and hosts without a Command key, as the
 * editor always has.
 */
export function menuItemForKey(
  menus: readonly NativeMenu[],
  event: ChordKeyEvent,
): NativeMenuItem | undefined {
  const list = chords(menus);
  const exact = list.find(([, chord]) => chordMatches(chord, event, false));
  if (exact || !event.ctrlKey || event.metaKey) return exact?.[0];
  return list.find(([, chord]) => chordMatches(chord, event, true))?.[0];
}

/** True when the chord holds Command or Control, so the browser may own it. */
export function isModifiedChord(chord: Chord): boolean {
  return chord.cmd || chord.ctrl;
}
