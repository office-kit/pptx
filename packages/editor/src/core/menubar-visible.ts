import type { NativeMenu, NativeMenuEntry } from './menubar-native.ts';

/**
 * The menus without the items `hide` names. A submenu left with no items goes
 * too, as does a menu, and separators never lead, trail or double up.
 */
export function pruneMenus(
  menus: readonly NativeMenu[],
  hide: (id: string) => boolean,
): readonly NativeMenu[] {
  return menus
    .filter((menu) => !hide(menu.id))
    .map((menu) => ({ ...menu, items: pruneEntries(menu.items, hide) }))
    .filter((menu) => menu.items.length > 0);
}

function pruneEntries(
  entries: readonly NativeMenuEntry[],
  hide: (id: string) => boolean,
): readonly NativeMenuEntry[] {
  const out: NativeMenuEntry[] = [];
  for (const entry of entries) {
    if (entry === '-') {
      if (out.length > 0 && out.at(-1) !== '-') out.push(entry);
      continue;
    }
    if (hide(entry.id)) continue;
    if (!entry.children?.length) {
      out.push(entry);
      continue;
    }
    const children = pruneEntries(entry.children, hide);
    if (children.length > 0) out.push({ ...entry, children });
  }
  if (out.at(-1) === '-') out.pop();
  return out;
}
