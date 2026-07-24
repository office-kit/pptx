// Editor localization.
//
// Strategy: the English string IS the message key, so wrapping a literal in
// `t('New')` needs no separate key table — only Japanese overrides are stored
// (`ja.ts`). A missing override falls back to the English key, so the UI is
// never blank in either language. Capability / category labels come from the
// manifest, which already carries `labelEn` / `labelJa`.
//
// `locale` is a module-level `$state`; because `t()` / `capLabel()` read it,
// any component template or `$derived` that calls them re-renders when the
// language switches. That is the whole reactivity story — no context, no props.

import { ja } from './ja.ts';

export type Locale = 'ja' | 'en';

const STORAGE_KEY = 'ok-editor-locale';

function detectInitial(): Locale {
  if (typeof localStorage !== 'undefined') {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored === 'ja' || stored === 'en') return stored;
  }
  if (typeof navigator !== 'undefined' && navigator.language?.toLowerCase().startsWith('ja')) {
    return 'ja';
  }
  return 'en';
}

let locale = $state<Locale>(detectInitial());

export function getLocale(): Locale {
  return locale;
}

export function setLocale(next: Locale): void {
  locale = next;
  if (typeof localStorage !== 'undefined') localStorage.setItem(STORAGE_KEY, next);
}

export const LOCALES: ReadonlyArray<{ id: Locale; label: string }> = [
  { id: 'en', label: 'English' },
  { id: 'ja', label: '日本語' },
];

/** Translate an English-keyed UI string to the active locale. */
export function t(key: string): string {
  if (locale === 'en') return key;
  return ja[key] ?? key;
}

/** A capability's label in the active locale. */
export function capLabel(cap: { labelEn: string; labelJa: string }): string {
  return locale === 'ja' ? cap.labelJa : cap.labelEn;
}

/** A category label (from `CATEGORY_LABELS`) in the active locale. */
export function catLabel(label: { en: string; ja: string }): string {
  return locale === 'ja' ? label.ja : label.en;
}
