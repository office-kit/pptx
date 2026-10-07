// The resolved capability manifest — generated catalogue + human overrides.
//
// Consumers (registry, ribbon, command palette, coverage test) import from
// here, never from the raw JSON. This is the single list the whole editor
// treats as "everything the library can author".

import generated from './capabilities.generated.json';
import { overrides } from './overrides.ts';
import type { Capability, CategoryId, ResolvedCapability } from './types.ts';

const base = generated.capabilities as unknown as Capability[];
const baseIds = new Set(base.map((c) => c.id));

// Guard: an override must refine an existing capability, never invent one.
// (Coverage is enforced against the generated set in
// `test/editor-capability-coverage.test.ts`; this keeps overrides honest at
// module load so a stale/misspelled key fails loudly instead of silently.)
for (const id of Object.keys(overrides)) {
  if (!baseIds.has(id)) {
    throw new Error(
      `Capability override "${id}" does not match any generated capability. Re-run manifest/generate.mjs or fix the key.`,
    );
  }
}

/** Default English label from a camelCase id: `setShapeFill` → "Set shape fill". */
function humanize(id: string): string {
  const words = id.replace(/([a-z0-9])([A-Z])/g, '$1 $2').split(/\s+/);
  return words
    .map((w, i) => (i === 0 ? (w[0] ?? '').toUpperCase() + w.slice(1) : w.toLowerCase()))
    .join(' ');
}

export const capabilities: readonly ResolvedCapability[] = base
  .map((cap): ResolvedCapability => {
    const o = overrides[cap.id] ?? {};
    return {
      ...cap,
      category: o.category ?? cap.category,
      params: o.params ?? cap.params,
      labelEn: o.labelEn ?? humanize(cap.id),
      labelJa: o.labelJa ?? humanize(cap.id),
      ...(o.ribbonGroup !== undefined ? { ribbonGroup: o.ribbonGroup } : {}),
      primary: o.primary ?? false,
    };
  })
  .sort((a, b) => a.id.localeCompare(b.id));

/** Fast id → capability lookup. */
export const capabilityById: ReadonlyMap<string, ResolvedCapability> = new Map(
  capabilities.map((c) => [c.id, c]),
);

export function capabilitiesByCategory(category: CategoryId): ResolvedCapability[] {
  return capabilities.filter((c) => c.category === category);
}

export function capabilitiesForOperand(
  operand: ResolvedCapability['operand'],
): ResolvedCapability[] {
  return capabilities.filter((c) => c.operand === operand);
}

export const CATEGORY_ORDER: readonly CategoryId[] = [
  'slide',
  'shape',
  'text',
  'paragraph',
  'fill',
  'stroke',
  'effect',
  'image',
  'table',
  'chart',
  'animation',
  'transition',
  'hyperlink',
  'comment',
  'notes',
  'slide-background',
  'section',
  'theme',
  'presentation',
  'misc',
];

export const CATEGORY_LABELS: Record<CategoryId, { en: string; ja: string }> = {
  slide: { en: 'Slides', ja: 'スライド' },
  shape: { en: 'Shapes', ja: '図形' },
  text: { en: 'Text', ja: 'テキスト' },
  paragraph: { en: 'Paragraph', ja: '段落' },
  fill: { en: 'Fill', ja: '塗りつぶし' },
  stroke: { en: 'Line', ja: '線' },
  effect: { en: 'Effects', ja: '効果' },
  image: { en: 'Picture', ja: '画像' },
  table: { en: 'Table', ja: '表' },
  chart: { en: 'Chart', ja: 'グラフ' },
  animation: { en: 'Animation', ja: 'アニメーション' },
  transition: { en: 'Transitions', ja: '画面切り替え' },
  hyperlink: { en: 'Hyperlink', ja: 'ハイパーリンク' },
  comment: { en: 'Comments', ja: 'コメント' },
  notes: { en: 'Notes', ja: 'ノート' },
  'slide-background': { en: 'Background', ja: '背景' },
  section: { en: 'Sections', ja: 'セクション' },
  theme: { en: 'Design', ja: 'デザイン' },
  presentation: { en: 'Document', ja: 'ドキュメント' },
  misc: { en: 'Other', ja: 'その他' },
};

export type { Capability, CategoryId, ResolvedCapability } from './types.ts';
