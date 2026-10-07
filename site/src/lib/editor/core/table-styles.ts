// PowerPoint's Table Styles gallery: the built-in styles grouped as PowerPoint
// groups them, their names in each UI language, and swatches drawn by the
// preview renderer with the slide's theme and the table's style options.
import {
  BUILTIN_TABLE_STYLES,
  addSlideTable,
  createPresentation,
  getSlideShapes,
  importSlide,
  inches,
  setShapeHidden,
  setSlideBackground,
  setSlideBackgroundGraphicsHidden,
  setSlideSize,
  setTableStyleFlags,
  setTableStyleId,
  type BuiltinTableStyle,
  type PresentationData,
  type SlideData,
  type SlideShapeData,
  type getTableStyleFlags,
} from '@office-kit/pptx';
import { renderSlideToSvg } from '@office-kit/pptx-preview';
import type { Locale } from '../i18n/i18n.svelte.ts';

export type TableStyleFlags = ReturnType<typeof getTableStyleFlags>;

// Headings as Mac PowerPoint words them (its mso strings 5a–5d).
export const TABLE_STYLE_GROUPS: readonly {
  readonly category: BuiltinTableStyle['category'];
  readonly en: string;
  readonly ja: string;
}[] = [
  { category: 'bestMatch', en: 'Best Match for Document', ja: 'ドキュメントに最適なスタイル' },
  { category: 'light', en: 'Light', ja: '淡色' },
  { category: 'medium', en: 'Medium', ja: '中間' },
  { category: 'dark', en: 'Dark', ja: '濃色' },
];

/** Styles per gallery row, as in PowerPoint: a family and its accents. */
export const TABLE_STYLES_PER_ROW = 7;

/** PowerPoint's "No Style, No Grid", which Clear Table applies. */
export const NO_STYLE_NO_GRID = BUILTIN_TABLE_STYLES[0].id;

// PowerPoint composes the localized names from patterns (its mso strings
// 4e–57); the English names are the `styleName` it writes.
const JA_FAMILY: Record<string, readonly [plain: string, accented: string]> = {
  Themed: ['テーマ スタイル', 'テーマ スタイル'],
  Light: ['スタイル (淡色)', '淡色スタイル'],
  Medium: ['スタイル (中間)', '中間スタイル'],
  Dark: ['スタイル (濃色)', '濃色スタイル'],
};
const JA_FIXED: Record<string, string> = {
  'No Style, No Grid': 'スタイルなし、表のグリッド線なし',
  'No Style, Table Grid': 'スタイルなし、表のグリッド線あり',
};
const NAME = /^(Themed|Light|Medium|Dark) Style (\d)(?: - Accent (\d)(?:\/Accent (\d))?)?$/;

/** A built-in style's name as PowerPoint shows it in `locale`. */
export function tableStyleName(name: string, locale: Locale): string {
  if (locale === 'en') return name;
  const fixed = JA_FIXED[name];
  if (fixed) return fixed;
  const match = NAME.exec(name);
  if (!match) return name;
  const [, family, number, accent, second] = match;
  const [plain, accented] = JA_FAMILY[family!]!;
  if (!accent) return `${plain} ${number}`;
  return `${accented} ${number} - アクセント ${accent}${second ? `/アクセント ${second}` : ''}`;
}

const FLAG_KEYS = ['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol'] as const;

interface SwatchDeck {
  readonly source: SlideData;
  readonly version: number;
  readonly flags: string;
  readonly pres: PresentationData;
  readonly slide: SlideData;
  readonly table: SlideShapeData;
  readonly images: Map<string, string>;
}

// A scratch deck holds a 5 × 5 table on a copy of the slide, so swatches use
// the slide's theme and color map. It is rebuilt when the slide, the deck
// version or the style options change; swatches are drawn on first request.
let deck: SwatchDeck | null = null;

function swatchDeck(source: SlideData, version: number, flags: TableStyleFlags): SwatchDeck {
  const key = FLAG_KEYS.map((flag) => (flags[flag] ? 1 : 0)).join('');
  if (deck && deck.source === source && deck.version === version && deck.flags === key) return deck;
  const pres = createPresentation();
  const slide = importSlide(pres, source);
  setShapeHidden(getSlideShapes(slide), true);
  // PowerPoint's minimum slide size is 1 inch; the swatch keeps its 76 × 56 shape.
  setSlideSize(pres, { width: inches(1.4), height: inches(1.04) });
  setSlideBackground(slide, '#FFFFFF');
  setSlideBackgroundGraphicsHidden(slide, true);
  const table = addSlideTable(slide, {
    x: inches(0.1),
    y: inches(0.1),
    w: inches(1.2),
    h: inches(0.84),
    rows: Array.from({ length: 5 }, () => Array.from({ length: 5 }, () => '')),
  });
  setTableStyleFlags(table, flags);
  deck = { source, version, flags: key, pres, slide, table, images: new Map() };
  return deck;
}

/**
 * Swatches (data: URLs) of `styleIds` on a 5 × 5 table with `flags`, in the
 * theme of `source`; `version` is the editor's document version.
 */
export function tableStyleSwatches(
  source: SlideData,
  version: number,
  flags: TableStyleFlags,
  styleIds: readonly string[],
): ReadonlyMap<string, string> {
  const { pres, slide, table, images } = swatchDeck(source, version, flags);
  for (const id of styleIds) {
    if (images.has(id)) continue;
    setTableStyleId(table, id);
    images.set(
      id,
      `data:image/svg+xml,${encodeURIComponent(renderSlideToSvg(pres, slide, { textLayout: 'svg' }))}`,
    );
  }
  return images;
}
