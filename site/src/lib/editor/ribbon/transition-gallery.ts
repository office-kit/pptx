import type { SlideTransition, TransitionEffect, TransitionOptions } from '@office-kit/pptx';

// Mac PowerPoint 16's Transitions gallery and Effect Options menus, in its
// order and wording (English and Japanese). Entries whose effect is a
// PowerPoint 2010+ extension (p14 / p15) are listed but unavailable: the
// library writes the ECMA-376 effect elements only.

export const P14_TRANSITION =
  'PowerPoint 2010 and later transitions are not supported by the library yet.';

/** What choosing a gallery tile or an Effect Options item writes. */
export type TransitionChoice = Pick<
  TransitionOptions,
  'effect' | 'direction' | 'orientation' | 'spokes' | 'thruBlack'
>;

export interface TransitionTile {
  readonly key: string;
  readonly en: string;
  readonly ja: string;
  readonly choice?: TransitionChoice;
  readonly unavailable?: string;
}

const tile = (en: string, ja: string, choice?: TransitionChoice): TransitionTile =>
  choice ? { key: en, en, ja, choice } : { key: en, en, ja, unavailable: P14_TRANSITION };

export const TRANSITION_TILES: readonly TransitionTile[] = [
  tile('None', 'なし', { effect: 'none' }),
  tile('Morph', '変形'),
  tile('Fade', 'フェード', { effect: 'fade' }),
  tile('Push', 'プッシュ', { effect: 'push', direction: 'u' }),
  tile('Wipe', 'ワイプ', { effect: 'wipe' }),
  tile('Split', 'スプリット', { effect: 'split', orientation: 'vert', direction: 'out' }),
  tile('Reveal', '出現'),
  tile('Cut', 'カット', { effect: 'cut' }),
  tile('Random Bars', 'ランダム ストライプ', { effect: 'randomBar', direction: 'vert' }),
  tile('Shape', '図形', { effect: 'circle' }),
  tile('Uncover', 'アンカバー', { effect: 'pull' }),
  tile('Cover', 'カバー', { effect: 'cover' }),
  tile('Flash', 'フラッシュ'),
  tile('Fall Over', 'フォール オーバー'),
  tile('Drape', 'ドレープ'),
  tile('Curtains', 'カーテン'),
  tile('Wind', '風'),
  tile('Prestige', 'プレステージ'),
  tile('Fracture', '割れる'),
  tile('Crush', 'クラッシュ'),
  tile('Peel Off', 'ピール オフ'),
  tile('Page Curl', 'ページ カール'),
  tile('Airplane', '飛行機'),
  tile('Origami', '折り紙'),
  tile('Dissolve', 'ディゾルブ', { effect: 'dissolve' }),
  tile('Checkerboard', 'チェッカーボード', { effect: 'checker' }),
  tile('Blinds', 'ブラインド', { effect: 'blinds', direction: 'vert' }),
  tile('Clock', '時計', { effect: 'wheel', spokes: 1 }),
  tile('Ripple', 'さざ波'),
  tile('Honeycomb', 'ハニカム'),
  tile('Glitter', 'キラキラ'),
  tile('Vortex', '渦巻き'),
  tile('Shred', '細断'),
  tile('Switch', 'スイッチ'),
  tile('Flip', 'フリップ'),
  tile('Gallery', 'ギャラリー'),
  tile('Cube', 'キューブ'),
  tile('Doors', 'ドア'),
  tile('Box', 'ボックス'),
  tile('Comb', 'コーム', { effect: 'comb' }),
  tile('Zoom', 'ズーム', { effect: 'zoom' }),
  tile('Random', 'ランダム', { effect: 'random' }),
  tile('Pan', 'パン'),
  tile('Ferris Wheel', '観覧車'),
  tile('Conveyor', 'コンベヤー'),
  tile('Rotate', '回転'),
  tile('Window', 'ウィンドウ'),
  tile('Orbit', 'オービット'),
  tile('Fly Through', 'フライスルー'),
];

// The gallery tile each written effect element belongs to. PowerPoint writes
// Wipe's diagonals as <p:strips>, Shape's variants as circle/diamond/plus,
// Clock's Wedge as <p:wedge> and Zoom's "Zoom and Rotate" as <p:newsflash>.
const TILE_OF_EFFECT: Readonly<Record<string, string>> = {
  none: 'None',
  fade: 'Fade',
  push: 'Push',
  wipe: 'Wipe',
  strips: 'Wipe',
  split: 'Split',
  cut: 'Cut',
  randomBar: 'Random Bars',
  circle: 'Shape',
  diamond: 'Shape',
  plus: 'Shape',
  pull: 'Uncover',
  cover: 'Cover',
  dissolve: 'Dissolve',
  checker: 'Checkerboard',
  blinds: 'Blinds',
  wheel: 'Clock',
  wedge: 'Clock',
  comb: 'Comb',
  zoom: 'Zoom',
  newsflash: 'Zoom',
  random: 'Random',
};

export const tileOfTransition = (transition: SlideTransition | null): string | null =>
  TILE_OF_EFFECT[transition?.effect ?? 'none'] ?? null;

export interface TransitionOption {
  readonly en: string;
  readonly ja: string;
  readonly choice?: TransitionChoice;
  readonly unavailable?: string;
  /** Arrow angle in degrees (0 points right) drawn on the item's thumbnail. */
  readonly arrow?: number;
}

const option = (
  en: string,
  ja: string,
  choice: TransitionChoice | null,
  arrow?: number,
): TransitionOption => ({
  en,
  ja,
  ...(choice ? { choice } : { unavailable: P14_TRANSITION }),
  ...(arrow === undefined ? {} : { arrow }),
});

// ECMA-376 side / corner tokens name the way the slide moves, so "From
// Bottom" is `u` (it moves up) and "From Top-Right" is `ld`.
const eight = (
  effect: TransitionEffect,
  corners: TransitionEffect = effect,
): TransitionOption[] => [
  option('From Right', '右から', { effect, direction: 'l' }, 180),
  option('From Top', '上から', { effect, direction: 'd' }, 270),
  option('From Left', '左から', { effect, direction: 'r' }, 0),
  option('From Bottom', '下から', { effect, direction: 'u' }, 90),
  option('From Top-Right', '右上から', { effect: corners, direction: 'ld' }, 225),
  option('From Bottom-Right', '右下から', { effect: corners, direction: 'lu' }, 135),
  option('From Top-Left', '左上から', { effect: corners, direction: 'rd' }, 315),
  option('From Bottom-Left', '左下から', { effect: corners, direction: 'ru' }, 45),
];
const orientations = (effect: TransitionEffect, verticalFirst: boolean): TransitionOption[] => {
  const vertical = option('Vertical', '縦', { effect, direction: 'vert' });
  const horizontal = option('Horizontal', '横', { effect, direction: 'horz' });
  return verticalFirst ? [vertical, horizontal] : [horizontal, vertical];
};

export const TRANSITION_OPTIONS: Readonly<Record<string, readonly TransitionOption[]>> = {
  Fade: [
    option('Smoothly', 'スムーズ', { effect: 'fade' }),
    option('Through Black', '黒いスクリーンから', { effect: 'fade', thruBlack: true }),
  ],
  Push: [
    option('From Bottom', '下から', { effect: 'push', direction: 'u' }, 90),
    option('From Left', '左から', { effect: 'push', direction: 'r' }, 0),
    option('From Right', '右から', { effect: 'push', direction: 'l' }, 180),
    option('From Top', '上から', { effect: 'push', direction: 'd' }, 270),
  ],
  Wipe: eight('wipe', 'strips'),
  Split: [
    option('Vertical Out', '縦 (外向き)', {
      effect: 'split',
      orientation: 'vert',
      direction: 'out',
    }),
    option('Vertical In', '縦 (内向き)', { effect: 'split', orientation: 'vert', direction: 'in' }),
    option('Horizontal Out', '横 (外向き)', {
      effect: 'split',
      orientation: 'horz',
      direction: 'out',
    }),
    option('Horizontal In', '横 (内向き)', {
      effect: 'split',
      orientation: 'horz',
      direction: 'in',
    }),
  ],
  'Random Bars': orientations('randomBar', true),
  Shape: [
    option('Circle', '円', { effect: 'circle' }),
    option('Diamond', 'ひし形', { effect: 'diamond' }),
    option('Plus', 'プラス', { effect: 'plus' }),
    option('In', 'イン', null),
    option('Out', 'アウト', null),
  ],
  Uncover: eight('pull'),
  Cover: eight('cover'),
  Checkerboard: [
    option('From Left', '左から', { effect: 'checker', direction: 'horz' }, 0),
    option('From Top', '上から', { effect: 'checker', direction: 'vert' }, 270),
  ],
  Blinds: orientations('blinds', true),
  Clock: [
    option('Clockwise', '時計回り', { effect: 'wheel', spokes: 1 }),
    option('Counterclockwise', '反時計回り', null),
    option('Wedge', 'くさび形', { effect: 'wedge' }),
  ],
  Comb: orientations('comb', false),
  Zoom: [
    option('In', 'イン', { effect: 'zoom', direction: 'in' }),
    option('Out', 'アウト', { effect: 'zoom', direction: 'out' }),
    option('Zoom and Rotate', 'ズームと回転', { effect: 'newsflash' }),
  ],
};

// The ECMA-376 attribute defaults, so a transition written without one still
// matches the option PowerPoint shows as checked.
const DEFAULT_DIRECTION: Readonly<Record<string, string>> = {
  push: 'l',
  wipe: 'l',
  cover: 'l',
  pull: 'l',
  strips: 'lu',
  randomBar: 'horz',
  blinds: 'horz',
  checker: 'horz',
  comb: 'horz',
  split: 'out',
  zoom: 'in',
};

export function optionMatches(choice: TransitionChoice, current: SlideTransition): boolean {
  if (choice.effect !== current.effect) return false;
  if (
    choice.direction !== undefined &&
    choice.direction !== (current.direction ?? DEFAULT_DIRECTION[current.effect])
  )
    return false;
  if (choice.orientation !== undefined && choice.orientation !== (current.orientation ?? 'horz'))
    return false;
  if (choice.spokes !== undefined && choice.spokes !== (current.spokes ?? 4)) return false;
  if (current.effect === 'fade' && Boolean(choice.thruBlack) !== Boolean(current.thruBlack))
    return false;
  return true;
}
