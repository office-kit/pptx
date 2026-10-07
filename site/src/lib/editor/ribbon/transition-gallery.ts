import type { SlideTransition, TransitionEffect, TransitionOptions } from '@office-kit/pptx';

// Mac PowerPoint 16's Transitions gallery and Effect Options menus, in its
// order and wording (English and Japanese). PowerPoint 2010 and later effects
// are written as PowerPoint writes them: p14 / p15 / p159 elements inside an
// mc:AlternateContent with a fade fallback.

/** What choosing a gallery tile or an Effect Options item writes. */
export type TransitionChoice = Pick<
  TransitionOptions,
  | 'effect'
  | 'direction'
  | 'orientation'
  | 'spokes'
  | 'thruBlack'
  | 'pattern'
  | 'isContent'
  | 'isInverted'
  | 'hasBounce'
  | 'preset'
  | 'invertX'
  | 'morphOption'
>;

export interface TransitionTile {
  readonly key: string;
  readonly en: string;
  readonly ja: string;
  /** What the tile writes: PowerPoint's default option for the effect. */
  readonly choice: TransitionChoice;
  /**
   * The Duration PowerPoint fills in when the effect is chosen, in
   * milliseconds. PowerPoint keeps one per effect rather than per speed; the
   * values are the ones Mac PowerPoint 16 shows (see POWERPOINT_PARITY.md).
   */
  readonly durationMs: number;
}

const tile = (
  en: string,
  ja: string,
  durationMs: number,
  choice: TransitionChoice,
): TransitionTile => ({ key: en, en, ja, choice, durationMs });

/** Duration shown, and written by a Duration edit, while a slide has no transition. */
export const NO_TRANSITION_DURATION_MS = 2000;

export const TRANSITION_TILES: readonly TransitionTile[] = [
  tile('None', 'なし', NO_TRANSITION_DURATION_MS, { effect: 'none' }),
  tile('Morph', '変形', 2000, { effect: 'morph', morphOption: 'byObject' }),
  tile('Fade', 'フェード', 700, { effect: 'fade' }),
  tile('Push', 'プッシュ', 1000, { effect: 'push', direction: 'u' }),
  tile('Wipe', 'ワイプ', 1000, { effect: 'wipe' }),
  tile('Split', 'スプリット', 1500, { effect: 'split', orientation: 'vert', direction: 'out' }),
  tile('Reveal', '出現', 1600, { effect: 'reveal', direction: 'l' }),
  tile('Cut', 'カット', 100, { effect: 'cut' }),
  tile('Random Bars', 'ランダム ストライプ', 1000, { effect: 'randomBar', direction: 'vert' }),
  tile('Shape', '図形', 2000, { effect: 'circle' }),
  tile('Uncover', 'アンカバー', 1000, { effect: 'pull' }),
  tile('Cover', 'カバー', 1000, { effect: 'cover' }),
  tile('Flash', 'フラッシュ', 750, { effect: 'flash' }),
  tile('Fall Over', 'フォール オーバー', 2000, { effect: 'prstTrans', preset: 'fallOver' }),
  tile('Drape', 'ドレープ', 2000, { effect: 'prstTrans', preset: 'drape' }),
  tile('Curtains', 'カーテン', 2250, { effect: 'prstTrans', preset: 'curtains' }),
  tile('Wind', '風', 2000, { effect: 'prstTrans', preset: 'wind', invertX: true }),
  tile('Prestige', 'プレステージ', 2000, { effect: 'prstTrans', preset: 'prestige' }),
  tile('Fracture', '割れる', 2000, { effect: 'prstTrans', preset: 'fracture' }),
  tile('Crush', 'クラッシュ', 2000, { effect: 'prstTrans', preset: 'crush' }),
  tile('Peel Off', 'ピール オフ', 1250, { effect: 'prstTrans', preset: 'peelOff' }),
  tile('Page Curl', 'ページ カール', 2000, { effect: 'prstTrans', preset: 'pageCurlDouble' }),
  tile('Airplane', '飛行機', 2000, { effect: 'prstTrans', preset: 'airplane', invertX: true }),
  tile('Origami', '折り紙', 2000, { effect: 'prstTrans', preset: 'origami', invertX: true }),
  tile('Dissolve', 'ディゾルブ', 1200, { effect: 'dissolve' }),
  tile('Checkerboard', 'チェッカーボード', 1500, { effect: 'checker' }),
  tile('Blinds', 'ブラインド', 1600, { effect: 'blinds', direction: 'vert' }),
  tile('Clock', '時計', 2000, { effect: 'wheel', spokes: 1 }),
  tile('Ripple', 'さざ波', 1400, { effect: 'ripple', direction: 'center' }),
  tile('Honeycomb', 'ハニカム', 3000, { effect: 'honeycomb' }),
  tile('Glitter', 'キラキラ', 2500, { effect: 'glitter', direction: 'r', pattern: 'hexagon' }),
  tile('Vortex', '渦巻き', 3000, { effect: 'vortex', direction: 'r' }),
  tile('Shred', '細断', 2500, { effect: 'shred', direction: 'in', pattern: 'strip' }),
  tile('Switch', 'スイッチ', 1250, { effect: 'switch', direction: 'r' }),
  tile('Flip', 'フリップ', 1250, { effect: 'flip', direction: 'r' }),
  tile('Gallery', 'ギャラリー', 1600, { effect: 'gallery', direction: 'l' }),
  tile('Cube', 'キューブ', 1250, { effect: 'prism', direction: 'l' }),
  tile('Doors', 'ドア', 1400, { effect: 'doors', direction: 'vert' }),
  tile('Box', 'ボックス', 1600, { effect: 'prism', direction: 'l', isInverted: true }),
  tile('Comb', 'コーム', 1000, { effect: 'comb' }),
  tile('Zoom', 'ズーム', 1200, { effect: 'warp', direction: 'in' }),
  tile('Random', 'ランダム', 2000, { effect: 'random' }),
  tile('Pan', 'パン', 1600, { effect: 'pan', direction: 'u' }),
  tile('Ferris Wheel', '観覧車', 2000, { effect: 'ferris', direction: 'l' }),
  tile('Conveyor', 'コンベヤー', 1600, { effect: 'conveyor', direction: 'l' }),
  tile('Rotate', '回転', 1200, { effect: 'prism', direction: 'l', isContent: true }),
  tile('Window', 'ウィンドウ', 1250, { effect: 'window', direction: 'vert' }),
  tile('Orbit', 'オービット', 1250, {
    effect: 'prism',
    direction: 'l',
    isContent: true,
    isInverted: true,
  }),
  tile('Fly Through', 'フライスルー', 1250, { effect: 'flythrough', direction: 'in' }),
];

// The gallery tile each written effect element belongs to. PowerPoint writes
// Wipe's diagonals as <p:strips>, Shape's variants as circle / diamond / plus
// and its In / Out as <p:zoom>, Clock's Counterclockwise and Wedge as
// <p14:wheelReverse> and <p:wedge>, Zoom as <p14:warp> and its "Zoom and
// Rotate" as <p:newsflash>. Cube, Box, Rotate and Orbit share <p14:prism>, and
// the PowerPoint 2013 effects share <p15:prstTrans>.
const TILE_OF_EFFECT: Readonly<Record<string, string>> = {
  none: 'None',
  morph: 'Morph',
  fade: 'Fade',
  push: 'Push',
  wipe: 'Wipe',
  strips: 'Wipe',
  split: 'Split',
  reveal: 'Reveal',
  cut: 'Cut',
  randomBar: 'Random Bars',
  circle: 'Shape',
  diamond: 'Shape',
  plus: 'Shape',
  zoom: 'Shape',
  pull: 'Uncover',
  cover: 'Cover',
  flash: 'Flash',
  dissolve: 'Dissolve',
  checker: 'Checkerboard',
  blinds: 'Blinds',
  wheel: 'Clock',
  wheelReverse: 'Clock',
  wedge: 'Clock',
  ripple: 'Ripple',
  honeycomb: 'Honeycomb',
  glitter: 'Glitter',
  vortex: 'Vortex',
  shred: 'Shred',
  switch: 'Switch',
  flip: 'Flip',
  gallery: 'Gallery',
  doors: 'Doors',
  comb: 'Comb',
  warp: 'Zoom',
  newsflash: 'Zoom',
  random: 'Random',
  pan: 'Pan',
  ferris: 'Ferris Wheel',
  conveyor: 'Conveyor',
  window: 'Window',
  flythrough: 'Fly Through',
};

const TILE_OF_PRESET: Readonly<Record<string, string>> = {
  fallOver: 'Fall Over',
  drape: 'Drape',
  curtains: 'Curtains',
  wind: 'Wind',
  prestige: 'Prestige',
  fracture: 'Fracture',
  crush: 'Crush',
  peelOff: 'Peel Off',
  pageCurlDouble: 'Page Curl',
  pageCurlSingle: 'Page Curl',
  airplane: 'Airplane',
  origami: 'Origami',
};

export const tileOfTransition = (transition: SlideTransition | null): string | null => {
  const effect = transition?.effect ?? 'none';
  if (effect === 'prstTrans') return TILE_OF_PRESET[transition?.preset ?? ''] ?? null;
  if (effect === 'prism') {
    const content = transition?.isContent === true;
    const inverted = transition?.isInverted === true;
    return content ? (inverted ? 'Orbit' : 'Rotate') : inverted ? 'Box' : 'Cube';
  }
  return TILE_OF_EFFECT[effect] ?? null;
};

/** The tile choosing `key` selects, for its default choice and duration. */
export const transitionTile = (key: string): TransitionTile | undefined =>
  TRANSITION_TILES.find((item) => item.key === key);

export interface TransitionOption {
  readonly en: string;
  readonly ja: string;
  readonly choice: TransitionChoice;
  /** Arrow angle in degrees (0 points right) drawn on the item's thumbnail. */
  readonly arrow?: number;
}

const option = (
  en: string,
  ja: string,
  choice: TransitionChoice,
  arrow?: number,
): TransitionOption => ({ en, ja, choice, ...(arrow === undefined ? {} : { arrow }) });

// ECMA-376 and [MS-PPTX] side / corner tokens name the way the slide moves, so
// "From Bottom" is `u` (it moves up) and "From Top-Right" is `ld`.
type Extra = Omit<TransitionChoice, 'effect' | 'direction'>;
const fromRight = (effect: TransitionEffect, extra: Extra = {}): TransitionOption =>
  option('From Right', '右から', { effect, direction: 'l', ...extra }, 180);
const fromLeft = (effect: TransitionEffect, extra: Extra = {}): TransitionOption =>
  option('From Left', '左から', { effect, direction: 'r', ...extra }, 0);
const fromTop = (effect: TransitionEffect, extra: Extra = {}): TransitionOption =>
  option('From Top', '上から', { effect, direction: 'd', ...extra }, 270);
const fromBottom = (effect: TransitionEffect, extra: Extra = {}): TransitionOption =>
  option('From Bottom', '下から', { effect, direction: 'u', ...extra }, 90);
const four = (effect: TransitionEffect, extra: Extra = {}): TransitionOption[] => [
  fromRight(effect, extra),
  fromTop(effect, extra),
  fromLeft(effect, extra),
  fromBottom(effect, extra),
];
const eight = (
  effect: TransitionEffect,
  corners: TransitionEffect = effect,
): TransitionOption[] => [
  ...four(effect),
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
// The two-way effects ([MS-PPTX] ST_TransitionLeftRightDirectionType).
const rightLeft = (effect: TransitionEffect): TransitionOption[] => [
  option('Right', '右', { effect, direction: 'r' }, 0),
  option('Left', '左', { effect, direction: 'l' }, 180),
];
const fromRightLeft = (effect: TransitionEffect): TransitionOption[] => [
  fromRight(effect),
  fromLeft(effect),
];
// The p15 presets point left; `invX` mirrors one to the right ([MS-PPTX]
// §2.4.3.8). The order is the menu's.
const presetSides = (
  preset: NonNullable<TransitionChoice['preset']>,
  rightFirst: boolean,
): TransitionOption[] => {
  const left = option('Left', '左', { effect: 'prstTrans', preset }, 180);
  const right = option('Right', '右', { effect: 'prstTrans', preset, invertX: true }, 0);
  return rightFirst ? [right, left] : [left, right];
};

export const TRANSITION_OPTIONS: Readonly<Record<string, readonly TransitionOption[]>> = {
  Morph: [
    option('Objects', 'オブジェクト', { effect: 'morph', morphOption: 'byObject' }),
    option('Words', '単語', { effect: 'morph', morphOption: 'byWord' }),
    option('Characters', '文字', { effect: 'morph', morphOption: 'byChar' }),
  ],
  Fade: [
    option('Smoothly', 'スムーズ', { effect: 'fade' }),
    option('Through Black', '黒いスクリーンから', { effect: 'fade', thruBlack: true }),
  ],
  Push: [fromBottom('push'), fromLeft('push'), fromRight('push'), fromTop('push')],
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
  Reveal: [
    option('Smoothly From Right', 'スムーズ (右から)', { effect: 'reveal', direction: 'l' }, 180),
    option('Smoothly From Left', 'スムーズ (左から)', { effect: 'reveal', direction: 'r' }, 0),
    option(
      'Through Black From Right',
      '黒いスクリーンから (右から)',
      { effect: 'reveal', direction: 'l', thruBlack: true },
      180,
    ),
    option(
      'Through Black From Left',
      '黒いスクリーンから (左から)',
      { effect: 'reveal', direction: 'r', thruBlack: true },
      0,
    ),
  ],
  Cut: [
    option('Smoothly', 'スムーズ', { effect: 'cut' }),
    option('Through Black', '黒いスクリーンから', { effect: 'cut', thruBlack: true }),
  ],
  'Random Bars': orientations('randomBar', true),
  Shape: [
    option('Circle', '円', { effect: 'circle' }),
    option('Diamond', 'ひし形', { effect: 'diamond' }),
    option('Plus', 'プラス', { effect: 'plus' }),
    option('In', 'イン', { effect: 'zoom', direction: 'in' }),
    option('Out', 'アウト', { effect: 'zoom', direction: 'out' }),
  ],
  Uncover: eight('pull'),
  Cover: eight('cover'),
  'Fall Over': presetSides('fallOver', false),
  Drape: presetSides('drape', false),
  Wind: presetSides('wind', true),
  'Peel Off': presetSides('peelOff', false),
  'Page Curl': [
    option('Double Left', '二重 (左)', { effect: 'prstTrans', preset: 'pageCurlDouble' }, 180),
    option(
      'Double Right',
      '二重 (右)',
      { effect: 'prstTrans', preset: 'pageCurlDouble', invertX: true },
      0,
    ),
    option('Single Left', '一重 (左)', { effect: 'prstTrans', preset: 'pageCurlSingle' }, 180),
    option(
      'Single Right',
      '一重 (右)',
      { effect: 'prstTrans', preset: 'pageCurlSingle', invertX: true },
      0,
    ),
  ],
  Airplane: presetSides('airplane', true),
  Origami: presetSides('origami', true),
  Checkerboard: [
    option('From Left', '左から', { effect: 'checker', direction: 'horz' }, 0),
    option('From Top', '上から', { effect: 'checker', direction: 'vert' }, 270),
  ],
  Blinds: orientations('blinds', true),
  Clock: [
    option('Clockwise', '時計回り', { effect: 'wheel', spokes: 1 }),
    option('Counterclockwise', '反時計回り', { effect: 'wheelReverse', spokes: 1 }),
    option('Wedge', 'くさび形', { effect: 'wedge' }),
  ],
  Ripple: [
    option('Center', '中央から', { effect: 'ripple', direction: 'center' }),
    option('From Top-Left', '左上から', { effect: 'ripple', direction: 'lu' }, 315),
    option('From Top-Right', '右上から', { effect: 'ripple', direction: 'ru' }, 225),
    option('From Bottom-Left', '左下から', { effect: 'ripple', direction: 'ld' }, 45),
    option('From Bottom-Right', '右下から', { effect: 'ripple', direction: 'rd' }, 135),
  ],
  Glitter: [
    ...(['hexagon', 'diamond'] as const).flatMap((pattern) => {
      const [en, ja] = pattern === 'hexagon' ? ['Hexagons', '六角形'] : ['Diamonds', 'ひし形'];
      return [
        option(
          `${en} from Left`,
          `${ja} (左から)`,
          { effect: 'glitter', direction: 'r', pattern },
          0,
        ),
        option(
          `${en} from Top`,
          `${ja} (上から)`,
          { effect: 'glitter', direction: 'd', pattern },
          270,
        ),
        option(
          `${en} from Right`,
          `${ja} (右から)`,
          { effect: 'glitter', direction: 'l', pattern },
          180,
        ),
        option(
          `${en} from Bottom`,
          `${ja} (下から)`,
          { effect: 'glitter', direction: 'u', pattern },
          90,
        ),
      ];
    }),
  ],
  Vortex: [fromLeft('vortex'), fromTop('vortex'), fromRight('vortex'), fromBottom('vortex')],
  Shred: [
    option('Strips In', 'ストリップ (内向き)', {
      effect: 'shred',
      pattern: 'strip',
      direction: 'in',
    }),
    option('Strips Out', 'ストリップ (外向き)', {
      effect: 'shred',
      pattern: 'strip',
      direction: 'out',
    }),
    option('Particles In', 'パーティクル (内向き)', {
      effect: 'shred',
      pattern: 'rectangle',
      direction: 'in',
    }),
    option('Particles Out', 'パーティクル (外向き)', {
      effect: 'shred',
      pattern: 'rectangle',
      direction: 'out',
    }),
  ],
  Switch: rightLeft('switch'),
  Flip: rightLeft('flip'),
  Gallery: fromRightLeft('gallery'),
  Cube: four('prism'),
  Doors: orientations('doors', true),
  Box: four('prism', { isInverted: true }),
  Comb: orientations('comb', false),
  Zoom: [
    option('In', 'イン', { effect: 'warp', direction: 'in' }),
    option('Out', 'アウト', { effect: 'warp', direction: 'out' }),
    option('Zoom and Rotate', 'ズームと回転', { effect: 'newsflash' }),
  ],
  Pan: [fromBottom('pan'), fromLeft('pan'), fromRight('pan'), fromTop('pan')],
  'Ferris Wheel': fromRightLeft('ferris'),
  Conveyor: fromRightLeft('conveyor'),
  Rotate: four('prism', { isContent: true }),
  Window: orientations('window', true),
  Orbit: four('prism', { isContent: true, isInverted: true }),
  'Fly Through': [
    option('In', 'イン', { effect: 'flythrough', direction: 'in' }),
    option('In with Bounce', 'イン (バウンド)', {
      effect: 'flythrough',
      direction: 'in',
      hasBounce: true,
    }),
    option('Out', 'アウト', { effect: 'flythrough', direction: 'out' }),
    option('Out with Bounce', 'アウト (バウンド)', {
      effect: 'flythrough',
      direction: 'out',
      hasBounce: true,
    }),
  ],
};

// The schema's attribute defaults (ECMA-376 pml.xsd, [MS-PPTX] §5.1), so a
// transition written without one still matches the option PowerPoint shows as
// checked. The left / right effects have no default direction.
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
  doors: 'horz',
  window: 'horz',
  split: 'out',
  zoom: 'out',
  warp: 'out',
  shred: 'in',
  flythrough: 'in',
  vortex: 'l',
  pan: 'l',
  glitter: 'l',
  prism: 'l',
  reveal: 'l',
  ripple: 'center',
};
const DEFAULT_PATTERN: Readonly<Record<string, string>> = { glitter: 'diamond', shred: 'strip' };
const FLAGS = ['thruBlack', 'isContent', 'isInverted', 'hasBounce', 'invertX'] as const;

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
  if (
    choice.pattern !== undefined &&
    choice.pattern !== (current.pattern ?? DEFAULT_PATTERN[current.effect])
  )
    return false;
  if (choice.preset !== undefined && choice.preset !== current.preset) return false;
  if ((choice.morphOption ?? 'byObject') !== (current.morphOption ?? 'byObject')) return false;
  return FLAGS.every((flag) => Boolean(choice[flag]) === Boolean(current[flag]));
}
