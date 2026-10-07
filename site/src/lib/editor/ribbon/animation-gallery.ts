import type {
  AnimationDirection,
  AnimationEffect,
  AnimationPatch,
  SlideAnimationStep,
} from '@office-kit/pptx';

// Mac PowerPoint 16's Entrance, Emphasis and Exit galleries and each effect's
// Effect Options menu, in its order and wording (English and Japanese).
// Presets the library does not write are listed but unavailable.

export const UNSUPPORTED_EFFECT = 'This effect is not supported by the library yet.';

export interface EffectTile {
  readonly key: string;
  readonly en: string;
  readonly ja: string;
  readonly kind: 'entrance' | 'emphasis' | 'exit';
  readonly effect?: AnimationEffect;
  readonly unavailable?: string;
}

const tiles = (
  kind: EffectTile['kind'],
  list: ReadonlyArray<readonly [string, string, AnimationEffect?]>,
): EffectTile[] =>
  list.map(([en, ja, effect]) =>
    effect
      ? { key: en, en, ja, kind, effect }
      : { key: en, en, ja, kind, unavailable: UNSUPPORTED_EFFECT },
  );

export const ENTRANCE_TILES: readonly EffectTile[] = tiles('entrance', [
  ['Appear', 'アピール', 'appear'],
  ['Blinds', 'ブラインド', 'blindsIn'],
  ['Checkerboard', 'チェッカーボード', 'checkerboardIn'],
  ['Dissolve In', 'ディゾルブイン', 'dissolveIn'],
  ['Fly In', 'スライドイン', 'flyIn'],
  ['Peek In', 'クロール イン', 'peekIn'],
  ['Random Bars', 'ランダム ストライプ', 'randomBarsIn'],
  ['Shape', '図形', 'shapeIn'],
  ['Split', 'スプリット', 'splitIn'],
  ['Strips', 'ストリップス', 'stripsIn'],
  ['Wedge', 'くさび形', 'wedgeIn'],
  ['Wheel', 'ホイール', 'wheelIn'],
  ['Wipe', 'ワイプ', 'wipeIn'],
  ['Expand', 'エクスパンド'],
  ['Fade', 'フェード', 'fadeIn'],
  ['Swivel', 'ターン'],
  ['Zoom', 'ズーム', 'zoomIn'],
  ['Center Revolve', 'センター リボルブ'],
  ['Boomerang', 'ブーメラン'],
  ['Bounce', 'バウンド'],
  ['Credits', 'クレジット タイトル'],
  ['Curve Up', 'カーブ (上)'],
  ['Drop', 'ドロップ'],
  ['Flip', 'フリップ'],
  ['Float', 'フロート'],
  ['Pinwheel', '風車'],
  ['Spiral In', 'スパイラル イン'],
  ['Basic Swivel', 'ターン (基本)'],
  ['Whip', 'ホイップ'],
]);

export const EMPHASIS_TILES: readonly EffectTile[] = tiles('emphasis', [
  ['Fill Color', '塗りつぶしの色'],
  ['Font Color', 'フォントの色'],
  ['Grow/Shrink', '拡大/収縮'],
  ['Line Color', '線の色'],
  ['Spin', 'スピン', 'spin'],
  ['Transparency', '透過性'],
  ['Bold Flash', 'ボールド フラッシュ'],
  ['Brush Color', 'ブラシの色'],
  ['Complementary Color', '補色'],
  ['Complementary Color 2', '補色 2'],
  ['Contrasting Color', '対照色'],
  ['Darken', '暗く'],
  ['Desaturate', '彩度を下げる'],
  ['Lighten', '明るく'],
  ['Object Color', 'オブジェクトの色'],
  ['Pulse', 'パルス'],
  ['Underline', '下線'],
  ['Color Pulse', 'カラー パルス'],
  ['Grow With Color', '拡大 (色付き)'],
  ['Shimmer', 'シマー'],
  ['Teeter', 'シーソー'],
  ['Blink', 'ブリンク'],
  ['Bold Reveal', 'ボールド表示'],
  ['Wave', 'ウェーブ'],
]);

// The Exit Effects gallery: the entrance gallery's counterparts, in the same
// order, with PowerPoint's exit names where they differ.
export const EXIT_TILES: readonly EffectTile[] = tiles('exit', [
  ['Disappear', 'クリア', 'disappear'],
  ['Blinds', 'ブラインド', 'blindsOut'],
  ['Checkerboard', 'チェッカーボード', 'checkerboardOut'],
  ['Dissolve Out', 'ディゾルブアウト', 'dissolveOut'],
  ['Fly Out', 'スライドアウト', 'flyOut'],
  ['Peek Out', 'クロール アウト', 'peekOut'],
  ['Random Bars', 'ランダム ストライプ', 'randomBarsOut'],
  ['Shape', '図形', 'shapeOut'],
  ['Split', 'スプリット', 'splitOut'],
  ['Strips', 'ストリップス', 'stripsOut'],
  ['Wedge', 'くさび形', 'wedgeOut'],
  ['Wheel', 'ホイール', 'wheelOut'],
  ['Wipe', 'ワイプ', 'wipeOut'],
  ['Contract', 'コントラクト'],
  ['Fade', 'フェード', 'fadeOut'],
  ['Swivel', 'ターン'],
  ['Zoom', 'ズーム', 'zoomOut'],
  ['Center Revolve', 'センター リボルブ'],
  ['Boomerang', 'ブーメラン'],
  ['Bounce', 'バウンド'],
  ['Credits', 'クレジット タイトル'],
  ['Curve Down', 'カーブ (下)'],
  ['Drop', 'ドロップ'],
  ['Flip', 'フリップ'],
  ['Float', 'フロート'],
  ['Pinwheel', '風車'],
  ['Spiral Out', 'スパイラル アウト'],
  ['Basic Swivel', 'ターン (基本)'],
  ['Whip', 'ホイップ'],
]);

/** One item of an Effect Options section. */
export interface EffectOption {
  readonly en: string;
  readonly ja: string;
  /** What choosing it changes. */
  readonly patch: AnimationPatch;
  /** Arrow angle in degrees (0 points right), for the direction items. */
  readonly arrow?: number;
}

export interface EffectOptionSection {
  readonly heading: string;
  readonly items: readonly EffectOption[];
}

// The arrow shows where the shape travels: "From Bottom" points up. An exit
// travels the other way, so its arrow is turned round and its label says "To".
const direction = (
  value: AnimationDirection,
  en: string,
  ja: string,
  arrow: number,
): EffectOption => ({ en, ja, arrow, patch: { direction: value } });

const FLY_DIRECTIONS: readonly EffectOption[] = [
  direction('bottom', 'From Bottom', '下から', 90),
  direction('bottomLeft', 'From Bottom-Left', '左下から', 45),
  direction('left', 'From Left', '左から', 0),
  direction('topLeft', 'From Top-Left', '左上から', -45),
  direction('top', 'From Top', '上から', -90),
  direction('topRight', 'From Top-Right', '右上から', -135),
  direction('right', 'From Right', '右から', 180),
  direction('bottomRight', 'From Bottom-Right', '右下から', 135),
];

const EDGE_DIRECTIONS: readonly EffectOption[] = [
  direction('bottom', 'From Bottom', '下から', 90),
  direction('left', 'From Left', '左から', 0),
  direction('right', 'From Right', '右から', 180),
  direction('top', 'From Top', '上から', -90),
];

// Strips are named after the corner they start from, the way PowerPoint's
// menu lists them.
const STRIP_DIRECTIONS: readonly EffectOption[] = [
  direction('bottomLeft', 'Left Down', '左下', 45),
  direction('topLeft', 'Left Up', '左上', -45),
  direction('bottomRight', 'Right Down', '右下', 135),
  direction('topRight', 'Right Up', '右上', -135),
];

const ORIENTATIONS: readonly EffectOption[] = [
  { en: 'Horizontal', ja: '横', patch: { orientation: 'horizontal' } },
  { en: 'Vertical', ja: '縦', patch: { orientation: 'vertical' } },
];

const CHECKERBOARD: readonly EffectOption[] = [
  { en: 'Across', ja: '横', patch: { orientation: 'horizontal' } },
  { en: 'Down', ja: '縦', patch: { orientation: 'vertical' } },
];

const SPLITS: readonly EffectOption[] = [
  { en: 'Horizontal In', ja: '横 (内側へ)', patch: { orientation: 'horizontal', inOut: 'in' } },
  { en: 'Horizontal Out', ja: '横 (外側へ)', patch: { orientation: 'horizontal', inOut: 'out' } },
  { en: 'Vertical In', ja: '縦 (内側へ)', patch: { orientation: 'vertical', inOut: 'in' } },
  { en: 'Vertical Out', ja: '縦 (外側へ)', patch: { orientation: 'vertical', inOut: 'out' } },
];

const IN_OUT: readonly EffectOption[] = [
  { en: 'In', ja: 'イン', patch: { inOut: 'in' } },
  { en: 'Out', ja: 'アウト', patch: { inOut: 'out' } },
];

const SHAPES: readonly EffectOption[] = [
  { en: 'Circle', ja: '円', patch: { shape: 'circle' } },
  { en: 'Box', ja: 'ボックス', patch: { shape: 'box' } },
  { en: 'Diamond', ja: 'ひし形', patch: { shape: 'diamond' } },
  { en: 'Plus', ja: 'プラス', patch: { shape: 'plus' } },
];

const SPOKES: readonly EffectOption[] = [1, 2, 3, 4, 8].map((spokes) => ({
  en: spokes === 1 ? '1 Spoke' : `${spokes} Spokes`,
  ja: `スポーク ${spokes}`,
  patch: { spokes },
}));

const section = (heading: string, items: readonly EffectOption[]): EffectOptionSection => ({
  heading,
  items,
});

/** The Effect Options sections above Sequence for an effect, in PowerPoint's order. */
export const effectOptionSections = (effect: AnimationEffect | null): EffectOptionSection[] => {
  switch (effect?.replace(/(In|Out)$/, '')) {
    case 'fly':
      return [section('Direction', FLY_DIRECTIONS)];
    case 'wipe':
    case 'peek':
      return [section('Direction', EDGE_DIRECTIONS)];
    case 'strips':
      return [section('Direction', STRIP_DIRECTIONS)];
    case 'blinds':
    case 'randomBars':
      return [section('Direction', ORIENTATIONS)];
    case 'checkerboard':
      return [section('Direction', CHECKERBOARD)];
    case 'split':
      return [section('Direction', SPLITS)];
    case 'shape':
      return [section('Direction', IN_OUT), section('Shapes', SHAPES)];
    case 'wheel':
      return [section('Spokes', SPOKES)];
    default:
      return [];
  }
};

/** Whether the step already has every value the option sets. */
export const optionChecked = (step: SlideAnimationStep, option: EffectOption): boolean =>
  Object.entries(option.patch).every(
    ([name, value]) => step[name as keyof SlideAnimationStep] === value,
  );

/** The label an option shows for an exit: "To Bottom" for "From Bottom". */
export const optionLabel = (
  option: EffectOption,
  exit: boolean,
): { readonly en: string; readonly ja: string } =>
  exit && option.en.startsWith('From ')
    ? { en: `To ${option.en.slice('From '.length)}`, ja: option.ja.replace(/から$/, 'へ') }
    : option;

/** The directions an effect takes, for the animation pane's direction list. */
export const effectDirections = (effect: AnimationEffect): AnimationDirection[] =>
  effectOptionSections(effect)
    .flatMap((s) => s.items)
    .flatMap((item) => (item.patch.direction === undefined ? [] : [item.patch.direction]));

export const UNSUPPORTED_SEQUENCE =
  'Animating all paragraphs at once is not supported by the library yet.';
