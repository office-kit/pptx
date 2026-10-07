import type {
  AnimationDirection,
  AnimationEffect,
  AnimationPatch,
  AnimationTextBuild,
  SlideAnimationStep,
} from '@office-kit/pptx';

// Mac PowerPoint 16's Entrance, Emphasis and Exit galleries and each effect's
// Effect Options menu, in its order and wording (English and Japanese). The
// keys are PowerPoint's gallery item ids (`<group>_<index>`, groups Basic,
// Subtle, Moderate and Exciting); the names are its AX labels, with the "&"
// the labels drop put back.

export interface EffectTile {
  readonly key: string;
  readonly en: string;
  readonly ja: string;
  readonly kind: 'entrance' | 'emphasis' | 'exit';
  readonly effect: AnimationEffect;
}

const tiles = (
  kind: EffectTile['kind'],
  list: ReadonlyArray<readonly [string, string, string, AnimationEffect]>,
): EffectTile[] => list.map(([key, en, ja, effect]) => ({ key, en, ja, kind, effect }));

export const ENTRANCE_TILES: readonly EffectTile[] = tiles('entrance', [
  ['0_0', 'Appear', '表示', 'appear'],
  ['0_1', 'Blinds', 'ブラインド', 'blindsIn'],
  ['0_2', 'Checkerboard', 'チェッカーボード', 'checkerboardIn'],
  ['0_3', 'Dissolve In', 'ディゾルブイン', 'dissolveIn'],
  ['0_4', 'Fly In', 'スライドイン', 'flyIn'],
  ['0_5', 'Peek In', 'ピークイン', 'peekIn'],
  ['0_6', 'Random Bars', 'ランダムストライプ', 'randomBarsIn'],
  ['0_7', 'Shape', '図形', 'shapeIn'],
  ['0_8', 'Split', 'スプリット', 'splitIn'],
  ['0_9', 'Strips', 'ストリップ', 'stripsIn'],
  ['0_10', 'Wedge', 'くさび形', 'wedgeIn'],
  ['0_11', 'Wheel', 'ホイール', 'wheelIn'],
  ['0_12', 'Wipe', 'ワイプ', 'wipeIn'],
  ['1_0', 'Expand', 'エクスパンド', 'expandIn'],
  ['1_1', 'Fade', 'フェード', 'fadeIn'],
  ['1_2', 'Swivel', 'ターン', 'swivelIn'],
  ['1_3', 'Zoom', 'ズーム', 'zoomIn'],
  ['2_0', 'Center Revolve', 'リボルブ', 'centerRevolveIn'],
  ['2_1', 'Float In', 'フロートイン', 'floatIn'],
  ['2_2', 'Grow & Turn', 'グローとターン', 'growTurnIn'],
  ['2_3', 'Rise Up', 'ライズ アップ', 'riseUpIn'],
  ['2_4', 'Spinner', 'スピナー', 'spinnerIn'],
  ['2_5', 'Basic Zoom', 'ベーシック ズーム', 'basicZoomIn'],
  ['2_6', 'Stretch', 'ストレッチ', 'stretchIn'],
  ['3_0', 'Boomerang', 'ブーメラン', 'boomerangIn'],
  ['3_1', 'Bounce', 'バウンド', 'bounceIn'],
  ['3_2', 'Credits', 'クレジット タイトル', 'creditsIn'],
  ['3_3', 'Curve Up', 'カーブ (上)', 'curveUpIn'],
  ['3_4', 'Drop', 'ドロップ', 'dropIn'],
  ['3_5', 'Flip', 'フリップ', 'flipIn'],
  ['3_6', 'Float', 'フロート', 'floatingIn'],
  ['3_7', 'Pinwheel', 'ピンウィール', 'pinwheelIn'],
  ['3_8', 'Spiral In', 'スパイラルイン', 'spiralIn'],
  ['3_9', 'Basic Swivel', 'ベーシック ターン', 'basicSwivelIn'],
  ['3_10', 'Whip', 'ホイップ', 'whipIn'],
]);

export const EMPHASIS_TILES: readonly EffectTile[] = tiles('emphasis', [
  ['0_0', 'Fill Color', '塗りつぶしの色', 'fillColor'],
  ['0_1', 'Font Color', 'フォントの色', 'fontColor'],
  ['0_2', 'Grow/Shrink', '拡大/収縮', 'growShrink'],
  ['0_3', 'Line Color', '線の色', 'lineColor'],
  ['0_4', 'Spin', 'スピン', 'spin'],
  ['0_5', 'Transparency', '透過性', 'transparency'],
  ['1_0', 'Bold Flash', 'ボールドフラッシュ', 'boldFlash'],
  ['1_1', 'Brush Color', 'ブラシの色', 'brushColor'],
  ['1_2', 'Complementary Color', '補色', 'complementaryColor'],
  ['1_3', 'Complementary Color 2', '補色 2', 'complementaryColor2'],
  ['1_4', 'Contrasting Color', 'カラー コントラスト', 'contrastingColor'],
  ['1_5', 'Darken', '暗く', 'darken'],
  ['1_6', 'Desaturate', '薄く', 'desaturate'],
  ['1_7', 'Lighten', '明るく', 'lighten'],
  ['1_8', 'Object Color', 'オブジェクト カラー', 'objectColor'],
  ['1_9', 'Pulse', 'パルス', 'pulse'],
  ['1_10', 'Underline', '下線', 'underline'],
  ['2_0', 'Color Pulse', 'カラー パルス', 'colorPulse'],
  ['2_1', 'Grow With Color', 'カラーで拡大', 'growWithColor'],
  ['2_2', 'Shimmer', 'シマー', 'shimmer'],
  ['2_3', 'Teeter', 'シーソー', 'teeter'],
  ['3_0', 'Blink', 'ブリンク', 'blink'],
  ['3_1', 'Bold Reveal', '太字表示', 'boldReveal'],
  ['3_2', 'Wave', 'ウェーブ', 'wave'],
]);

export const EXIT_TILES: readonly EffectTile[] = tiles('exit', [
  ['0_0', 'Blinds', 'ブラインド', 'blindsOut'],
  ['0_1', 'Checkerboard', 'チェッカーボード', 'checkerboardOut'],
  ['0_2', 'Disappear', 'クリア', 'disappear'],
  ['0_3', 'Dissolve Out', 'ディゾルブアウト', 'dissolveOut'],
  ['0_4', 'Fly Out', 'スライドアウト', 'flyOut'],
  ['0_5', 'Peek Out', 'ピークアウト', 'peekOut'],
  ['0_6', 'Random Bars', 'ランダムストライプ', 'randomBarsOut'],
  ['0_7', 'Shape', '図形', 'shapeOut'],
  ['0_8', 'Split', 'スプリット', 'splitOut'],
  ['0_9', 'Strips', 'ストリップ', 'stripsOut'],
  ['0_10', 'Wedge', 'くさび形', 'wedgeOut'],
  ['0_11', 'Wheel', 'ホイール', 'wheelOut'],
  ['0_12', 'Wipe', 'ワイプ', 'wipeOut'],
  ['1_0', 'Contract', 'コントラクト', 'contractOut'],
  ['1_1', 'Fade', 'フェード', 'fadeOut'],
  ['1_2', 'Swivel', 'ターン', 'swivelOut'],
  ['1_3', 'Zoom', 'ズーム', 'zoomOut'],
  ['2_0', 'Center Revolve', 'リボルブ', 'centerRevolveOut'],
  ['2_1', 'Collapse', 'コラプス', 'collapseOut'],
  ['2_2', 'Float Out', 'フロートアウト', 'floatOut'],
  ['2_3', 'Shrink & Turn', '縮小および回転', 'shrinkTurnOut'],
  ['2_4', 'Sink Down', 'シンク', 'sinkDownOut'],
  ['2_5', 'Spinner', 'スピナー', 'spinnerOut'],
  ['2_6', 'Basic Zoom', 'ベーシック ズーム', 'basicZoomOut'],
  ['2_7', 'Stretchy', 'ゴム', 'stretchyOut'],
  ['3_0', 'Boomerang', 'ブーメラン', 'boomerangOut'],
  ['3_1', 'Bounce', 'バウンド', 'bounceOut'],
  ['3_2', 'Credits', 'クレジット タイトル', 'creditsOut'],
  ['3_3', 'Curve Down', 'カーブ (下)', 'curveDownOut'],
  ['3_4', 'Drop', 'ドロップ', 'dropOut'],
  ['3_5', 'Flip', 'フリップ', 'flipOut'],
  ['3_6', 'Float', 'フロート', 'floatingOut'],
  ['3_7', 'Pinwheel', 'ピンウィール', 'pinwheelOut'],
  ['3_8', 'Spiral Out', 'スパイラルアウト', 'spiralOut'],
  ['3_9', 'Basic Swivel', 'ベーシック ターン', 'basicSwivelOut'],
  ['3_10', 'Whip', 'ホイップ', 'whipOut'],
]);

/** Every gallery's tiles, for lists that offer them all (the animation pane). */
export const ALL_TILES: readonly EffectTile[] = [...ENTRANCE_TILES, ...EMPHASIS_TILES, ...EXIT_TILES];

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

/** Effect Options ▸ Sequence, in PowerPoint's order. */
export const SEQUENCE: ReadonlyArray<{ readonly build: AnimationTextBuild; readonly en: string }> = [
  { build: 'asOneObject', en: 'As One Object' },
  { build: 'allAtOnce', en: 'All at Once' },
  { build: 'byParagraph', en: 'By Paragraph' },
];

// PowerPoint writes no build entry for Fill Color and Line Color: they colour
// the shape itself, so they have no Sequence to offer.
const UNBUILT: ReadonlySet<AnimationEffect> = new Set(['fillColor', 'lineColor']);

/** Whether the effect takes a Sequence (a text build). */
export const takesSequence = (effect: AnimationEffect | null): boolean =>
  effect !== null && !UNBUILT.has(effect);
