import type { AnimationDirection, AnimationEffect } from '@office-kit/pptx';

// Mac PowerPoint 16's Entrance and Emphasis galleries and the fly effects'
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
  ['Blinds', 'ブラインド'],
  ['Checkerboard', 'チェッカーボード'],
  ['Dissolve In', 'ディゾルブイン'],
  ['Fly In', 'スライドイン', 'flyIn'],
  ['Peek In', 'クロール イン'],
  ['Random Bars', 'ランダム ストライプ'],
  ['Shape', '図形'],
  ['Split', 'スプリット'],
  ['Strips', 'ストリップス'],
  ['Wedge', 'くさび形'],
  ['Wheel', 'ホイール'],
  ['Wipe', 'ワイプ'],
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

// The editor's exit presets; PowerPoint's Exit Effects gallery opens from its
// own ribbon button.
export const EXIT_TILES: readonly EffectTile[] = tiles('exit', [
  ['Disappear', 'クリア', 'disappear'],
  ['Fade Out', 'フェードアウト', 'fadeOut'],
  ['Fly Out', 'スライドアウト', 'flyOut'],
  ['Zoom Out', 'ズーム (終了)', 'zoomOut'],
]);

export const UNSUPPORTED_DIRECTION = 'Diagonal directions are not supported by the library yet.';

export interface DirectionOption {
  readonly en: string;
  readonly ja: string;
  readonly direction?: AnimationDirection;
  /** Arrow angle in degrees (0 points right). */
  readonly arrow: number;
}

// The arrow shows where the shape travels: "From Bottom" points up.
export const FLY_DIRECTIONS: readonly DirectionOption[] = [
  { en: 'From Bottom', ja: '下から', direction: 'bottom', arrow: 90 },
  { en: 'From Bottom-Left', ja: '左下から', arrow: 45 },
  { en: 'From Left', ja: '左から', direction: 'left', arrow: 0 },
  { en: 'From Top-Left', ja: '左上から', arrow: -45 },
  { en: 'From Top', ja: '上から', direction: 'top', arrow: -90 },
  { en: 'From Top-Right', ja: '右上から', arrow: -135 },
  { en: 'From Right', ja: '右から', direction: 'right', arrow: 180 },
  { en: 'From Bottom-Right', ja: '右下から', arrow: 135 },
];

export const UNSUPPORTED_SEQUENCE =
  'Animating all paragraphs at once is not supported by the library yet.';
