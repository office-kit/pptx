// Human-authored refinements over the generated capability manifest.
//
// The generator gives every capability a working default (a humanized label and
// a parsed parameter schema). This file upgrades the ones that benefit from a
// hand-tuned schema, a bilingual label, a ribbon home, or `primary` prominence.
// Entries here NEVER remove capabilities — coverage stays exhaustive regardless
// of how much of this file is filled in. Unlisted capabilities simply use their
// generated defaults, which are still fully reachable via the command palette.

import { generatedOverrides } from './overrides.generated.ts';
import type { CapabilityOverride, ParamSpec } from './types.ts';

// A run's own outline, shadow and glow — `<a:ln>` and `<a:effectLst>` inside
// `<a:rPr>`. They are part of `TextFormat` in the library; the generated
// schema predates them, so the fields are added here rather than by hand-
// writing the whole parameter again.
const CHARACTER_EFFECT_FIELDS: readonly ParamSpec[] = [
  {
    name: 'outline',
    type: 'TextOutline',
    kind: 'object',
    optional: true,
    label: 'Text outline',
    fields: [
      { name: 'color', type: 'string', kind: 'color', optional: true, label: 'Color' },
      { name: 'widthEmu', type: 'number', kind: 'emu', optional: true, label: 'Width' },
    ],
  },
  {
    name: 'shadow',
    type: 'ShadowOptions',
    kind: 'object',
    optional: true,
    label: 'Text shadow',
    fields: [
      { name: 'color', type: 'string', kind: 'color', optional: true, label: 'Color' },
      { name: 'blurEmu', type: 'number', kind: 'emu', optional: true, label: 'Blur' },
      { name: 'offsetEmu', type: 'number', kind: 'emu', optional: true, label: 'Distance' },
      { name: 'angleDeg', type: 'number', kind: 'number', optional: true, label: 'Angle' },
      { name: 'opacity', type: 'number', kind: 'number', optional: true, label: 'Opacity' },
    ],
  },
  {
    name: 'glow',
    type: 'GlowOptions',
    kind: 'object',
    optional: true,
    label: 'Text glow',
    fields: [
      { name: 'color', type: 'string', kind: 'color', optional: false, label: 'Color' },
      { name: 'radiusEmu', type: 'number', kind: 'emu', optional: true, label: 'Radius' },
      { name: 'opacity', type: 'number', kind: 'number', optional: true, label: 'Opacity' },
    ],
  },
];

const withCharacterEffects = (format: ParamSpec): ParamSpec => ({
  ...format,
  fields: [...(format.fields ?? []), ...CHARACTER_EFFECT_FIELDS],
});

// Hand-authored refinements. Merged on top of `generatedOverrides` (the
// workflow-enriched field schemas), so a hand entry wins for the same id.
const handOverrides: Record<string, CapabilityOverride> = {
  addSlidePlaceholder: {
    labelEn: 'Add placeholder',
    labelJa: 'プレースホルダーを追加',
    params: [
      {
        name: 'type',
        type: 'string',
        kind: 'enum',
        optional: false,
        label: 'Placeholder',
        // `ST_PlaceholderType`. Only slots a layout actually reserves can be
        // added, so an unused token here simply reports nothing to add.
        enumValues: ['title', 'body', 'sldNum', 'dt', 'ftr', 'pic', 'chart', 'tbl'],
        default: 'sldNum',
      },
    ],
  },
  setShapeTextField: {
    labelEn: 'Insert field',
    labelJa: 'フィールドを挿入',
    ribbonGroup: 'text',
    params: [
      {
        name: 'type',
        type: 'string',
        kind: 'enum',
        optional: false,
        label: 'Field',
        // `ST_TextFieldType` is an open string; these are the tokens
        // PowerPoint itself writes and every reader substitutes. The thirteen
        // `datetime` variants differ only in format, so the dialog offers the
        // two PowerPoint's own Insert menu does.
        enumValues: ['slidenum', 'datetime1', 'datetime2', 'footer', 'headerfooter'],
        default: 'slidenum',
      },
      {
        name: 'options',
        type: '{ text?: string }',
        kind: 'object',
        optional: true,
        label: 'Options',
        fields: [
          {
            name: 'text',
            type: 'string',
            kind: 'string',
            optional: true,
            label: 'Cached text',
          },
        ],
      },
    ],
  },
  setShapeTextFormat: {
    labelEn: 'Set Shape Text Format',
    labelJa: '図形テキストの書式を設定',
    params: [
      withCharacterEffects(generatedOverrides.setShapeTextFormat!.params![0]!),
      ...generatedOverrides.setShapeTextFormat!.params!.slice(1),
      {
        name: 'options',
        type: '{ range?: { start: number; end: number }; reset?: boolean }',
        kind: 'object',
        optional: true,
        label: 'Options',
        fields: [
          {
            name: 'reset',
            type: 'boolean',
            kind: 'boolean',
            optional: true,
            label: 'Clear text formatting',
          },
          {
            name: 'range',
            type: '{ start: number; end: number }',
            kind: 'object',
            optional: true,
            label: 'Text range',
            fields: [
              { name: 'start', type: 'number', kind: 'number', optional: false, label: 'Start' },
              { name: 'end', type: 'number', kind: 'number', optional: false, label: 'End' },
            ],
          },
        ],
      },
    ],
  },
  setCommentText: { labelEn: 'Edit Comments', labelJa: 'コメントを編集' },
  setCommentStatus: {
    labelEn: 'Resolve or Reopen Comment',
    labelJa: 'コメントを解決 / 再開',
    params: [
      {
        name: 'comment',
        type: 'SlideCommentData',
        kind: 'object',
        optional: false,
        label: 'Comment',
      },
      {
        name: 'status',
        type: 'CommentStatus',
        kind: 'enum',
        optional: false,
        label: 'Status',
        enumValues: ['active', 'resolved', 'closed'],
      },
    ],
  },
  splitTableCell: { labelJa: '結合セルを分割' },
  addSectionHeaderSlide: { labelJa: 'セクション見出しスライドの追加' },
  addMissingSlidePlaceholders: {
    labelEn: 'Restore deleted placeholders',
    labelJa: '削除したプレースホルダーを復元',
    category: 'slide',
  },
  resetSlidePlaceholderTextFormatting: {
    labelEn: 'Reset placeholder text formatting',
    labelJa: 'プレースホルダーの文字書式を戻す',
    category: 'slide',
  },
  resetSlideLayout: {
    labelEn: 'Reset layout',
    labelJa: 'レイアウトをリセット',
    category: 'slide',
  },
  resetSlidePlaceholderGeometry: {
    labelEn: 'Reset placeholder positions',
    labelJa: 'プレースホルダーの配置を戻す',
  },
  addSlide: { labelEn: 'New slide from layout', labelJa: 'レイアウトからスライドを追加' },
  addSlideAt: { labelJa: '指定位置にスライドを追加' },
  addSlideChart: { labelJa: 'グラフの挿入' },
  addSlideImage: { labelJa: '画像の挿入' },
  addSlideLine: { labelJa: '線の挿入' },
  addSlideMedia: { labelJa: '音声・動画の挿入' },
  addSlideShape: { labelJa: '図形の挿入' },
  addSlideTable: { labelJa: '表の挿入' },
  addSlideTextBox: { labelJa: 'テキストボックスの挿入' },
  appendShapeText: { labelJa: 'テキストの追記' },
  appendSlideNotes: { labelJa: 'スピーカーノートの追記' },
  clearAllHyperlinks: { labelJa: 'すべてのリンクを削除' },
  clearAllSlideComments: { labelJa: 'すべてのスライドのコメントを削除' },
  clearAllSlideNotes: { labelJa: 'すべてのスピーカーノートを削除' },
  clearShapeEffects: { labelJa: '図形の効果を解除' },
  clearShapeFill: { labelJa: '図形の塗りつぶしを解除' },
  clearShapeStroke: { labelJa: '図形の枠線を解除' },
  clearSlideAnimations: { labelJa: 'アニメーションを解除' },
  clearSlideBackground: { labelJa: '背景を解除' },
  clearSlideComments: { labelJa: 'スライドのコメントを削除' },
  clearSlideHyperlinks: { labelJa: 'スライドのリンクを削除' },
  clearSlideShapes: { labelJa: 'スライド内の図形をすべて削除' },
  clearSlideTransition: { labelJa: '画面切り替えを解除' },
  clearTableCellFill: { labelJa: 'セルの塗りつぶしを解除' },
  compactPackage: { labelJa: '未使用データを削除' },
  copyShape: { labelJa: '図形のコピー' },
  createPresentation: { labelJa: 'プレゼンテーションの作成' },
  duplicateSlideAt: { labelJa: '指定位置にスライドを複製' },
  importSlide: { labelJa: 'スライドのインポート' },
  incrementRevision: { labelJa: 'リビジョン番号の更新' },
  insertTableColumn: { labelJa: '列の挿入' },
  mergePresentations: { labelJa: 'プレゼンテーションの結合' },
  removeSlideNotes: { labelJa: 'スピーカーノートの削除' },
  removeTableColumn: { labelJa: '列の削除' },
  removeTableRow: { labelJa: '行の削除' },
  removeThumbnail: { labelJa: 'サムネイルの削除' },
  renameShape: { labelJa: '図形名の変更' },
  replaceHyperlink: { labelJa: 'リンク先の置換' },
  replaceTextInNotes: { labelJa: 'ノート内のテキストを置換' },
  replaceTextInPresentation: { labelJa: 'プレゼンテーション全体のテキストを置換' },
  replaceTextInSlide: { labelJa: 'スライド内のテキストを置換' },
  replaceTextInSlideNotes: { labelJa: 'スライドのノート内のテキストを置換' },
  reverseSlides: { labelJa: 'スライドの順序を反転' },
  setChartSpec: { labelJa: 'グラフのデータと設定' },
  setCoreProperties: { labelJa: '文書の基本情報' },
  setExtendedProperties: { labelJa: '文書の詳細情報' },
  setMediaPartBytes: { labelJa: 'メディアデータの置換' },
  setParagraphLevel: { labelJa: '段落の階層' },
  setShapeAltTitle: { labelJa: '代替テキストのタイトル' },
  setShapeDescription: { labelJa: '代替テキストの説明' },
  setShapeHidden: { labelJa: '図形の表示・非表示' },
  setShapeHyperlink: { labelJa: '図形のリンク' },
  setShapeImage: { labelJa: '画像の置換' },
  setShapeImageBrightness: { labelJa: '画像の明るさ' },
  setShapeImageContrast: { labelJa: '画像のコントラスト' },
  setShapeImageFill: { labelJa: '画像による塗りつぶし' },
  setShapeImageOpacity: { labelJa: '画像の不透明度' },
  setShapeParagraphs: { labelJa: '段落と文字書式の編集' },
  setShapePosition: { labelJa: '図形の位置' },
  setShapeRunHyperlink: { labelJa: '文字列のリンク' },
  setShapeRunText: { labelJa: '文字列の編集' },
  setShapeSize: { labelJa: '図形のサイズ' },
  setShapeStrokeCap: { labelJa: '線の端の形状' },
  setShapeStrokeCompound: { labelJa: '線の種類' },
  setShapeStrokeJoin: { labelJa: '線の接合部の形状' },
  setShapeTextBodyRotationDeg: { labelJa: 'テキストの回転' },
  setShapeZIndex: { labelJa: '図形の重なり順' },
  setSlideBackground: { labelJa: '背景色' },
  setSlideBackgroundImage: { labelJa: '背景画像' },
  setSlideBody: { labelJa: 'スライド本文の編集' },
  setSlideLayout: { labelJa: 'スライドのレイアウト' },
  setSlideNotes: { labelJa: 'スピーカーノートの編集' },
  setSlideTitle: { labelJa: 'スライドタイトルの編集' },
  setTableCellClickAction: { labelEn: 'Set Table Cell Link', labelJa: '表セルのリンクを設定' },
  setTableCellAnchor: { labelJa: 'セルの縦方向の配置' },
  setTableCellFill: { labelJa: 'セルの塗りつぶし' },
  setTableCellParagraphs: { labelJa: 'セルの段落と文字書式の編集' },
  setTableCellText: { labelJa: 'セルのテキスト編集' },
  setTableColumnWidth: { labelJa: '列の幅' },
  setTableRowHeight: { labelJa: '行の高さ' },
  setTableStyleId: { labelJa: '表のスタイル' },
  setThumbnail: { labelJa: 'サムネイルの設定' },
  sortSlides: { labelJa: 'スライドの並べ替え' },
  swapSlides: { labelJa: 'スライドの入れ替え' },
  touchModified: { labelJa: '最終更新日時の更新' },
  ungroupShapes: { labelJa: 'グループ解除' },
  // --- Slides ------------------------------------------------------------
  addBlankSlide: {
    labelEn: 'Blank slide',
    labelJa: '白紙のスライド',
    ribbonGroup: 'slides',
    primary: true,
    params: [],
  },
  addTitleSlide: {
    labelEn: 'Title slide',
    labelJa: 'タイトルスライド',
    ribbonGroup: 'slides',
    primary: true,
    params: [{ name: 'title', type: 'string', kind: 'string', optional: false, label: 'Title' }],
  },
  addContentSlide: {
    labelEn: 'Title + content slide',
    labelJa: 'タイトルとコンテンツ',
    ribbonGroup: 'slides',
    primary: true,
    params: [
      { name: 'title', type: 'string', kind: 'string', optional: true, label: 'Title' },
      { name: 'body', type: 'string', kind: 'string', optional: true, label: 'Body' },
    ],
  },
  duplicateSlide: {
    labelEn: 'Duplicate slide',
    labelJa: 'スライドの複製',
    ribbonGroup: 'slides',
    primary: true,
    params: [],
  },
  removeSlide: {
    labelEn: 'Delete slide',
    labelJa: 'スライドの削除',
    ribbonGroup: 'slides',
    primary: true,
    params: [],
  },
  moveSlide: { labelEn: 'Move slide', labelJa: 'スライドの移動', ribbonGroup: 'slides' },
  setSlideHidden: {
    labelEn: 'Hide slide',
    labelJa: 'スライドを非表示',
    ribbonGroup: 'slides',
    params: [
      { name: 'hidden', type: 'boolean', kind: 'boolean', optional: false, label: 'Hidden' },
    ],
  },
  setSlideSize: { labelEn: 'Slide size', labelJa: 'スライドのサイズ', ribbonGroup: 'design' },

  // --- Shapes ------------------------------------------------------------
  setShapeFill: {
    labelEn: 'Shape fill',
    labelJa: '図形の塗りつぶし',
    ribbonGroup: 'shape-styles',
    primary: true,
    params: [{ name: 'color', type: 'string', kind: 'color', optional: false, label: 'Color' }],
  },
  setShapeStroke: {
    labelEn: 'Shape outline',
    labelJa: '図形の枠線',
    ribbonGroup: 'shape-styles',
    primary: true,
    // The library takes a single `options` object — flattening it would pass
    // the wrong positional args. Render it as a field group.
    params: [
      {
        name: 'options',
        type: '{ color?: string; widthEmu?: number }',
        kind: 'object',
        optional: false,
        fields: [
          { name: 'color', type: 'string', kind: 'color', optional: true, label: 'Color' },
          { name: 'widthEmu', type: 'number', kind: 'emu', optional: true, label: 'Width' },
        ],
      },
    ],
  },
  setShapeNoFill: {
    labelEn: 'No fill',
    labelJa: '塗りつぶしなし',
    ribbonGroup: 'shape-styles',
    params: [],
  },
  setShapeNoStroke: {
    labelEn: 'No outline',
    labelJa: '枠線なし',
    ribbonGroup: 'shape-styles',
    params: [],
  },
  removeShape: {
    labelEn: 'Delete shape',
    labelJa: '図形の削除',
    ribbonGroup: 'arrange',
    primary: true,
    params: [],
  },
  bringShapeToFront: {
    labelEn: 'Bring to front',
    labelJa: '最前面へ',
    ribbonGroup: 'arrange',
    params: [],
  },
  sendShapeToBack: {
    labelEn: 'Send to back',
    labelJa: '最背面へ',
    ribbonGroup: 'arrange',
    params: [],
  },
  bringShapeForward: {
    labelEn: 'Bring forward',
    labelJa: '前面へ',
    ribbonGroup: 'arrange',
    params: [],
  },
  sendShapeBackward: {
    labelEn: 'Send backward',
    labelJa: '背面へ',
    ribbonGroup: 'arrange',
    params: [],
  },
  setShapePreset: {
    labelEn: 'Change shape',
    labelJa: '形状の変更',
    params: [
      {
        name: 'preset',
        type: 'PresetShape',
        kind: 'enum',
        optional: false,
        label: 'Shape',
        enumValues: [
          'rect',
          'roundRect',
          'ellipse',
          'triangle',
          'diamond',
          'pentagon',
          'hexagon',
          'star5',
          'heart',
        ],
        default: 'rect',
      },
    ],
  },
  setShapeRotation: {
    labelEn: 'Rotation',
    labelJa: '回転',
    ribbonGroup: 'arrange',
    params: [
      {
        name: 'degrees',
        type: 'number',
        kind: 'number',
        optional: false,
        label: 'Degrees',
        default: '0',
      },
    ],
  },

  // --- Effects / advanced fills (field-based nested dialogs) --------------
  setShapeGradientFill: {
    labelEn: 'Gradient fill',
    labelJa: 'グラデーション',
    ribbonGroup: 'shape-styles',
    params: [
      {
        name: 'options',
        type: 'GradientFillOptions',
        kind: 'object',
        optional: false,
        fields: [
          {
            name: 'stops',
            type: 'GradientStop[]',
            kind: 'array',
            optional: false,
            label: 'Color stops',
            item: {
              name: 'stop',
              type: 'GradientStop',
              kind: 'object',
              optional: false,
              fields: [
                {
                  name: 'offset',
                  type: 'number',
                  kind: 'number',
                  optional: false,
                  label: 'Offset (0–1)',
                },
                { name: 'color', type: 'string', kind: 'color', optional: false, label: 'Color' },
              ],
            },
          },
          {
            name: 'angleDeg',
            type: 'number',
            kind: 'number',
            optional: true,
            label: 'Angle (°)',
            default: '90',
          },
          {
            name: 'path',
            type: "'linear'|'circle'|'rect'|'shape'",
            kind: 'enum',
            optional: true,
            label: 'Path',
            enumValues: ['linear', 'circle', 'rect', 'shape'],
          },
        ],
      },
    ],
  },
  setShapePatternFill: {
    labelEn: 'Pattern fill',
    labelJa: 'パターン',
    ribbonGroup: 'shape-styles',
    params: [
      {
        name: 'options',
        type: 'PatternFillOptions',
        kind: 'object',
        optional: false,
        fields: [
          {
            name: 'preset',
            type: 'PatternPreset',
            kind: 'string',
            optional: false,
            label: 'Preset (e.g. pct50, dkUpDiag, wave)',
          },
          {
            name: 'foreground',
            type: 'string',
            kind: 'color',
            optional: false,
            label: 'Foreground',
          },
          {
            name: 'background',
            type: 'string',
            kind: 'color',
            optional: false,
            label: 'Background',
          },
        ],
      },
    ],
  },
  setShapeShadow: {
    labelEn: 'Shadow',
    labelJa: '影',
    ribbonGroup: 'effects',
    params: [
      {
        name: 'options',
        type: 'ShadowOptions',
        kind: 'object',
        optional: true,
        fields: [
          { name: 'color', type: 'string', kind: 'color', optional: true, label: 'Color' },
          { name: 'blurEmu', type: 'Emu', kind: 'emu', optional: true, label: 'Blur' },
          { name: 'offsetEmu', type: 'Emu', kind: 'emu', optional: true, label: 'Offset' },
          {
            name: 'angleDeg',
            type: 'number',
            kind: 'number',
            optional: true,
            label: 'Angle (°)',
            default: '45',
          },
          {
            name: 'opacity',
            type: 'number',
            kind: 'number',
            optional: true,
            label: 'Opacity (0–1)',
          },
        ],
      },
    ],
  },
  setShapeGlow: {
    labelEn: 'Glow',
    labelJa: '光彩',
    ribbonGroup: 'effects',
    params: [
      {
        name: 'options',
        type: 'GlowOptions',
        kind: 'object',
        optional: false,
        fields: [
          { name: 'color', type: 'string', kind: 'color', optional: false, label: 'Color' },
          { name: 'radiusEmu', type: 'Emu', kind: 'emu', optional: true, label: 'Radius' },
        ],
      },
    ],
  },
  setSlideTransition: {
    labelEn: 'Transition',
    labelJa: '画面切り替え',
    ribbonGroup: 'transition',
    primary: true,
    params: [
      {
        name: 'options',
        type: 'TransitionOptions',
        kind: 'object',
        optional: false,
        fields: [
          {
            name: 'effect',
            type: 'TransitionEffect',
            kind: 'enum',
            optional: false,
            label: 'Effect',
            enumValues: [
              'none',
              'fade',
              'push',
              'cover',
              'wipe',
              'split',
              'cut',
              'dissolve',
              'checker',
              'blinds',
              'randomBar',
              'zoom',
              'circle',
              'diamond',
              'plus',
              'wheel',
              'wedge',
              'newsflash',
            ],
          },
          {
            name: 'speed',
            type: "'slow'|'med'|'fast'",
            kind: 'enum',
            optional: true,
            label: 'Speed',
            enumValues: ['slow', 'med', 'fast'],
          },
          {
            name: 'direction',
            type: 'string',
            kind: 'string',
            optional: true,
            label: 'Direction (effect-specific, e.g. l/r/u/d)',
          },
          { name: 'spokes', type: 'number', kind: 'number', optional: true, label: 'Wheel spokes' },
          {
            name: 'thruBlack',
            type: 'boolean',
            kind: 'boolean',
            optional: true,
            label: 'Through black (fade)',
          },
        ],
      },
    ],
  },
  setShapeAnimation: {
    labelEn: 'Animation',
    labelJa: 'アニメーション',
    ribbonGroup: 'animation',
    primary: true,
    params: [
      {
        name: 'opts',
        type: 'AnimationOptions',
        kind: 'object',
        optional: false,
        fields: [
          {
            name: 'effect',
            type: 'AnimationEffect',
            kind: 'enum',
            optional: false,
            label: 'Effect',
            enumValues: ['fadeIn', 'fadeOut', 'appear', 'disappear'],
          },
          {
            name: 'durationMs',
            type: 'number',
            kind: 'number',
            optional: true,
            label: 'Duration (ms)',
            default: '500',
          },
        ],
      },
    ],
  },
  updateSlideAnimation: {
    labelEn: 'Edit animation',
    labelJa: 'アニメーションの変更',
    ribbonGroup: 'animation',
    params: [
      { name: 'id', type: 'number', kind: 'number', optional: false, label: 'Animation' },
      {
        name: 'patch',
        type: 'AnimationPatch',
        kind: 'object',
        optional: false,
        fields: [
          {
            name: 'effect',
            type: 'AnimationEffect',
            kind: 'enum',
            optional: true,
            label: 'Effect',
            enumValues: ['fadeIn', 'fadeOut', 'appear', 'disappear'],
          },
          {
            name: 'start',
            type: 'AnimationStartCondition',
            kind: 'enum',
            optional: true,
            label: 'Start',
            enumValues: ['click', 'withPrevious', 'afterPrevious'],
          },
          {
            name: 'durationMs',
            type: 'number',
            kind: 'number',
            optional: true,
            label: 'Duration (ms)',
          },
          { name: 'delayMs', type: 'number', kind: 'number', optional: true, label: 'Delay (ms)' },
          {
            name: 'byParagraph',
            type: 'boolean',
            kind: 'boolean',
            optional: true,
            label: 'By paragraph',
          },
        ],
      },
    ],
  },
  removeSlideAnimation: {
    labelEn: 'Delete animation',
    labelJa: 'アニメーションの削除',
    ribbonGroup: 'animation',
    params: [{ name: 'id', type: 'number', kind: 'number', optional: false, label: 'Animation' }],
  },
  moveSlideAnimation: {
    labelEn: 'Reorder animation',
    labelJa: 'アニメーションの並べ替え',
    ribbonGroup: 'animation',
    params: [
      { name: 'id', type: 'number', kind: 'number', optional: false, label: 'Animation' },
      { name: 'index', type: 'number', kind: 'number', optional: false, label: 'Position' },
    ],
  },
  setShapeImageCrop: {
    labelEn: 'Crop image',
    labelJa: '画像のトリミング',
    ribbonGroup: 'picture',
    params: [
      {
        name: 'crop',
        type: 'ImageCrop',
        kind: 'object',
        optional: false,
        fields: [
          { name: 'left', type: 'number', kind: 'number', optional: true, label: 'Left (0–1)' },
          { name: 'top', type: 'number', kind: 'number', optional: true, label: 'Top (0–1)' },
          { name: 'right', type: 'number', kind: 'number', optional: true, label: 'Right (0–1)' },
          { name: 'bottom', type: 'number', kind: 'number', optional: true, label: 'Bottom (0–1)' },
        ],
      },
    ],
  },
};

// The effective override map: workflow-enriched field schemas as the base,
// hand-authored refinements layered on top (hand wins per id).
export const overrides: Record<string, CapabilityOverride> = {
  ...generatedOverrides,
  ...handOverrides,
};
