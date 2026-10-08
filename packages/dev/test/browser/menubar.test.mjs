import assert from 'node:assert/strict';
import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { chromium } from 'playwright';
import {
  getParagraphPropertiesEffective,
  getShapeKind,
  getSlideShapes,
  getSlides,
  loadPresentation,
} from '@office-kit/pptx';
import { startPreview, waitForState } from '../helpers/server.mjs';
import { installRichTextSelection } from '../helpers/rich-text.mjs';

// The reference desktop app's (Mac, 16.113.3) menu bar (AX capture of every menu with nothing
// selected, English and Japanese UI, 2026-10-07): order, separators, submenus
// and shortcut glyphs. The application menu is the browser's and is left
// out; so are the lists that name the capture machine's state (recent files,
// subtitle languages, microphones). Undo and Repeat name the last edit, and
// the Window menu lists this deck, as the reference desktop app lists its documents.
const NATIVE = {
  en: {
    File: [
      'New Presentation ⌘N',
      'New from Template... ⇧⌘P',
      'Open... ⌘O',
      'Open Recent ▸',
      '----',
      'Close ⌘W',
      'Save ⌘S',
      'Save As... ⇧⌘S',
      'Save as Template...',
      'Export...',
      'Move...',
      'Rename...',
      '----',
      'Browse Version History',
      '----',
      'Share ▸',
      '----',
      'Always Open Read-Only',
      'Restrict Permissions ▸',
      'Passwords...',
      '----',
      'Compress Pictures...',
      '----',
      'Page Setup...',
      'Print... ⌘P',
      '----',
      'Properties',
    ],
    'File ▸ Open Recent': ['More... ⇧⌘O'],
    'File ▸ Share': ['Share...', '----', 'Send Presentation', 'Send PDF'],
    Edit: [
      "Can't Undo ⌘Z",
      "Can't Repeat ⌘Y",
      '----',
      'Cut ⌘X',
      'Copy ⌘C',
      '----',
      'Paste ⌘V',
      'Paste Special... ⌃⌘V',
      'Paste and Match Formatting ⌥⇧⌘V',
      '----',
      'Clear',
      'Select All ⌘A',
      'Duplicate ⌘D',
      'Delete Slide',
      '----',
      'Remove Section',
      'Rename Section',
      '----',
      'Find ▸',
      'Select Data...',
      '----',
      'Toggle Drawing ⌃⌘Z',
      '----',
      'AutoFill ▸',
      'Start Dictation 🎤',
      'Emoji & Symbols 🌐',
    ],
    'Edit ▸ Find': [
      'Find... ⌘F',
      'Find Next ⌘G',
      'Find Previous ⇧⌘G',
      'Advanced Find... ⌃F',
      '----',
      'Replace... ⌃H',
      'Replace Fonts...',
    ],
    'Edit ▸ AutoFill': ['Contact…', 'Passwords…', 'Credit Card…'],
    View: [
      'Normal ⌘1',
      'Slide Sorter ⌘2',
      'Notes Page ⌘3',
      'Outline View ⌘4',
      'Reading View ⌘5',
      'Presenter View ⌥↩',
      'Slide Show ⇧⌘↩',
      '----',
      'Show Slides ⌃⇧⇥',
      '----',
      'Master ▸',
      '----',
      'Ribbon ⌥⌘R',
      'Message Bar',
      '----',
      'Header and Footer...',
      'Markup',
      'Advanced Markup',
      '----',
      'Ruler',
      'Grid and Guides ▸',
      'Zoom ▸',
      'Enter Full Screen 🌐F',
    ],
    'View ▸ Master': ['Slide Master ⌥⌘1', 'Handout Master ⌥⌘2', 'Notes Master ⌥⌘3'],
    'View ▸ Grid and Guides': [
      'Smart Guides',
      'Guides ⌃⌥⌘G',
      'Gridlines ⌘^',
      '----',
      'Snap to Grid',
      '----',
      'Grid Options...',
    ],
    'View ▸ Zoom': ['Fit to Window ⌥⌘O', 'Zoom In ⌘+', 'Zoom Out ⌘-', '----', 'Zoom...'],
    Insert: [
      'New Slide ⇧⌘N',
      'Duplicate Slide ⇧⌘D',
      'Slides From ▸',
      '----',
      'Section',
      'Comment ⇧⌘M',
      '----',
      'WordArt',
      'Header and Footer...',
      'Date and Time...',
      'Slide Number',
      '----',
      'Table...',
      'Chart ▸',
      'SmartArt ▸',
      '----',
      'Picture ▸',
      'Audio ▸',
      'Video ▸',
      'Cameo',
      'Equation',
      'Symbol...',
      'Shape ▸',
      '----',
      'Icons...',
      '----',
      '3D Models ▸',
      '----',
      'Zoom ▸',
      '----',
      'Action Buttons ▸',
      'Action Settings...',
      '----',
      'Object...',
      'Hyperlink... ⌘K',
    ],
    'Insert ▸ Slides From': ['Reuse Slides...', 'Outline...'],
    'Insert ▸ Chart': [
      'Column',
      'Bar',
      'Line',
      'Area',
      'Pie',
      'Treemap',
      'Sunburst',
      'Histogram',
      'Pareto',
      'Box and Whisker',
      'X Y (Scatter)',
      'Waterfall',
      'Funnel',
      'Stock',
      'Surface',
      'Radar',
      'Filled Map',
    ],
    'Insert ▸ SmartArt': [
      'List',
      'Process',
      'Cycle',
      'Hierarchy',
      'Relationship',
      'Matrix',
      'Pyramid',
      'Picture',
      'Timeline',
      'Meet The Team',
      'Text Card',
    ],
    'Insert ▸ Picture': [
      'Photo Browser…',
      'Picture from File...',
      'Stock Images...',
      'Online Pictures...',
      'Brand Images...',
    ],
    'Insert ▸ Audio': ['Audio Browser...', 'Audio from File...', '----', 'Record Audio...'],
    'Insert ▸ Video': [
      'Movie Browser...',
      'Movie from File...',
      'Stock Videos...',
      'Online Movie...',
    ],
    'Insert ▸ Shape': ['Rectangle', 'Rounded Rectangle', 'Triangle', 'Oval', 'Line'],
    'Insert ▸ 3D Models': ['Insert 3D Model From', 'Stock 3D Models...', 'This Device...'],
    'Insert ▸ Zoom': ['Summary Zoom', 'Section Zoom', 'Slide Zoom'],
    'Insert ▸ Action Buttons': [
      'Action Buttons',
      'Back or Previous',
      'Forward or Next',
      'Beginning',
      'End',
      'Home',
      'Information',
      'Return',
      'Movie',
      'Document',
      'Sound',
      'Help',
      'Custom',
    ],
    Format: [
      'Font... ⌘T',
      'Paragraph... ⌥⌘M',
      '----',
      'Bullets and Numbering...',
      '----',
      'Columns...',
      'Alignment ▸',
      'More Options...',
      '----',
      'Pick Up Object Style ⇧⌘C',
      'Apply To Defaults ⇧⌘V',
      'Animation Painter ⌥⇧⌘C',
      '----',
      'Replace Fonts...',
      '----',
      'Theme Colors...',
      'Slide Background... ⇧⌘2',
      '----',
      'Crop ⇧C',
      '----',
      'Format Object... ⇧⌘1',
    ],
    'Format ▸ Alignment': [
      'Align Left ⌘L',
      'Center ⌘E',
      'Align Right ⌘R',
      'Justify',
      'Distributed',
    ],
    Arrange: [
      'Reorder Objects',
      'Reorder Overlapping Objects',
      '----',
      'Bring to Front ⇧⌘F',
      'Send to Back ⇧⌘B',
      'Bring Forward ⌥⇧⌘F',
      'Send Backward ⌥⇧⌘B',
      '----',
      'Group ⌥⌘G',
      'Ungroup ⌥⇧⌘G',
      'Regroup ⌥⌘J',
      '----',
      'Rotate or Flip ▸',
      'Align or Distribute ▸',
      '----',
      'Selection Pane... ⌥⌘U',
    ],
    'Arrange ▸ Rotate or Flip': [
      'Rotate Left 90°',
      'Rotate Right 90°',
      'Flip Horizontal',
      'Flip Vertical',
      '----',
      'More Rotation Options...',
    ],
    'Arrange ▸ Align or Distribute': [
      'Align Left',
      'Align Center',
      'Align Right',
      'Align Top',
      'Align Middle',
      'Align Bottom',
      '----',
      'Distribute Horizontally',
      'Distribute Vertically',
      '----',
      'Align to Slide',
      'Align Selected Objects',
    ],
    Tools: [
      'Spelling...',
      'Thesaurus... ⌃⌥⌘R',
      'Translate...',
      'Set Proofing Language...',
      'AutoCorrect Options...',
      '----',
      'Check Accessibility',
      '----',
      'Macro ▸',
      'Add-ins...',
    ],
    'Tools ▸ Macro': ['Macros...', 'Visual Basic Editor'],
    'Slide Show': [
      'Play from Start ⇧⌘↩',
      'Play from Current Slide ⌘↩',
      '----',
      'Custom Slide Show ▸',
      '----',
      'Rehearse with Coach',
      '----',
      'Presenter View ⌥↩',
      'Rehearse Timings',
      'Record Slide Show',
      '----',
      'Hide Slide',
      'Set Up Show...',
      '----',
      'Always Use Subtitles',
      'Subtitle Settings ▸',
    ],
    'Slide Show ▸ Subtitle Settings': [
      'Spoken Language ▸',
      'Subtitle Language ▸',
      '----',
      'Microphone ▸',
      '----',
      'Bottom (Overlaid)',
      'Top (Overlaid)',
      'Below Slide',
      'Above Slide',
      '----',
      'System Caption Preferences...',
    ],
    Window: [
      'Minimize ⌘M',
      'Minimize All',
      'Zoom',
      'Zoom All',
      'Fill 🌐⌃F',
      'Center 🌐⌃C',
      '----',
      'Move & Resize ▸',
      'Full Screen Tile',
      '----',
      'Remove Window from Set',
      '----',
      'New Window',
      'Arrange All',
      'Next Pane',
      '----',
      'Bring All to Front',
      'Arrange in Front',
      '----',
      'deck',
    ],
    'Window ▸ Move & Resize': [
      'Halves',
      'Left',
      'Right',
      'Top',
      'Bottom',
      '----',
      'Quarters',
      'Top Left',
      'Top Right',
      'Bottom Left',
      'Bottom Right',
      '----',
      'Arrange',
      'Left & Right',
      'Left & Quarters',
      'Right & Left',
      'Right & Quarters',
      'Top & Bottom',
      'Top & Quarters',
      'Bottom & Top',
      'Bottom & Quarters',
      'Quarters',
      '----',
      'Return to Previous Size 🌐⌃R',
    ],
    Help: [
      'Editor Help ⌘?',
      '----',
      'Feedback',
      '----',
      'Clear Application Data',
      '----',
      'Check for Updates',
    ],
  },
  ja: {
    ファイル: [
      '新規作成 ⌘N',
      'テンプレートから新規作成... ⇧⌘P',
      '開く... ⌘O',
      '最近使ったファイル ▸',
      '----',
      '閉じる ⌘W',
      '保存 ⌘S',
      '名前を付けて保存... ⇧⌘S',
      'テンプレートとして保存...',
      'エクスポート...',
      '移動...',
      '名前の変更...',
      '----',
      'バージョン履歴の表示',
      '----',
      '共有 ▸',
      '----',
      '常に読み取り専用で開く',
      'アクセスの制限 ▸',
      'パスワード...',
      '----',
      '図の圧縮...',
      '----',
      'ページ設定...',
      'プリント... ⌘P',
      '----',
      'プロパティ',
    ],
    'ファイル ▸ 最近使ったファイル': ['その他... ⇧⌘O'],
    'ファイル ▸ 共有': ['共有...', '----', 'プレゼンテーションを送信', 'PDF を送信'],
    編集: [
      '元に戻せません ⌘Z',
      '繰り返しできません ⌘Y',
      '----',
      'カット ⌘X',
      'コピー ⌘C',
      '----',
      'ペースト ⌘V',
      '形式を選択してペースト... ⌃⌘V',
      'ペーストしてスタイルを合わせる ⌥⇧⌘V',
      '----',
      'クリア',
      'すべてを選択 ⌘A',
      '複製 ⌘D',
      'スライドの削除',
      '----',
      'セクションの削除',
      'セクション名の変更',
      '----',
      '検索 ▸',
      'データの選択...',
      '----',
      '描画の切り替え ⌃⌘Z',
      '----',
      '自動入力 ▸',
      '音声入力を開始 🎤',
      '絵文字と記号 🌐',
    ],
    '編集 ▸ 検索': [
      '検索... ⌘F',
      '次を検索 ⌘G',
      '前を検索 ⇧⌘G',
      '高度な検索... ⌃F',
      '----',
      '置換... ⌃H',
      'フォントの置換...',
    ],
    '編集 ▸ 自動入力': ['連絡先…', 'パスワード…', 'クレジットカード…'],
    表示: [
      '標準 ⌘1',
      'スライド一覧 ⌘2',
      'ノート ⌘3',
      'アウトライン表示 ⌘4',
      '閲覧表示 ⌘5',
      '発表者ツール ⌥↩',
      'スライド ショー ⇧⌘↩',
      '----',
      'アウトラインの表示 ⌃⇧⇥',
      '----',
      'マスター ▸',
      '----',
      'リボン ⌥⌘R',
      'メッセージ バー',
      '----',
      'ヘッダーとフッター...',
      'コメントと注釈',
      '高度なマークアップ',
      '----',
      'ルーラー',
      'グリッドとガイド ▸',
      'ズーム ▸',
      '全画面表示 🌐F',
    ],
    '表示 ▸ マスター': ['スライド マスター ⌥⌘1', '配布資料マスター ⌥⌘2', 'ノート マスター ⌥⌘3'],
    '表示 ▸ グリッドとガイド': [
      'スマート ガイド',
      'ガイド ⌃⌥⌘G',
      'グリッド線 ⌘^',
      '----',
      'グリッドに合わせる',
      '----',
      'グリッド オプション',
    ],
    '表示 ▸ ズーム': [
      'ウインドウに合わせる ⌥⌘O',
      'ズーム イン ⌘+',
      'ズーム アウト ⌘-',
      '----',
      'ズーム...',
    ],
    挿入: [
      '新しいスライド ⇧⌘N',
      'スライドの複製 ⇧⌘D',
      'スライドの指定 ▸',
      '----',
      'セクション',
      'コメント ⇧⌘M',
      '----',
      'ワードアート',
      'ヘッダーとフッター...',
      '日付と時刻...',
      'スライド番号',
      '----',
      '表...',
      'グラフ ▸',
      'SmartArt ▸',
      '----',
      '画像 ▸',
      'オーディオ ▸',
      'ビデオ ▸',
      'レリーフ',
      '数式',
      '記号/文字...',
      '図形 ▸',
      '----',
      'アイコン...',
      '----',
      '3D モデル ▸',
      '----',
      'ズーム ▸',
      '----',
      '動作設定ボタン ▸',
      'オブジェクトの動作設定...',
      '----',
      'オブジェクト...',
      'ハイパーリンク... ⌘K',
    ],
    '挿入 ▸ スライドの指定': ['スライドの再利用...', 'アウトライン表示...'],
    '挿入 ▸ グラフ': [
      '縦棒',
      '横棒',
      '折れ線',
      '面',
      '円',
      'ツリーマップ',
      'サンバースト',
      'ヒストグラム',
      'パレート図',
      '箱ひげ図',
      '散布図',
      'ウォーターフォール',
      'じょうご',
      '株価',
      '等高線',
      'レーダー',
      '塗り分けマップ',
    ],
    '挿入 ▸ SmartArt': [
      'リスト',
      '手順',
      '循環',
      '階層構造',
      '集合関係',
      'マトリックス',
      'ピラミッド',
      '図',
      'タイムライン',
      'チームに会う',
      'テキスト カード',
    ],
    '挿入 ▸ 画像': [
      '写真ブラウザー…',
      '画像をファイルから挿入...',
      'ストック画像...',
      'オンライン画像...',
      'ブランド イメージ...',
    ],
    '挿入 ▸ オーディオ': [
      'サウンド ブラウザー...',
      'オーディーをファイルから挿入...',
      '----',
      'オーディオの録音...',
    ],
    '挿入 ▸ ビデオ': [
      'ムービー ブラウザー...',
      'ファイルからムービー...',
      'ストック ビデオ...',
      'オンライン ビデオ...',
    ],
    '挿入 ▸ 図形': ['正方形/長方形', '角丸四角形', '三角形', '円/楕円', '線'],
    '挿入 ▸ 3D モデル': ['3D モデルを挿入', '3D モデルのストック...', 'このデバイス...'],
    '挿入 ▸ ズーム': ['サマリー ズーム', 'セクション ズーム', 'スライド ズーム'],
    '挿入 ▸ 動作設定ボタン': [
      '動作設定ボタン',
      '戻る/前へ',
      '進む/次へ',
      '先頭',
      '終了',
      'ホーム',
      '情報',
      'Return キー',
      'ムービー',
      '文書',
      'サウンド',
      'ヘルプ',
      'ユーザー定義',
    ],
    フォーマット: [
      'フォント... ⌘T',
      '段落... ⌥⌘M',
      '----',
      '箇条書きと段落番号...',
      '----',
      '段組み...',
      '配置 ▸',
      'その他のオプション...',
      '----',
      'オブジェクト スタイルのコピー ⌥⌘C',
      '標準に適用 ⇧⌘V',
      'アニメーションのコピー/貼り付け ⌥⇧⌘C',
      '----',
      'フォントの置換...',
      '----',
      'テーマの配色...',
      'スライドの背景... ⇧⌘2',
      '----',
      'トリミング ⇧C',
      '----',
      'オブジェクトの書式設定... ⇧⌘1',
    ],
    'フォーマット ▸ 配置': ['左揃え ⌘L', '中央揃え ⌘E', '右揃え ⌘R', '両端揃え', '均等割り付け'],
    配置: [
      'オブジェクトの並べ替え',
      '重なり合ったオブジェクトの並べ替え',
      '----',
      '最前面へ移動 ⇧⌘F',
      '最背面へ移動 ⇧⌘B',
      '前面へ移動 ⌥⇧⌘F',
      '背面へ移動 ⌥⇧⌘B',
      '----',
      'グループ化 ⌥⌘G',
      'グループ解除 ⌥⇧⌘G',
      '再グループ化 ⌥⌘J',
      '----',
      '回転/反転 ▸',
      '配置/整列 ▸',
      '----',
      '選択ウインドウ... ⌥⌘U',
    ],
    '配置 ▸ 回転/反転': [
      '左へ 90 度回転',
      '右へ 90 度回転',
      '左右反転',
      '上下反転',
      '----',
      'その他の回転オプション...',
    ],
    '配置 ▸ 配置/整列': [
      '左揃え',
      '左右中央揃え',
      '右揃え',
      '上揃え',
      '上下中央揃え',
      '下揃え',
      '----',
      '左右に整列',
      '上下に整列',
      '----',
      'スライドに合わせて配置',
      '選択したオブジェクトの配置',
    ],
    ツール: [
      'スペル チェック...',
      '類義語辞典... ⌃⌥⌘R',
      '翻訳...',
      '校正言語の設定...',
      'オートコレクトのオプション...',
      '----',
      'アクセシビリティ チェック',
      '----',
      'マクロ ▸',
      'アドイン',
    ],
    'ツール ▸ マクロ': ['マクロ...', 'Visual Basic Editor'],
    'スライド ショー': [
      '最初から再生 ⇧⌘↩',
      '現在のスライドから再生 ⌘↩',
      '----',
      '目的別スライド ショー ▸',
      '----',
      '発表者ツール ⌥↩',
      'リハーサル',
      'スライド ショーの記録',
      '----',
      '非表示スライドに設定',
      'スライド ショーの設定...',
      '----',
      '常に字幕を使用する',
      '字幕の設定 ▸',
    ],
    'スライド ショー ▸ 字幕の設定': [
      '話し手の言語 ▸',
      '字幕の言語 ▸',
      '----',
      'マイク ▸',
      '----',
      '下部 (重ねて表示)',
      '上部 (重ねて表示)',
      'スライドの下',
      'スライドの上',
      '----',
      'システムのキャプション設定...',
    ],
    ウィンドウ: [
      'しまう ⌘M',
      'ズーム',
      '----',
      '新しいウインドウを開く',
      'ウインドウの整列',
      '次のペイン',
      '----',
      'すべてを手前に移動',
      'すべてをしまう',
      'すべてを拡大/縮小',
      '画面全体に表示 🌐⌃F',
      '中央に配置 🌐⌃C',
      '----',
      '移動とサイズ変更 ▸',
      'フルスクリーンのタイル表示',
      '----',
      'ウインドウをセットから削除',
      'ウインドウを整理',
      '----',
      'deck',
    ],
    'ウィンドウ ▸ 移動とサイズ変更': [
      '2分割',
      '左',
      '右',
      '上',
      '下',
      '----',
      '4分割',
      '左上',
      '右上',
      '左下',
      '右下',
      '----',
      '配置',
      '左と右',
      '左と4分割',
      '右と左',
      '右と4分割',
      '上と下',
      '上と4分割',
      '下と上',
      '下と4分割',
      '4分割',
      '----',
      '前のサイズに戻す 🌐⌃R',
    ],
    ヘルプ: [
      'エディター ヘルプ ⌘?',
      '----',
      'フィードバック',
      '----',
      'アプリケーション データのクリア',
      '----',
      '更新プログラムのチェック',
    ],
  },
};

const DECK = `import { Presentation, Slide, Shape } from '@office-kit/pptx-dsl';
export default (
  <Presentation>
    <Slide>
      <Shape preset="rect" x={1} y={1} width={3} height={1.5} text="Hello world" />
      <Shape preset="ellipse" x={6} y={3} width={2} height={1} text="B" />
    </Slide>
  </Presentation>
);
`;
const SHOTS = process.env.MENUBAR_SHOTS;
const TITLES = {
  en: [
    'File',
    'Edit',
    'View',
    'Insert',
    'Format',
    'Arrange',
    'Tools',
    'Slide Show',
    'Window',
    'Help',
  ],
  ja: [
    'ファイル',
    '編集',
    '表示',
    '挿入',
    'フォーマット',
    '配置',
    'ツール',
    'スライド ショー',
    'ウィンドウ',
    'ヘルプ',
  ],
};

// The reference desktop app's enabled state with nothing selected, a shape selected and text
// selected (native capture, menubar/SUMMARY.md). `✓` marks a checked item.
const STATES = [
  ['Edit ▸ Cut', 'off', 'on', 'on'],
  ['Edit ▸ Copy', 'off', 'on', 'on'],
  ['Edit ▸ Duplicate', 'off', 'on', 'off'],
  ['Insert ▸ Symbol...', 'off', 'off', 'on'],
  ['Insert ▸ Action Settings...', 'off', 'on', 'on'],
  ['Insert ▸ Hyperlink...', 'off', 'on', 'on'],
  ['Format ▸ Font...', 'off', 'on', 'on'],
  ['Format ▸ Paragraph...', 'off', 'on', 'on'],
  ['Format ▸ Columns...', 'off', 'on', 'on'],
  ['Format ▸ Alignment ▸ Align Left', 'off', 'on', 'on'],
  ['Format ▸ Alignment ▸ Center', 'off', 'on ✓', 'on ✓'],
  ['Format ▸ Alignment ▸ Align Right', 'off', 'on', 'on'],
  ['Format ▸ Alignment ▸ Justify', 'off', 'on', 'on'],
  ['Format ▸ Alignment ▸ Distributed', 'off', 'on', 'on'],
  ['Format ▸ More Options...', 'off', 'on', 'on'],
  ['Format ▸ Pick Up Object Style', 'off', 'on', 'on'],
  ['Format ▸ Apply To Defaults|Apply Object Style', 'off', 'on', 'on'],
  ['Format ▸ Crop', 'off', 'off', 'off'],
  ['Format ▸ Format Object...', 'off', 'on', 'on'],
  ['Arrange ▸ Bring to Front', 'off', 'on', 'on'],
  ['Arrange ▸ Send to Back', 'off', 'on', 'on'],
  ['Arrange ▸ Bring Forward', 'off', 'on', 'on'],
  ['Arrange ▸ Send Backward', 'off', 'on', 'on'],
  ['Arrange ▸ Rotate or Flip ▸ Rotate Left 90°', 'off', 'on', 'on'],
  ['Arrange ▸ Rotate or Flip ▸ Rotate Right 90°', 'off', 'on', 'on'],
  ['Arrange ▸ Rotate or Flip ▸ Flip Horizontal', 'off', 'on', 'on'],
  ['Arrange ▸ Rotate or Flip ▸ Flip Vertical', 'off', 'on', 'on'],
  ['Arrange ▸ Rotate or Flip ▸ More Rotation Options...', 'off', 'on', 'on'],
  ['Arrange ▸ Align or Distribute ▸ Align Left', 'off', 'on', 'on'],
  ['Arrange ▸ Align or Distribute ▸ Align Center', 'off', 'on', 'on'],
  ['Arrange ▸ Align or Distribute ▸ Align Right', 'off', 'on', 'on'],
  ['Arrange ▸ Align or Distribute ▸ Align Top', 'off', 'on', 'on'],
  ['Arrange ▸ Align or Distribute ▸ Align Middle', 'off', 'on', 'on'],
  ['Arrange ▸ Align or Distribute ▸ Align Bottom', 'off', 'on', 'on'],
  ['Arrange ▸ Align or Distribute ▸ Distribute Horizontally', 'off', 'on', 'on'],
  ['Arrange ▸ Align or Distribute ▸ Distribute Vertically', 'off', 'on', 'on'],
  ['Arrange ▸ Align or Distribute ▸ Align to Slide', 'off ✓', 'on ✓', 'on ✓'],
  ['View ▸ Normal', 'on ✓', 'on ✓', 'on ✓'],
  ['View ▸ Slide Sorter', 'on', 'on', 'on'],
  ['View ▸ Notes Page', 'on', 'on', 'on'],
  ['View ▸ Outline View', 'on', 'on', 'on'],
  ['View ▸ Reading View', 'on', 'on', 'on'],
  ['View ▸ Master ▸ Slide Master', 'on', 'on', 'on'],
  ['View ▸ Master ▸ Handout Master', 'on', 'on', 'on'],
  ['View ▸ Master ▸ Notes Master', 'on', 'on', 'on'],
];
// The editor cannot run these yet; they stay disabled with the reason.
const UNSUPPORTED = new Map([
  ['Format ▸ Bullets and Numbering...', 'The editor has no Bullets and Numbering dialog yet.'],
  ['Tools ▸ Thesaurus...', 'The thesaurus needs an online reference service.'],
]);

async function open(file) {
  const preview = await startPreview(file);
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ viewport: { width: 1500, height: 1000 } });
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  const page = await context.newPage();
  await installRichTextSelection(page);
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto(preview.url);
  const editor = page.frameLocator('#editor-frame');
  await editor.getByText('Saved to this project', { exact: true }).waitFor();
  const deck = async () =>
    loadPresentation(new Uint8Array(await (await fetch(`${preview.url}/deck.pptx`)).arrayBuffer()));
  const changed = async (action) => {
    const before = (await waitForState(preview.url, () => true)).revision;
    await action();
    await waitForState(preview.url, (state) => state.revision !== before);
  };
  return { preview, browser, page, editor, errors, deck, changed };
}

async function deckFile() {
  const dir = await mkdtemp(join(tmpdir(), 'office-menubar-'));
  const file = join(dir, 'deck.tsx');
  await writeFile(file, DECK);
  return { dir, file };
}

async function shot(page, name) {
  if (!SHOTS) return;
  await mkdir(SHOTS, { recursive: true });
  await page.screenshot({ path: join(SHOTS, `${name}.png`) });
}

/** "Label shortcut ▸" per row, "----" per separator. */
function entries(menu) {
  return menu
    .locator(':scope > .item, :scope > .branch > .item, :scope > .sep')
    .evaluateAll((nodes) =>
      nodes.map((node) => {
        if (node.classList.contains('sep')) return '----';
        return [
          node.getAttribute('aria-label'),
          node.querySelector('kbd')?.textContent,
          node.getAttribute('aria-haspopup') ? '▸' : '',
        ]
          .filter(Boolean)
          .join(' ');
      }),
    );
}

/** Opens the menu at `path` (title ▸ submenu …) and returns its deepest menu. */
async function openPath(editor, path) {
  const bar = editor.getByRole('menubar');
  if ((await bar.locator('.title[aria-expanded="true"]').count()) > 0) {
    await bar.locator('.title[aria-expanded="true"]').click();
  }
  await bar.getByRole('menuitem', { name: path[0], exact: true }).click();
  for (let depth = 1; depth < path.length; depth++) {
    await bar
      .locator(`.menu[data-depth="${depth - 1}"] > .branch > .item`)
      .and(editor.getByRole('menuitem', { name: path[depth], exact: true }))
      .hover();
  }
  return bar.locator(`.menu[data-depth="${path.length - 1}"]`);
}

async function closeMenus(editor) {
  const open = editor.locator('.menubar .title[aria-expanded="true"]');
  if ((await open.count()) > 0) await open.click();
}

/** "on"/"off", plus " ✓" when checked, for the item at `path`. */
async function itemState(editor, path) {
  const [label, alternate] = path.at(-1).split('|');
  const menu = await openPath(editor, path.slice(0, -1));
  const rows = menu.locator(':scope > .item, :scope > .branch > .item');
  let item = rows.and(editor.getByLabel(label, { exact: true }));
  if (alternate && (await item.count()) === 0)
    item = rows.and(editor.getByLabel(alternate, { exact: true }));
  const state = await item.evaluate(
    (node) =>
      (node.disabled ? 'off' : 'on') + (node.getAttribute('aria-checked') === 'true' ? ' ✓' : ''),
  );
  await closeMenus(editor);
  return state;
}

test(
  'the menu bar has the reference desktop app’s menus, separators, submenus and shortcuts in English and Japanese',
  { timeout: 300000 },
  async () => {
    const { dir, file } = await deckFile();
    let session;
    try {
      session = await open(file);
      const { page, editor, errors } = session;
      for (const locale of ['en', 'ja']) {
        if (locale === 'ja') {
          await editor.locator('.lang select').selectOption('ja');
          await editor.getByRole('menuitem', { name: 'ファイル', exact: true }).waitFor();
        }
        const bar = editor.locator('.menubar');
        assert.deepEqual(await bar.locator('.title').allTextContents(), TITLES[locale]);
        for (const [key, expected] of Object.entries(NATIVE[locale])) {
          const menu = await openPath(editor, key.split(' ▸ '));
          assert.deepEqual(await entries(menu), expected, `${locale}: ${key}`);
          if (!key.includes(' ▸ ')) await shot(page, `${locale}-${key.replace(/\s+/g, '-')}`);
        }
        await closeMenus(editor);
      }
      // A disabled item names the reason.
      const tools = await openPath(editor, ['ツール']);
      assert.equal(
        await tools
          .getByRole('menuitem', { name: '類義語辞典...', exact: true })
          .getAttribute('title'),
        '類義語辞典にはオンラインの参照サービスが必要です。',
      );
      await closeMenus(editor);
      assert.deepEqual(errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'menu items are enabled and checked as in the reference desktop app with nothing, a shape or text selected',
  { timeout: 300000 },
  async () => {
    const { dir, file } = await deckFile();
    let session;
    try {
      session = await open(file);
      const { page, editor, errors } = session;
      const shape = editor.locator('.hit').first();
      const input = editor.locator('.canvas-shell .inline-edit');
      // A click in a shape's text edits it; Escape leaves the shape selected,
      // and a second Escape clears the selection.
      const contexts = {
        async nothing() {
          await shape.click();
          await page.keyboard.press('Escape');
          await page.keyboard.press('Escape');
        },
        async shape() {
          await shape.click();
          await page.keyboard.press('Escape');
        },
        async text() {
          await shape.click();
          await input.waitFor();
          await input.evaluate((node) => {
            node.focus({ preventScroll: true });
            window.selectEditorText(node, 0, 6);
            node.dispatchEvent(new Event('select', { bubbles: true }));
          });
        },
      };
      for (const [index, name] of ['nothing', 'shape', 'text'].entries()) {
        await contexts[name]();
        const actual = [];
        for (const row of STATES)
          actual.push([row[0], await itemState(editor, row[0].split(' ▸ '))]);
        assert.deepEqual(
          actual,
          STATES.map((row) => [row[0], row[index + 1]]),
          name,
        );
        for (const [path, reason] of UNSUPPORTED) {
          const menu = await openPath(editor, path.split(' ▸ ').slice(0, -1));
          const item = menu.getByRole('menuitem', { name: path.split(' ▸ ').at(-1), exact: true });
          assert.equal(await item.isDisabled(), true, path);
          assert.equal(await item.getAttribute('title'), reason);
          await closeMenus(editor);
        }
        if (name === 'text') {
          // The menus leave the caret and selection in the text.
          assert.deepEqual(
            await input.evaluate((node) => ({
              focused: node === node.ownerDocument.activeElement,
              selected: node.ownerDocument.getSelection().toString(),
            })),
            { focused: true, selected: 'Hello ' },
          );
          await openPath(editor, ['Format']);
          await shot(page, 'en-Format-text-selected');
          await closeMenus(editor);
          await page.keyboard.press('Escape');
        }
      }
      assert.deepEqual(errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);

test(
  'the reference desktop app’s keyboard shortcuts run their menu commands',
  { timeout: 300000 },
  async () => {
    const { dir, file } = await deckFile();
    let session;
    try {
      session = await open(file);
      const { page, editor, errors, deck, changed } = session;
      const shapes = async () => getSlideShapes(getSlides(await deck())[0]);
      const first = editor.locator('.hit').first();
      // A click in the text edits it; Escape leaves the shape selected.
      const selectFirst = async () => {
        await page.keyboard.press('Escape');
        await page.keyboard.press('Escape');
        await first.click();
        await page.keyboard.press('Escape');
      };
      const toast = (text) => editor.getByText(text, { exact: true }).last();

      // ⌘D duplicates the selected shape.
      await selectFirst();
      await changed(() => page.keyboard.press('Meta+d'));
      assert.equal((await shapes()).length, 3);
      await changed(() => page.keyboard.press('Meta+z'));
      assert.equal((await shapes()).length, 2);

      // ⌥⌘G groups and ⌥⇧⌘G ungroups; ⌘G alone is Find Next, not Group.
      await selectFirst();
      await page.keyboard.press('Meta+a');
      await page.keyboard.press('Meta+g');
      await page.waitForTimeout(200);
      assert.equal((await shapes()).length, 2);
      await changed(() => page.keyboard.press('Meta+Alt+g'));
      // The slide lists the group, then its members.
      assert.deepEqual(
        (await shapes()).map((shape) => getShapeKind(shape)),
        ['group', 'shape', 'shape'],
      );
      await changed(() => page.keyboard.press('Meta+Alt+Shift+g'));
      assert.deepEqual(
        (await shapes()).map((shape) => getShapeKind(shape)),
        ['shape', 'shape'],
      );

      // ⌘T opens Font, ⌘K the link dialog, ⇧⌘1 the Format pane.
      await selectFirst();
      await page.keyboard.press('Meta+t');
      const font = editor.getByRole('dialog', { name: 'Font', exact: true });
      await font.waitFor();
      await font.getByRole('button', { name: 'Cancel', exact: true }).click();
      await font.waitFor({ state: 'detached' });
      await page.keyboard.press('Meta+k');
      const link = editor.getByRole('dialog', { name: 'Edit link', exact: true });
      await link.waitFor();
      await link.getByRole('button', { name: 'Close', exact: true }).click();
      await link.waitFor({ state: 'detached' });
      await selectFirst();
      await page.keyboard.press('Meta+Shift+Digit1');
      await editor.getByRole('button', { name: 'Close Format Shape', exact: true }).waitFor();

      // ⇧⌘C picks up the object style in English; ⌥⌘C is not a shortcut there.
      await selectFirst();
      await page.keyboard.press('Meta+Alt+c');
      await page.waitForTimeout(200);
      assert.equal(await editor.getByText('Formatting copied', { exact: true }).count(), 0);
      await page.keyboard.press('Meta+Shift+c');
      await toast('Formatting copied').waitFor();

      // ⌘E and ⌘L align the text being edited.
      const align = async () =>
        getParagraphPropertiesEffective(await deck(), (await shapes())[0], 0).align;
      const input = editor.locator('.canvas-shell .inline-edit');
      await first.click();
      await input.waitFor();
      await input.evaluate((node) => {
        node.focus({ preventScroll: true });
        window.selectEditorText(node, 0, 5);
        node.dispatchEvent(new Event('select', { bubbles: true }));
      });
      await changed(() => page.keyboard.press('Meta+l'));
      assert.equal(await align(), 'left');
      await changed(() => page.keyboard.press('Meta+e'));
      assert.equal(await align(), 'center');
      await page.keyboard.press('Meta+Enter');

      // ⌘4 and ⌘1 switch views; ⇧⌘N adds a slide; ⌥⌘R hides the ribbon.
      await page.keyboard.press('Meta+4');
      await editor.getByRole('navigation', { name: 'Outline View', exact: true }).waitFor();
      await page.keyboard.press('Meta+1');
      await editor
        .getByRole('navigation', { name: 'Outline View', exact: true })
        .waitFor({ state: 'detached' });
      await changed(() => editor.locator('body').press('Meta+Shift+n'));
      assert.equal(getSlides(await deck()).length, 2);
      await page.keyboard.press('Meta+Alt+r');
      const home = editor.getByRole('tab', { name: 'Home', exact: true });
      await home.waitFor({ state: 'detached' });
      await page.keyboard.press('Meta+Alt+r');
      await home.waitFor();

      // ⌘? (Help ▸ Editor Help) opens the command search.
      await page.keyboard.press('Meta+Shift+Slash');
      await editor.getByRole('dialog', { name: 'Command palette', exact: true }).waitFor();
      await page.keyboard.press('Escape');

      // The reference desktop app's Japanese build picks up the object style with ⌥⌘C instead.
      await editor.locator('.lang select').selectOption('ja');
      await editor.locator('.nav [data-slide-index="0"]').click();
      await selectFirst();
      await page.keyboard.press('Meta+Shift+c');
      await page.waitForTimeout(200);
      assert.equal(await editor.getByText('書式をコピーしました', { exact: true }).count(), 0);
      await page.keyboard.press('Meta+Alt+c');
      await toast('書式をコピーしました').waitFor();
      assert.deepEqual(errors, []);
    } finally {
      await session?.browser.close();
      await session?.preview.close();
      await rm(dir, { recursive: true, force: true });
    }
  },
);
