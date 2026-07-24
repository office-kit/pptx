// Japanese overrides. Keyed by the English UI string (see i18n.svelte.ts).
// Anything not listed falls back to its English key.

export const ja: Record<string, string> = {
  // TopBar / file
  New: '新規',
  Open: '開く',
  Save: '保存',
  'Open .pptx': '.pptx を開く',
  'Save as .pptx': '.pptx として保存',
  'Undo (Ctrl+Z)': '元に戻す (Ctrl+Z)',
  'Redo (Ctrl+Y)': 'やり直し (Ctrl+Y)',
  Editor: 'エディター',
  'Unsaved changes': '未保存の変更',
  'Command palette (Ctrl+K)': 'コマンドパレット (Ctrl+K)',
  'All capabilities': 'すべての機能',
  Opened: '開きました:',
  'Open failed': '読み込みに失敗しました',
  'Saved .pptx': '.pptx を保存しました',
  'Save failed': '保存に失敗しました',
  Language: '言語',

  // Ribbon tabs
  Home: 'ホーム',
  Insert: '挿入',
  Design: 'デザイン',
  Transitions: '画面切り替え',
  Animations: 'アニメーション',
  'Shape Format': '図形の書式',
  Table: '表',

  // Ribbon groups
  Slides: 'スライド',
  Font: 'フォント',
  Paragraph: '段落',
  Drawing: '図形描画',
  Arrange: '配置',
  Editing: '編集',
  Tables: '表',
  Illustrations: '図',
  Text: 'テキスト',
  Comments: 'コメント',
  'Slide setup': 'スライドの設定',
  Background: '背景',
  Theme: 'テーマ',
  Transition: '画面切り替え',
  Animation: 'アニメーション',
  Fill: '塗りつぶし',
  Outline: '枠線',
  Effects: '効果',
  'Size & rotate': 'サイズと回転',
  'Rows & columns': '行と列',
  Cell: 'セル',
  'Table style': '表のスタイル',

  // Ribbon explicit item labels
  'Text format': '文字の書式',
  'Run format': 'ラン書式',
  Bullets: '箇条書き',
  Replace: '置換',

  // Context menu
  Cut: '切り取り',
  Copy: 'コピー',
  Paste: '貼り付け',
  Duplicate: '複製',
  Delete: '削除',
  'Bring to front': '最前面へ移動',
  'Bring forward': '前面へ移動',
  'Send backward': '背面へ移動',
  'Send to back': '最背面へ移動',
  Group: 'グループ化',
  Ungroup: 'グループ解除',
  'Select all': 'すべて選択',

  // Status bar
  Slide: 'スライド',
  'No selection': '選択なし',
  Fit: '全体表示',
  'Zoom in (Ctrl+=)': '拡大 (Ctrl+=)',
  'Zoom out (Ctrl+-)': '縮小 (Ctrl+-)',
  'Reset to 100%': '100% に戻す',
  'Fit (Ctrl+0)': '全体表示 (Ctrl+0)',

  // Slide navigator
  'New slide': '新しいスライド',
  'Move slide': 'スライドを移動',

  // Command palette
  'Search capabilities…': '機能を検索…（例: gradient, table, transition）',
  'No capability matches': '該当する機能がありません',
  navigate: '移動',
  run: '実行',
  close: '閉じる',

  // Bespoke property controls
  'Fill & outline': '塗りつぶしと枠線',
  'Position & size (in)': '位置とサイズ（インチ）',
  Rotation: '回転',

  // Properties panel
  actions: '個の操作',
  'All applicable capabilities': '適用可能なすべての機能',
  Shape: '図形',
  shapes: '個の図形',
  'Table cell': '表のセル',

  // Command dialog
  Run: '実行',
  Cancel: 'キャンセル',
  Apply: '適用',
  Close: '閉じる',
  'This command takes no arguments.': 'この操作に引数はありません。',
  'operates on': '対象',
  category: 'カテゴリ',
  returns: '戻り値',
};
