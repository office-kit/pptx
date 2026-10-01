# 作業継続メモ（2026-10-01）

## 再開時の指示

2026-10-01 にユーザーが完成までの作業再開を明示した。以前の停止指示は解除済み。PR #287 を唯一の提出先として実装・実機比較・検証を続ける。

最終目標は、インストール済み **Mac デスクトップ版 PowerPoint と全操作の UI・動作を揃える**こと。既存機能だけでなく未実装操作も対象と明示されている。全体は未完成であり、今回の PR 整理を完全一致の達成と扱わない。

1. `CLAUDE.md` を読む。既存の公開 API を使い、機能ごとに API を重複させない。pnpm を使う。
2. `git status` と PR #287 の状態・CI を確認する。今回の修正は `fix(editor): preserve custom tabs during direct text editing` のコミットを探す。
3. このファイルと `packages/dev/POWERPOINT_PARITY.md` を読む。後者は履歴形式で、古い「未実装」が後段で解決されている場合があるため現コードと照合する。
4. 下記の「次に進める項目」から、実機確認・失敗するテスト・修正・検証の順で進める。

## Git / PR

- リポジトリ: `/Users/baseballyama/git/pptx`
- ローカルブランチ: `integrate/pptx-editor-pr287`
- **唯一の PR**: https://github.com/office-kit/pptx/pull/287
- リモートブランチ: `feat/pptx-editor`
- push: `git push origin HEAD:feat/pptx-editor`
- 新しい PR を勝手に並立させない。マージや強制 push は今回依頼されていない。
- `.pnpm-store/` は以前からの untracked ディレクトリ。コミット・削除対象にしない。
- 古い `feat/mac-powerpoint-parity` ブランチは完全統合されていない。丸ごとマージせず、必要な機能を現実装と比較する。

## この区切りの実装

- 段落の個別タブ・既定タブ間隔の公開 API と、Paragraph > Tabs ダイアログの保存・取消・Undo を実装済み。
- タブダイアログの Set / Clear / Clear All の有効状態と位置欄リセットを Mac 実機で確認。
- 通常プレビューでは、個別タブを含む本文に SVG 配置とブラウザー実フォントの計測を使用。既定タブは HTML の `tab-size` で反映。
- 今回は直接編集時の個別タブにも左・中央・右・小数点揃えを反映。実際のタブ文字を保持するため、選択・コピーの UTF-16 オフセットを変えない。ズーム変化時は再配置する。
- 直接編集で 5.08 cm が既定 2.54 cm に戻る失敗をテストで再現し、修正後に入力・Undo を含めて確認。

主なファイル:

- `site/src/lib/editor/core/inline-text-html.ts`: 有効段落プロパティから編集用 HTML を構築。
- `site/src/lib/editor/core/editing-tabs.ts`: タブ文字の幅と後続フィールド揃えを計算。
- `site/src/lib/editor/ui/RichTextInput.svelte`: HTML 再描画後にタブ配置と選択復元。
- `site/src/lib/editor/core/rich-text-dom.ts`: 編集文字列と選択オフセット。
- `packages/preview/src/text-layout.ts`: SVG タブ配置。
- `packages/preview/src/browser-measure.ts`: ブラウザーフォント計測。
- `packages/dev/test/browser/custom-tab-spacing.test.mjs`: 4 揃えの表示・直接編集・入力・Undo。
- `packages/dev/test/browser/default-tab-spacing.test.mjs`: 既定間隔の表示・直接編集。

## 次に進める項目

最初はタブとルーラーの境界条件を実機と比較するのが自然な続き。

- 複数のタブ、書式の異なる run をまたぐフィールド、改行・折り返し、狭い本文での配置。
- 中央・右寄せ段落、箇条書き、インデント、表セル、回転・縦書き、極端なズーム。
- タブ前後への入力、範囲選択・コピー・貼り付け、IME、Undo/Redo、保存後再読込。
- 編集用配置は `offsetLeft` による整数位置と canvas 幅を使う。サブピクセル精度・字間・折り返し・段落配置がネイティブと完全一致すると主張しない。
- 小数点は現在 ASCII `.` を基準にする。ロケールによる小数記号やネイティブ処理を要確認。
- ルーラーからのタブ追加・移動・削除、タブ種別切替は実装し、保存・取消・Undo のブラウザー検証が通過。インデントのドラッグは一部対応しているが、ネイティブのスナップ・混在表示・ライブ再配置は未確認。
- Paragraph Typography のフラグは API/UI で編集できるが、プレビューの改行規則への完全反映は未実装。

Outline View の表示・文字編集・タイトル分割を実装。折りたたみ、階層変更、本文のレベル表示、スライドをまたぐ選択などは残る。全体の残件には リボン・メニューの完全一致、残る書式/画像補正/テクスチャ操作、背景画像のネイティブセッション記憶との差、図形に沿うグラデーション等がある。網羅的な全操作監査はまだ終わっていない。既存の比較記録と現コードから優先順位を決める。

## Mac 実機

- CUA から `com.microsoft.Powerpoint` に接続する。過去には接続断があったが直近の監査では復旧済み。再開時に状態を取り直す。
- 直近の参照: `/private/tmp/pptx-native-selection-lock-audit.pptx`。元の `/tmp/pptx-macro-audit/support.pptx` も比較に使われた。
- ユーザーは参照文書への一時変更と確認後の取り消しを許可済み。
- 直近の Tabs / Paragraph 監査は両方 Cancel で終了し、未復元の変更は残していない。今回の編集タブ修正ではネイティブ文書を変更していない。
- 実機の変更は比較後に Undo し、保存済み状態・Undo 状態で復元を確認する。ネイティブ操作に AppleScript を使わない。

## 検証と環境

- この区切りで format、lint、ルート TypeScript、Svelte（0 errors / 0 warnings）、エディタ build、上記 2 ブラウザーテストが通過。
- PR の Testing に最後の全体テスト結果と CI 状況を記載する。過去の全ブラウザー一括実行は 189/190 で、その後失敗箇所の個別修正が通過した状態。全ブラウザーが最終 HEAD で一括成功済みとは扱わない。
- Node: `/Users/baseballyama/.nodenv/versions/26.6.0/bin/node`。サブプロセスを起動する検証前には必ず同ディレクトリを PATH の先頭にする。
- pnpm の shim が停止する問題があった。今回の検証はインストール済みローカル CLI を直接 Node で起動した。npm/yarn へ切り替えてはいない。
- ブラウザーテスト実行中に build を走らせない。共有 dist を更新すると検証対象が変わる。
- DSL 型チェックは core と preview の build 後に行う。古い dist のチェック結果を信頼しない。

再現用のコマンド例（リポジトリルート）:

```sh
export PATH=/Users/baseballyama/.nodenv/versions/26.6.0/bin:$PATH
node node_modules/oxfmt/bin/oxfmt --check
node node_modules/oxlint/bin/oxlint
node node_modules/typescript/bin/tsc --noEmit
node node_modules/vitest/vitest.mjs run --maxWorkers=2
node node_modules/tsdown/dist/run.mjs
node packages/dev/build-editor.mjs
node --test packages/dev/test/browser/custom-tab-spacing.test.mjs packages/dev/test/browser/default-tab-spacing.test.mjs
```

Svelte は `site` で `node node_modules/svelte-check/bin/svelte-check --tsconfig ./tsconfig.json`。
ブラウザー起動には sandbox 外実行が必要だった。許可済みの git/PR 更新を毎回ユーザーに再確認せず、必要な環境承認だけを処理する。

## 再開後の進捗（2026-10-01）

- PR #287 の Node 24 CI で失敗した表セルの文字色テストを現行の More Colors 入力へ修正し、個別再実行で通過。
- main 909cefa6 を統合。ソース対応付けの shape/cell/paragraph 属性を既存のアニメーション対象属性と重複させず保持。関連 59 件と全ユニット 3,048 件が通過（109 件 skip）。
- ルーラーのタブ操作は選択段落へ適用。4 揃えの追加、ドラッグ移動・ルーラー外で削除、Escape 取消、Undo、再読込を新規ブラウザーテストで確認。既存インデント検証も通過。
- Mac 実機は再接続済み。ルーラークリックでタブ追加を確認して Undo。Ruler off に復元。Outline View の普通の Text box が一覧に現れないことを確認し Normal 120% へ復元。Undo disabled。
- ネイティブのドラッグは AXError.failure となり未確認。混在タブ表示・回転/縦書き・ライブ再配置・スナップは未完了。全操作・全 UI の一致は達成していない。
- Outline View の文字編集・タイトル末尾/途中の Enter・Undo・保存・表示切替時の確定を実装。日本語/英語のブラウザーテストが通過。site ユニットは 81 件通過、Svelte は 0 errors / 0 warnings。
- 全ブラウザー検証を再実行中。最初の実行は terminal-client.js/css のローカルビルド漏れで Agents 画面が開けず中止。`packages/dev` で `node build-terminal.mjs` を実行し、完全な成果物からやり直した。部分的な成功を全体成功と扱わない。
