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

Outline View の表示・文字編集・タイトル分割を実装。アイコンのダブルクリックによる折りたたみと保存・Undo も実装。階層変更、本文のレベル表示、スライドをまたぐ選択などは残る。全体の残件には リボン・メニューの完全一致、残る書式/画像補正/テクスチャ操作、背景画像のネイティブセッション記憶との差、図形に沿うグラデーション等がある。網羅的な全操作監査はまだ終わっていない。既存の比較記録と現コードから優先順位を決める。

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
- アウトラインのアイコン右クリックに折りたたみ/展開のサブメニューと全スライド操作を実装。バッチ操作は 1 回で Undo。本文編集後のアイコン右クリックで図形用メニューが出る不具合も修正。日本語/英語の独立ビルド検証 2 件、コア全 3,051 件が通過（109 skip）。
- 全ブラウザー再実行は 212 件中 210 件成功、アウトライン 2 件失敗。実行中にテストを更新して共有ビルドとの差が生じたため、最新の通常ビルドでアウトラインと関連メニューを再検証する。これを全件成功と記載しない。
- 最新の通常ビルドでアウトライン日本語/英語、既存 canvas context、context menu keyboard の計 4 件が成功。format/lint、型検査、Svelte 0 errors / 0 warnings、core/editor build も通過。
- 参照 `/private/tmp/pptx-outline-audit/body.pptx` は Normal 120%、Undo disabled に復元済み。Outline View に入り Return、Undo で AX の本文入力要素へ入れる。`selectText` で本文先頭/途中へ移動し Tab を押すと `a:pPr lvl="1"` が保存された。Shift+Tab は操作接続経由の結果が不確かで未確認。全メニュー一致、階層表示/編集、スライドをまたぐ選択は未実装。

- 追加: OutlineText の本文 Tab は実段落の範囲を求めて `setParagraphLevel` を適用する。書式・文字列の保持、保存、Undo、選択終端をブラウザーテストで確認する。階層の表示、Shift+Tab、タイトルの降格、Promote/Demote メニューは引き続き未完成。

## Outline shortcuts / Promote follow-up

- Command-4 opens Outline View; Command-1 leaves it and commits pending text. The View menu includes the shortcut label.
- Native Promote of a level-zero body paragraph splits the slide: promoted text becomes the new title and following body paragraphs move with it. Implement this separately from changing `a:pPr/@lvl`, preserving paragraph/run formatting and relationships. Native audit restored the reference document with Undo disabled.

- タイトル分割でリンクが消えるブラウザー回帰を再現し、既存 `setShapeParagraphs` に `{ source, range? }` 入力を追加して修正。文字範囲の段落 XML・フィールド・リンクをコピーし、宛先の本文設定は保持。英語/日本語のブラウザー検証が通過。
- 実機 `/private/tmp/pptx-outline-audit/body.pptx` では、レベル 0 の本文を 2 段落選択して Promote すると、それぞれ別スライドのタイトルになる。単一段落では後続本文も新スライドへ移る。昇格操作は未実装。比較後は Undo、Normal 120%、1 スライド、Undo disabled に復元・保存済み。

- アウトラインのスライドアイコン右クリックに Move Up / Move Down を追加。英語/日本語の独立ビルドで複数選択・境界・Undo/Redo・保存順序・再読込が通過。全ブラウザー実行中の共有ビルドは変更していないため、完了後の通常ビルドでも `outline-slides.test.mjs` の英日 2 件が通過。

- Paragraph-level range updates now use one text-body commit through the existing `getParagraphLevel` / `setParagraphLevel` APIs. Outline Tab no longer performs repeated whole-text scans/commits. Core range tests cover shapes and table cells.
- Clipboard boundary review reproduced and fixed invalid format metadata and sub-point HTML font sizes. Editor unit suite: 83 passed; new HTML-size browser case passed. Frozen full-suite run `/tmp/pptx-browser-10e68754.log` finished: 212 passed, zero failures. Core/editor were then rebuilt and DSL/Svelte checks passed. Focused ordinary-build validation is recorded below.

- CI `10e68754` の Node 24 はアウトラインの遅延 select イベントで 1 件失敗（211 件成功）。view 切替で破棄済みの textarea に select を配送する英日テストで再現し、イベントの currentTarget を読むよう修正。タイトル分割後のフォーカス処理も破棄前の ownerDocument を保持する。独立ビルド 2 件成功、通常ビルドの関連 11 件も全て成功（アウトライン編集・移動・選択イベント各英日、書式付き貼り付け、HTML 貼り付け・フォントサイズ境界）。

## 継続実装: 本文の昇格

- `promoteOutlineBody` を `site/src/lib/editor/core/outline.ts` に追加。入れ子段落は一段上げ、選択したルート段落ごとに新しいスライドを作る。Shift+Tab と本文の右クリック Promote が使用する。
- 現レイアウトにタイトル・本文枠がない場合は挿入前にエラーを表示する。Title Demote、スライドをまたぐ文字選択、書式・階層のアウトライン表示、本文メニューの残りは未完成。
- コアの既存 API に一括処理を追加: `addSlideAt(pres,index,options[])`、`setShapeParagraphs(targets,{source,ranges})`、`getShapeParagraphElements(shape)`。元の単一対象 API も存続。
- 実機比較はこの段階では参照ファイルへの追加変更なし。本文単一ルート・複数ルート・入れ子 Promote は前段の実機結果を参照。複数レベルが混在する選択・特殊レイアウトは更に比較が必要。

- 混在レベルを実機比較済み。選択した Second（ルート）は新タイトル、Third（レベル 1）は新本文のレベル 0 となる。保存 XML も確認。参照文書は 2 回 Undo、Normal 120%、1 スライド、Undo disabled に復元して保存。
- 英語/日本語のブラウザー検証は右クリックによる複数ルート昇格・Undo/Redo まで両方成功。末尾空段落と混在レベルは保存・再読込のユニット検証を追加。

## 継続実装: クリップボードとタイトル降格

- アウトラインの Copy/Cut/Paste は未確定入力と文字書式を保持し、外部 HTML も受け付ける。Undo・保存を英日ブラウザーで確認。
- タイトルの Tab / Demote は前スライド本文へタイトルと本文を結合する。先頭スライドは変更しない。追加図形や宛先本文なしは変更前にエラー。該当レイアウトの実機比較は残る。
- `setShapeParagraphs(target,{sources})` で段落 XML・リンク・フィールドを一括結合する。元データを全てコピーした後で宛先を更新する。
- 実機でタイトルの降格と後続本文のレベル保持を確認。参照 body.pptx は Undo、Normal 120%、1 枚、Undo disabled に復元・保存済み。未復元変更なし。
- コア 3,068 成功 / 109 skip、site 88 成功、関連ブラウザー英日 6 件成功。Svelte 0 errors / 0 warnings、root 型検査・core/editor build 成功。全 UI・全操作一致は未完成。
- 次は本文右クリックのクリップボード操作、アウトラインの書式/階層表示、スライド間選択など。過去の未実装記述は本節で更新される。

## 最新追記: アウトラインの右クリックとドラッグ

- 右クリック Cut/Copy/Paste を実装。HTML 書式保持と非同期クリップボード応答中のフォーカス移動を検証。
- スライドアイコンのドラッグ並べ替えを実装。連続・非連続の複数選択、Undo、保存後の再読込を英日 4 ケースで検証。Mac の追加選択は Command-click。
- 残件表は POWERPOINT_PARITY.md の冒頭を参照。ドラッグの実機との完全一致やスクロール、本文の視覚的階層・書式表示、スライド間の文字選択などは未完了。

## 最新追記: 本文段落の上下移動

- 本文右クリックの Move Up/Down を追加。選択段落を隣の段落と入れ替え、段落 XML・書式・リンクを保持する。既存 `setShapeParagraphs` の単一対象にも `{ source, ranges }` を追加。
- 実機では後続の入れ子段落は自動で追従しない。先頭本文では Move Up が無効。比較後は元の本文・レベル、Normal 120%、Undo disabled に復元・保存済み。未復元変更なし。
- コア 3,069 成功 / 109 skip、site 89 成功。タイトル文字やスライド境界をまたぐ移動は未検証。全操作一致は未完成。

## 最新追記: スライド参照と並べ替え

- 絶対参照と `./` を含む参照で、読み込み・並べ替え・単一移動が失敗する不具合を再現して修正。複数スライドの移動は一括 sort に変更。英日ブラウザー 2 件、site 89 件成功。詳細は POWERPOINT_PARITY.md。
- 実機タイトル文字の Move Up は前スライド末尾の本文を現在スライド先頭へ、Move Down は現在スライド先頭の本文を前スライド末尾へ移す。スライドアイコンの移動とは違う。実装は残件。
- 実機参照は Normal 120%、1 枚、Undo disabled に復元して保存済み。未復元変更なし。

## 最新追記: タイトル文字の上下移動

- 上記のタイトル境界移動を実装。前後スライドの本文を既存段落 API で結合・再配分し、段落レベル・リンク・書式を保持。先頭タイトルと移動元本文が空の場合は無効。
- 本文プレースホルダーの欠けたレイアウト、複雑なスライド境界選択は残件。
- コア 3,071 成功 / 109 skip、site 91 成功、英日アウトライン browser 2 件成功。通常ビルドで保存・Undo・選択・書式を確認。全 UI 一致は未完成。

## 最新追記: タイトル降格の確認と本文枠復元

- 追加オブジェクトがあるタイトル降格は Yes/No ダイアログを表示。No は文書を変更せず、Yes 後の Undo はスライドとオブジェクトを復元。
- 前スライドの本文枠が削除済みならレイアウトから復元。Title Only ではレイアウトを維持し、マスター由来の本文枠を追加。実機保存 XML の unmatched idx を確認。
- コア 3,072 成功 / 109 skip、site 93 成功、英日ブラウザー 4 件成功。型検査 0 errors / 0 warnings。実機の一時変更はすべて取り消して保存済み。
- 全操作一致は未完成。残件は POWERPOINT_PARITY.md 冒頭を参照。
