## 2026-10-01: 追加回帰と残件の確認

- 42f5b21d が PR #287 の remote HEAD に反映済み。最新 CI run36871528222 は Static / OOXML 成功、Node 22/24/26 と Preview fidelity は確認時実行中。
- 同じアニメーション付きスライドが連続する custom show の専用回帰を追加。1回目の表示で animation cursor 0→1、次の出現で0へ戻ることを確認。show-properties-playback は6/6成功、対象format/lint成功。
- 追加レビューで、再生対象の Custom Show を削除しても showProperties の参照 ID が残る問題を発見。未修正。削除後の設定・保存再読込・Undoの回帰を先に追加して修正すること。
- PowerPoint の画面取得は引き続き0×0で失敗。変更・復元待ちはなし。ユーザーの元プレビューは再読み込みしていない。

## 2026-10-01: スライドショー統合検証

- 旧 c433 ビルドの全ブラウザ検証 session9572 は終了: **244 passed / 0 failed**（`/tmp/pptx-browser-c433.log`）。このセッションを再 poll しない。
- 最新ソースで core/preview/DSL/dev の依存順 build 成功（`/tmp/pptx-show-integrated-build.log`）。dev 型検査成功、site 0 errors / 0 warnings、site 117 tests passed。
- Custom Shows は順序・重複を扱う Add/Remove/Up/Down UI と native dialog 化を完了。ブラウザで 3,1,3 の保存後 PPTX、編集・複製・削除・Undo・再読込・日英 UI を検証し成功。
- 再生統合の重複スライド位置・Home/End・タイマー再始動を追加修正。新規設定画面1件、再生5件、既存presentation/animation/transition15件が成功。compact-editorも日英・900/1500pxで1件成功（`/tmp/pptx-show-compact-final.log`）。Home/Endと非表示端の追加回帰も成功（agent session6327終了）。重複した同一スライドのアニメーション再始動についてはコードを確認したが専用ブラウザー回帰は未追加。
- PowerPoint 実機は本ターンも画面サイズ 0×0 で取得失敗。操作・変更はしておらず、復元待ちなし。window/kiosk 実動作とナレーション再生は未対応。全 UI/操作の一致は未完成。
- PR #287 は OPEN、push 先 feat/pptx-editor のまま。直近 push 4ce の CI は Static/Node22/Node26/OOXML 成功、Node24 確認時進行中、既知の Preview fidelity 失敗は継続。ユーザーの元プレビューの編集とソース競合は保持。

## 2026-10-01: スライドショー統合レビュー継続中

- root全テスト3103 passed /109 skipped、format/lint/typecheck成功（`/tmp/pptx-show-core-final.log`、session7701完了）。新規APIのmanifest漏れを修正し、paletteから各専用dialogを開く導線も追加。新コードの共有build/browser検証はまだ未実行。
- build.tsとserver.tsの`/state`にshowProperties/customShows(slideIndices)を追加。outline_selectionがpage.ts再生統合を担当。初回実装はindexOfで重複スライド位置を失うので差戻し済み、showCursor主体へ修正中。poll時の位置保持、全hidden、同一slide連続時animation再始動も依頼。
- slide_insertionはCustomShowsDialogをcheckbox方式から順序付きリスト(Add/Remove/Up/Down、重複可能)へ改善中。coreの未知XML保持とMap化も担当。entries.shiftのO(n²)、custShowLst未知子削除を指摘し修正依頼済み。custom-shows browser回帰追加中。
- outline_selectionのshow-properties-dialog.test.mjsはwaitForStateが/editor/stateを読むのにshowPropertiesを参照する問題を指摘、/state取得に修正依頼。新再生テストは未実行。
- 全browser旧c433検証session9572は同じ実行のまま生存、slide-numbers付近まで進行し観測失敗なし。完了まで共有distを書き換えない。
- 本ターンではcommit/pushなし。全コードが作業ツリーにあり、両agent作業も未統合扱い。新機能にwindow/kiosk動作・narration音声制御は未対応とchangesetに明記。完全一致未完了。

## 2026-10-01: スライドショー機能の移植再開（作業中）

- PR #287 の push 済み HEAD は4ce7310d。ローカルHEAD a276af20はアウトライン実装状況の文書更新。作業ツリーにはスライドショー機能の移植途中の変更がある。
- 旧branchの未統合機能として Custom Shows、Set Up Slide Show、発表中インク操作を確認。旧branchでもこれらの実機比較は未完了で、ブラウザー検証済みと実機一致を混同しない。
- 親担当: show-properties APIと回帰を移植。表示モード、範囲、ループ、ナレーション、アニメーション、タイミングの保存。対象5件成功、custom-showsと合わせ8件成功、root型検査・対象format/lint成功。xmllintが利用可能で、テスト内のXSD検証も実施。kiosk restartの既定300000はbundled pml.xsdに明記されている。
- slide_insertion担当: Custom Shows API/編集UI/再生統合。outline_selection担当: ShowPropertiesDialogと導線。両agentと編集範囲を調整済み。まだUIの完成・ビルド・ブラウザー検証・commit/pushは未完了。
- 同じ全ブラウザー検証session9572は生存中。/tmp/pptx-browser-c433.logはアウトライン付近まで進行、観測時失敗なし。完了まで共有distを上書きしない。この検証対象はc433のビルドであり新しい変更の検証にはならない。
- 最新CI run36866882930（4ce）: Static/Node22/Node26/OOXML成功、Node24実行中、Preview fidelity失敗。05-preset-shapes slide1の単語折り返し差分は同じ。OfficeのlatinLnBrk既定falseだけでは緊急分割の有無を断定できず、根拠なくbaselineや実装を変えない。
- PowerPointの画面取得は再度0×0で失敗。Document Controlにも接続セッションなし。実機編集・復元待ちはない。ユーザープレビューの編集とソース競合を保持する。

## 2026-10-01: 日本語禁則処理の書式境界

- c4335475までPR #287へpush済み。その後、句点だけ太字にすると「甲乙／。」になる回帰を失敗するテストで再現し修正。論理wordをrun横断で分割してから元runへ対応付け、書式・文字サイズを維持。単一runの狭幅時の緊急改行は保持。
- 検証: root3093 passed / 109 skipped（/tmp/pptx-cjk-root-tests.log）、root/preview型検査、lint、対象format成功。previewは実行中ブラウザー検証のdistを保つため /tmp/pptx-cjk-preview-build に分離build成功。
- c4335475の全ブラウザー検証はsession9572、/tmp/pptx-browser-c433.logで実行中。既存のhandleをpollすること。観測時点では失敗なし。完了まで共有distを上書きしない。
- c4335475のCI run36864790119: Static/Node22/Node26/OOXML成功、Node24実行中、Preview fidelity失敗。既知の05-preset-shapes slide1 fgSSIM .7731→.6161は未解決。PowerPoint実機の画面取得は0×0で失敗し、実機変更・復元待ちはない。
- 検討したグラフ回転無視の修正は撤回済み。MS-OI29500 §21.3.2.28のOffice制限はchart drawings内のgraphic framesが対象であり、スライド上のp:xfrmへ適用できる根拠ではない。p:xfrm自体はrot/flipH/flipVを許す。根拠なくこの修正を再適用しない。
- 既存ユーザープレビューの編集内容とソース競合を保持。勝手にreload/競合解消しない。全操作のPowerPoint一致は未完成。

## 2026-10-01: アウトライン書式の統合検証

- ヘッダー40px・重複バー削除・保存状態統合は実装済み。同じPR #287を使用し、ユーザーのプレビュー資料と既存タブは保持。
- 8a9a1fddまでpush済み。CIはStatic/Node22/Node26/OOXML成功、Node24確認時実行中、Preview fidelity失敗。描画一致の未解決事項は下記を参照。
- アウトラインの複数フィールドにまたがる文字・段落書式をリボンへ接続。選択文字だけのBold適用と1回のUndoを実ブラウザーで検証。queued caretがリボンのフォーカスを奪う不具合も旧ビルドで再現後に修正。
- 統合検証: アウトラインのコピー/切り取り/通常入力/貼り付け/IME/左右キー/右クリック/フォーカス保護9/9成功（/tmp/pptx-outline-integration-browser.log）。ヘッダー・フォント・既存インライン編集3/3成功（/tmp/pptx-final-header-font.log）。site117件、Svelte 0 errors / 0 warnings、lint、dev依存込みbuild成功。
- 最終統合ブラウザー15/15成功（/tmp/pptx-outline-final-browser.log）。最終Svelte 0 errors / 0 warnings、format/lint成功。
- カーソル位置の次回入力書式（Bold+Italicの連続指定）とキャンバス選択時の旧範囲解除も実装。Svelte $state はAPIオブジェクトをproxy化するため、所有権の解除はオブジェクトでなくapplyコールバックの同一性を比較する。キャンバス誤適用を旧ビルドで再現し修正後成功。全操作のPowerPoint一致は未完成。

## 2026-10-01: 書式境界の折り返しとリボン選択を継続

- PR #287 の push 済み HEAD は63d7a728。前ターンは跨ぎ選択の修正・テスト・pushを完了した進捗ありのターン。
- 書式runの境界をLatin単語の改行位置として扱う不具合を失敗する2テストで再現し修正。Latin途中改行無効時は複数runを一つの単語として幅判定し、各runの書式は維持。関連58件、全root3091件（109 skipped）、lint、preview型検査・build成功。単一runの長い単語の既存fidelity差分を解決する変更ではない。
- PowerPoint実機の再接続は画面0×0で失敗。資料変更・復元待ちはない。
- アウトラインのリボン連携とqueued caretのフォーカス保護を別担当が実装・検証中。これらの完成・統合結果は次の追記を確認すること。

## 2026-10-01: ヘッダーと文字編集の統合検証

- 最新プレビューのヘッダー40px、重複操作バー除去、保存状態のTopBar統合をChrome1494127186で目視確認。日英900/1500pxのブラウザーテスト成功。ユーザーの旧タブと4スライドの編集・競合通知は保持。
- 共通selectRichTextのBR境界ずれを失敗するテストで再現して修正。BR・ブロック境界のUTF-16位置対応を共通化。
- 統合検証: Svelte 0 errors / 0 warnings、site113件、lint、対象format、dev build成功。ヘッダー・既存インライン編集・既存アウトライン日英4/4成功（/tmp/pptx-compact-outline-integrated.log）。フォント候補と透明な編集背景2/2成功（/tmp/pptx-font-background-final.log）。
- アウトライン跨ぎ選択の統合ブラウザー8/8成功（/tmp/pptx-outline-integrated-final.log）。コピー・カット・ペースト・通常入力・Shift+Enter・実IME確定とUndo・左右キーでの選択解除・並べ替え後の追加編集を検証。IMEは全スライドの文字列を照合。モデル再登録時の旧cleanupと古いcaret復元を修正。右クリックで選択が解除される不具合も修正し、実際のCopyメニューからOSクリップボードへ書き込む追加テスト1/1成功（/tmp/pptx-outline-context-corrected.log）。
- c91c3363のCIはStatic/Node22/Node26/OOXML成功、Preview fidelity失敗、Node24確認時実行中（run36858403975）。fidelityの失敗は同じ05-preset-shapes slide1（fg-SSIM 0.7731→0.6161）。latinLnBrk false/省略時の長い単語の緊急折返しはPowerPoint実機未確認。基準値は変更していない。
- 追加の選択ライフサイクル確認: 置換直後のrAFより先にRibbonへクリックした場合のフォーカス復元、跨ぎ選択へのRibbon書式適用は未検証。キーボード/クリップボードの成功と混同しない。
- 最新PowerPoint画面取得はcgWindowNotFoundで失敗。実機変更・復元待ちはない。全操作一致は未完成。

# 作業継続メモ（2026-10-01）

## 最新検証: Latin 改行の描画差分を調査中

- PR #287 の remote HEAD は `30b180ed`。CI run `36856771139` の Preview fidelity が失敗。`05-preset-shapes.pptx` slide 1 の fg-SSIM が 0.7731 → 0.6161。Static、Node 22/26、OOXML validator は成功（Node 24 は確認時 pending）。
- 取得した画像 `/tmp/pptx-fidelity-30b/05-preset-shapes/slide-1.{ours,gt}.png` では、狭い図形の triangle / diamond / pentagon / star5 / leftRightArrow が ours で1行、LibreOffice で2行。直前の Latin 分割変更に対応する差。PowerPoint 一致の証拠として CI の LibreOffice 結果だけを使わない。基準値は変更していない。
- `latinLnBrk=true` で全幅より短い単語が残り幅で分割されない不具合を修正。Latin 分割を有効にしても日本語の句点は前の文字と保持する回帰も追加。false/省略時の全幅を超える単語の緊急折り返しは実機未確認のため変更せず、fidelity の失敗は未解決として残す。
- 今回の検証: 関連ユニット56件、全ユニット3089件成功（109 skipped）、root型検査・lint・対象format・dev依存込みbuild成功。ヘッダー日英900/1500px、フォント選択・Undo、透明な文字編集背景のブラウザー3件成功（`/tmp/pptx-header-current.log`）。
- 実機の再接続は `cgWindowNotFound`。文書変更・復元待ちはなし。`hangingPunct` は既定値の根拠のみ確認でき、文字ごとの位置計算は未確認のため未実装。
- 全ブラウザー実行 `70356` は終了（exit 1）。230件中228成功、フォント候補キー操作とノート splitter 高さの2件が失敗。旧 `16ae1bcc` ビルドに対する既知の失敗で、修正後の隔離テストは以前に成功済み。最新版一括成功とは扱わない。共有 dist の実行中制約は解除、以後のビルドは親担当で調整。
- 未コミットのアウトライン選択は修正中。クリックによる解除、連続入力後のフォーカス、後方選択、BR/ブロック境界のオフセットを追加レビュー。完成済みとして取り込まない。

## 2026-10-01: ノートの書式保持と Latin 単語の改行

- ノート欄の入力を範囲置換として記録し、編集していない run 書式・フィールドを保持。`setSlideNotes` に optional `range` / `preserveFormatting` を追加し、既定の全置換動作は維持。
- 正しい OOXML のフィクスチャで、旧 editor が italic run を消す失敗を再現。隔離ビルドでは離れた2か所の編集・保存・Undo と既存ノート欄操作が成功（`/tmp/pptx-notes-browser-red.log`、`/tmp/pptx-notes-browser-green.log`、`/tmp/pptx-notes-browser.log`）。ノートのリッチテキスト表示・書式ツールバー全体の対応ではない。
- SVG 文字配置へ `latinLnBrk` を伝播。Microsoft MS-OI29500 §2.1.1406 の Office 偏差に従い、省略時 false、明示 true の場合に長い Latin 単語を途中で分割。`eaLnBrk` は禁則処理を指し、CJK 全体の分割禁止という近似実装は採用しない。hangingPunct 対応も未実装。
- root 型検査、全ユニット3087件、Svelte0 errors/warnings、siteテスト、lint、対象ファイルのformat、隔離 core/preview/editor build 成功。共有distは旧ビルドで走る全ブラウザ検証のため上書きしていない。
- アウトライン範囲選択は引き続き作業中。forward copy/cut/Undo に加え、通常入力・paste・backward のブラウザ検証を要求している。完成済みとして扱わない。
- PowerPoint の最新接続試行も画面取得0×0で失敗。今回の実機変更・復元待ちはない。ユーザーのプレビュー資料とChromeタブは変更していない。

## 2026-10-01: フォントのキー操作・ノート高さ・白黒画像

- フォント検索欄から ArrowUp で末尾候補、ArrowDown で先頭候補へ移動するよう修正。ノート欄の最大高さと splitter の値を実際の表示に揃え、画面縮小にも追従。
- DrawingML biLevel の前に輝度変換を行い、飽和色や grayscale + duotone の組み合わせでも白黒化。閾値を32段階へ丸める近似を廃止。PowerPoint実機での色変換の完全一致は未確認。
- root 型検査、全ユニット3081件、site97件、Svelte0 errors/warnings、preview型検査・隔離ビルド成功。隔離したeditorビルドでフォント・ノートのブラウザー2件成功（`/tmp/pptx-font-notes-green.log`）。
- 全ブラウザー検証は旧ビルド16ae1bccに対して進行中（exec70356、`/tmp/pptx-full-browser-oct1.log`）。追加したフォント・ノート回帰は旧ビルドでは失敗する。実行中は共有distを上書きしない。最終版の全ブラウザー成功とは扱わない。
- アウトラインの複数フィールド選択は別担当が作業中。ブラウザー検証未完了のため、この変更のコミットには含めない。
- ユーザーのプレビュー資料は現在4スライドでユーザーの編集あり。既存Chromeタブに競合通知があり、勝手に再読み込み・競合解決しない。独立タブでヘッダー省スペース化を確認済み。

## 2026-10-01: 選択文字・テーマのフォント候補

- フォント候補に選択中の文字の Latin / East Asian / Complex Script とテーマの見出し・本文フォントを追加。独自名も選択肢に残る。OS の全フォント列挙ではない。
- `%` 付きの画像トリミング値を正しく読み込むよう修正（例: `-12.5%` は `-0.125`）。保存・再読み込み回帰を追加。
- format/lint/root・DSL型検査、全ユニット3076件、site96件、Svelte0 errors/warnings、core/editor build成功。ヘッダー日英900/1500px、フォント適用・保存・Undo、文字背景のブラウザー3件成功。
- クロススライド文字選択は引き続き未実装。共有アンカーと複数図形への一括編集の設計が必要。未完成の試作は採用していない。
- PowerPoint は起動中だが操作対象ウィンドウを取得できず、この回の実機変更・復元待ちはない。ユーザーの資料とタブは変更・再読み込みしていない。

## 2026-10-01: ヘッダー改善の再検証と画像の負値トリミング

- `24b84ffa` のヘッダー省スペース化・透明な文字編集背景・フォント選択は、独立ブラウザーテスト3件で再確認済み（日本語、900px/1500pxを含む）。ユーザー資料 `/tmp/pptx-user-preview/deck.tsx` は変更していない。
- `srcRect` の負値（画像外への拡張）を API・画像パネル・塗りつぶし復元・画像プレビューで保持するよう修正。保存再読込および画素検証を追加。全ユニット3075件、site 96件、format/lint/types/build、Svelte検証が通過。
- 小数点タブの区切り文字推定と段落禁則CSSの近似対応は、意味的な根拠が不足するため今回採用していない。未実装項目として残る。
- 全操作のPowerPoint一致は未完成。ヘッダー改善を全体完了と扱わない。

## 最新の目標とユーザー用プレビュー

- 完全一致まで継続する goal が有効。「OpenOffice 仕様」は **Office Open XML（OOXML）** の意味とユーザーが確認済み。OOXML で実現できない機能は対象外にできる。単に未実装・検証困難という理由で対象外にしない。
- ユーザー用プレビューは `http://127.0.0.1:4173/`。`/tmp/pptx-user-preview/deck.tsx` に review サンプルをコピーして起動。3 スライドとエディタの表示、page error なしを確認。Chrome に開いて Agents を閉じ、編集領域を広げた。
- サーバーの exec session は `84895`。ユーザーが編集するため、この資料・保存データをテストで上書きしない。再起動前にサーバーが生存しているか確認する。

## ユーザーの使用中フィードバックへの修正

- 編集テキストの白背景を透明化し、元の文字のみを一時的に隠して二重描画を防止。図形・表セルの塗りは保持。
- フォント欄に候補一覧・検索を追加。任意フォント名の直接入力も維持。OS のインストール済みフォント列挙ではなく共通候補一覧。
- 埋め込み時の重複 Undo/説明バーとブランド表示を除き、保存状態を TopBar に統合。外側ヘッダーは40pxに縮小。
- 独立ブラウザーテスト3件成功（背景・保存/再読込・セル、フォント部分選択/保存/Undo、日英1500/900pxヘッダー）。Svelte check 0 errors / 0 warnings、dev build成功。
- プレビューを再起動済み。ユーザー資料は上書きしていない。全操作のPowerPoint完全一致はまだ未完成。

## 最新: 3 系統の並列実装

- ユーザーがサブエージェントによる並列実装を明示的に依頼。利用上限の親 1 + 子 2 で、新規スライド、アウトライン移動、選択ウィンドウを分担した。今後も独立したファイル群を担当させ、共有 dist の build とブラウザ検証は親が直列に調整する。
- 下記の Title Slide 後の New Slide の未対応は解消。通常・アウトラインのメニューと Cmd/Ctrl+Shift+N が同じポリシーを使う。同じマスターの Title and Content を優先し、未使用レイアウトも公開 API で所属マスターを判定する。
- アウトラインの先頭/末尾で上下矢印を押すと隣のタイトル/本文へ移動し、入力中の変更を確定する。スライドをまたぐテキスト範囲選択は未実装。
- Selection Pane で複数の選択済み兄弟を一緒にドラッグできる。相対順・選択・グループ境界・一括 Undo を保持する。別担当の独立レビューで具体的な不具合は見つからなかった。
- PowerPoint の Arrange を開く際に画面取得 0×0 が再発。この試行で参照文書の内容は変更しておらず、復元待ちはない。複数図形ドラッグのネイティブ比較は未完了。
- 全操作一致は未完成。PR #287 のみを更新し、完全一致や根拠のない進捗率を報告しない。
- 検証: root format/lint/typecheck 成功、3,073 tests passed / 109 skipped、site 96/96、Svelte check、DSL typecheck、core/editor build 成功。ログは `/tmp/pptx-parallel-root-quality.log`、`/tmp/pptx-parallel-site-quality.log`。直前 HEAD `1d6927d4` の CI は全成功（run 36836258164）。
- 統合後の browser 9/9 成功: `new-slide-layout`、`outline-view`、`outline-slide-menu`、`selection-pane`。ログ `/tmp/pptx-parallel-browser.log`。最終 HEAD の全ブラウザ一括検証ではない。

## 接続復旧と New Slide 比較（2026-10-01）

- PowerPoint の完全終了後、CUA で起動・参照資料の読み込み・Layout メニュー操作が成功。直前の画面サイズ 0×0 エラーはこの比較では再発していない。
- `/private/tmp/pptx-outline-audit/body.pptx`（Normal、129%）を開き、Title and Content から Title Slide に一時変更して Cmd+Shift+N を実行。追加された第 2 スライドの Layout は **Title and Content** が選択されていた。
- Escape、Undo 2 回、保存で復元。1 スライド、Normal、129%、Undo disabled、Saved to my Mac を確認した。
- `ContextMenu.svelte` のアウトライン New slide は現在の layout をそのまま使うため、タイトルレイアウト後の例外は未対応。通常表示の同メニューも `addBlankSlide` を呼ぶ。修正時は同じマスター内の適切な本文レイアウトを選ぶこと。複数マスターを無視して資料内の最初の本文レイアウトを選ばない。
- `7b6d401c` の GitHub CI は Static、Node 22/24/26、Preview fidelity、OOXML、PR template が全て成功（run 36832667431）。全操作一致の完了を意味しない。

## 最新の追記: アウトライン入力中の継承書式

- 複数マスターの資料で、書式表示を有効にしたアウトラインの入力直後に 20px の文字が 10.6667px になる問題をブラウザーテストで再現。入力プレビューの段落・run 書式解決へ元の shape を `inheritanceSource` として渡して修正した。
- `outline-slide-menu.test.mjs` は独立した第 2 マスターを持つ資料を使い、自動保存前の表示、保存後の書式、Undo、New/Duplicate/Delete を英語・日本語で検証。2/2 通過。site ユニット 93/93、format/lint、Svelte 0 errors / 0 warnings、エディタ build が通過。
- ログ: `/tmp/pptx-outline-inheritance-before.log`（修正前の再現）、`/tmp/pptx-outline-inheritance-after.log`、`/tmp/pptx-outline-inheritance-quality.log`。最終 HEAD の全ブラウザー一括成功を意味しない。
- ユーザーの「開きました」後、PowerPoint `body` の画面取得に一度成功（Normal 160%、Undo disabled）。Layout を開く操作で再び `screen capture size invalid (0.0,0.0)` となり、全ウィンドウ一覧も取得不能。文書内容の変更はしていない。Title Slide 後の New Slide の例外確認は未完了。
- PR 全体レビューは進行中。比較基準 `origin/main`、751 ファイル / +85,338 / -2,663、テスト 251 ファイル。公開 API とグループ・リンク・ページサイズ変更を確認中であり、全ファイルのレビュー済みとは扱わない。

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

## 最新追記: アウトラインの階層表示

- textarea を共通 RichTextInput に置換し、段落レベル・箇条書き・連番を表示。書式非表示の実機で本文24px行高・約10pxレベル刻み・タイトル太字を確認。
- blur とコンポーネント破棄が同時発生しても編集を二重適用しないよう、未確定編集キューは同期的に空にする。
- 隔離ビルドで英日アウトライン12件、通常本文回帰12件、HTML解析1件が成功。追加の空段落・改行・日本語・文字列としてのHTML貼り付けも英日成功。通常ビルド全件結果は後続追記を参照。
- Show Formatting の切替で実機の文字サイズが変わることを確認したが、この機能とスライド間選択は未実装。参照は2回切替で元に戻し、Normal 120%、Undo disabled、保存済み。
- guides.ts の二重線形検索を Set/Map 化。既存ガイド6テスト成功。

- 本文・タイトルの右クリックにも既存の Collapse/Expand サブメニューを共通化して追加。全スライド対象・キーボード移動・Undo を英日 browser で検証。最新 root は3,072成功 / 109 skip、Svelteは0 errors / 0 warnings。

## 検証追記

- `4b61cb1e` の通常ビルド全ブラウザテストは223件すべて成功。CIも全件成功。階層表示の差分はこの全件実行より後なので、最新差分の対象テスト結果と区別する。
- Show Formatting の再監査ではアクセシビリティ経由の選択が安定せず、切替成功を再現できなかった。同名のOOXML属性はSlide Sorter用であり、アウトライン設定の保存方法は未確認。
- 実機アウトラインのCommand-Shift-Nで、Title and Contentの後に同じレイアウトの空スライドが入ることを保存XMLで確認。追加をUndoし、Normal 120%、1枚、Undo disabledに戻して保存済み。

- 階層表示・文字メニューの通常ビルド検証: アウトライン/クリップボード13件成功、format/lint/root型検査、DSL型検査、core/editorビルド成功。

## 最新追記: 本文メニューのスライド操作

- アウトライン本文から追加・複製・削除を実装。追加は選択スライドのレイアウトを継承し、複製は本文以外の追加オブジェクトも保持。各操作と Undo、保存結果を英日2件で確認。
- Title and Content は実機と一致。Title Slide の後の特殊なレイアウト選択は未検証。

### 2026-10-01 アウトライン書式表示

Show Formatting の実機切り替えを再現し、英日メニューとセッション内表示設定を追加。表示切り替えで文書の Undo・保存を発生させない。太字・斜体などは継承書式を解決して描画し、編集保存時は元の文字サイズを維持する。実機の一時太字を Undo し、Normal / 120% / 保存済み / Undo 無効へ復元した。

書式付きアウトラインの文字サイズは現時点では 1/4 スケール近似。サイズ上限、正確な行高、ネイティブ初期値・設定の持続範囲は未検証なので、完全一致と扱わない。クロススライド選択、残るメニュー、全 PR レビューも未完了。

pnpm 12.5.1 の Node エントリは `bin/pnpm.mjs`（`.cjs` ではない）。依存再配置が発生して取得先の DNS で停止したため、`install --frozen-lockfile --registry=https://registry.npmjs.org` で復元。ロックファイル変更なし。
