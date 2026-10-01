## 2026-10-02: キャンバスの縦横比固定

- 検証: 修正前のブラウザー回帰で比率2→2.263851へ崩れることを再現。修正後のresize-geometry 2件、動画リボン1件、site121件、Svelte 0 errors/warnings、editor build成功。Undo後の再選択をテストに追加。全体型検査で前回のCommandDoc変更に対するFakeDocのsetDocumentSetting欠落を検出し、テスト用クラスも追従。
- Mac PowerPointの参照動画で実測: 固定ONの右下ドラッグで幅10.16→11.93cm・高さ5.72→6.71cm。右辺ドラッグでは幅だけ11.93cmへ変化し高さ5.72cmを維持。固定OFFの右下ドラッグでは幅11.93cm・高さ6.00cmになり自由変形。
- キャンバスはこれまでShiftしか見ていなかったため、ドラッグ開始時に保存済みの縦横比固定を読み、角ハンドルだけに適用する修正を追加。
- 参照資料 `/tmp/pptx-poster-audit/trim-reset.pptx` は全リサイズをUndoし、固定ONに戻して保存。XMLで位置914400/914400、寸法3657600/2057400、noChangeAspect=1を確認。未復元の変更なし。
- Shift併用時のPowerPoint実測、古いUndo/Redo履歴をまたぐ縦横比設定、動画書式の未実装項目は引き続き必要。全体の完全一致は未達成。

## 2026-10-02: 縦横比固定の保存

- 検証: core全体3190 passed / 109 skipped、site121 passed、Svelte 0 errors/warnings、format/lint/typecheck、core/editor build成功。動画リボン・compact-editor・resize-geometryブラウザーテスト成功。

- `isShapeAspectRatioLocked` / `setShapeAspectRatioLocked`でDrawingMLの`noChangeAspect`を読み書きし、動画リボンとサイズパネルを文書の設定に連動。固定解除後の幅変更では高さを維持する。
- Mac実機の`/tmp/pptx-poster-audit/trim-reset.pptx`でチェック解除→保存を確認: `a:picLocks`の`noChangeAspect`属性が消え、チェックを戻すと`1`が復元。両操作でUndoは無効のまま。チェックを元に戻して保存・XML確認済み。
- 設定変更自体はUndo項目を増やさず現在の履歴スナップショットを更新。次のサイズ変更のUndo後も設定を保持する。より古いUndo/Redo履歴をまたぐ設定の挙動、ドラッグハンドルとShift併用時の実機一致は追加確認が必要。

## 2026-10-02: 動画書式リボンの専用化

- 検証: 専用リボンの実再生、サイズ保存・縦横比維持・Undo、メニューのキーボード操作のブラウザーテスト成功。既存のヘッダー、フォント、文字編集背景、表紙画像の4件も成功。site120件、format/lint/typecheck、Svelteチェック、editor build成功。

- 実機で Video Format の順序を確認: Preview、Adjust (Corrections / Color / Poster Frame / Reset)、Video Styles、Alt Text、Arrange、Size、Format Pane。従来実装はタブ名だけ変更して図形用コマンドを表示していた。
- 専用リボンにPlay/Pause、Poster Frame、Video Border / Effects、Alt Text、Arrange、直接Height / Width、Format Paneを実装。完全一致にはギャラリー、Corrections / Color / Crop / 独立Reset、配置操作の個別ボタンなどが残る。縦横比固定は既存SizePositionSectionと同様ローカル状態で、OOXML noChangeAspectとの同期は未実装。
- トリム後Poster Frame Resetの実験資料は `/tmp/pptx-poster-audit/trim-{before,after}-reset.pptx`。開始100msのOOXML保存を確認したが、Reset後のPNGが汎用再生アイコンで、動画の先頭/トリム開始どちらのフレームかを判定できなかった。この実験をReset仕様の根拠には使わない。比較元 `trim-reset.pptx` はUndoを2回行い保存し、Undo無効・trim要素消滅を確認済み。

## 2026-10-02: 表紙画像のリセットと省スペース表示の再確認

- Mac PowerPoint の Poster Frame > Reset は先頭フレームで表紙PNGを置換し、動画本体・再生設定を保持する。`/private/tmp/pptx-poster-audit/{before-reset,after-reset}.pptx` が比較資料。Undoして保存済み。
- 同じ操作をエディターに追加。独立した動画デコーダーで先頭フレームを取得し、選択や文書の変更があれば結果を破棄する。既存の setShapeImage とUndo履歴を利用。
- 検証: Resetブラウザー1件（先頭フレーム画像・動画バイト列・再生設定・非ゼロ停止位置・Undo/Redo・再読込・日本語メニュー）成功。site120件、Svelte 0 errors/warnings、format/lint、editor build成功。
- 未完了: トリム済み動画でのResetの実機比較、Video Format全体の配置/効果の一致。全体の完全一致は未達成。
- ヘッダー、フォント選択、文字編集背景のブラウザー3件再検証成功。独立した確認プレビューは http://127.0.0.1:4175 （`/tmp/pptx-header-preview/deck.tsx`）。元の4173の編集データは変更していない。

## 2026-10-02: 動画の表紙画像

- Video Format > Poster Frame から Current Frame / Image from File... を実装。既存 setShapeImage を利用し動画本体・再生設定を保持。再生/シークまで動画要素を隠して表紙画像を表示。
- 実機 `/private/tmp/pptx-poster-audit/deck.pptx` を開き、Mac の項目名と Current Frame が選択直後に無効・再生後に有効になることを確認。メニューには Reset も存在するが未実装。リボン全体のVideo Format配置/スタイルも未一致。実機資料の内容変更はなし（Undo無効を確認）。
- ブラウザーで表紙の永続化、動画バイト列/再生設定保持、Undo、再読込、日英メニューと初期非表示を検証。format/lint/root typecheck、Svelte 0 errors/warnings、site120件、editor build成功。
- 次はPoster Frame > Reset の実機動作・保存OOXMLを確認して実装。その後、Video Formatの動画専用レイアウトと効果の差分を進める。全体の完全一致は未達成。4173の未保存編集は変更していない。

## 2026-10-02: 回転時の再生バーとヘッダー確認

- Mac PowerPointで音声を90度回転しても再生バーは水平で図形の下に配置されることを確認。一時変更はUndo済み。office-kitの再現テスト失敗後、バーをキャンバス直下に配置して回転から独立させた。
- ブラウザー5件成功（`/tmp/pptx-transform-browser.log`）：回転動画の実再生/デコード、音声ブックマーク/トリム、動画切替/端の配置、ヘッダー日英900/1500px。グループ子メディア選択の比較は未完了。
- Svelte 0 errors/warnings、format/lint/root typecheck成功。core変更なし。
- 確認用4174をChromeの新しいタブで開いた。省スペース化は反映済み。4173の未保存編集は変更していない。
- 全体の完全一致は未達成。次候補はVideo Poster FrameのUI（既存setShapeImageを再利用可能）。

## 2026-10-02: メディア再生バーの外観と保存操作

- 実機で確認した38px高の再生バー、塗りつぶしシーク軌道、SVG再生/移動/ミュートへ変更。インライン音量スライダーを除去し、音量は再生リボンから変更する。小さな図形で左側ナビに隠れてクリックできない問題をキャンバス内への位置補正で修正。
- Save Media As を再生リボンに追加。埋め込み音声・動画の元データとファイル名を保存。オンラインメディアでは無効。
- 音声の実再生、トリム/ループ/フェード、ミュートUndo、動画の実デコードフレームとスライド切替時停止、右端/画面幅変更時のバー位置、保存したバイト列の一致をブラウザー検証。回転/グループ内メディアのバー配置は未検証。
- 検証: 統合ブラウザー6件成功 `/tmp/pptx-controls-final-browser.log`、site120件、Svelte 0 errors/warnings、format/lint/root typecheck/dev build成功。core変更なし（直前の3186 passedを継承、今回は再実行なし）。
- 全体の完全一致は未達成。PR287に統合。4173の未保存編集には触れない。実機文書の変更・復元待ちはない。

## 2026-10-02: インライン再生・ブックマークUIの統合

- 音声・動画の選択時に再生バーを表示。±250ms、ブックマーク追加/削除、音量・ミュートの保存、選択変更時停止を実装。トリム終端、巻戻し、ループ、フェードを再生に反映。
- core 3186 passed /109 skipped、site 120 passed、format/lint/typecheck、Svelte 0 errors/warnings、core/dev build成功。ブラウザーはブックマーク保存/再読込/削除、実再生1秒超、旧要素停止、±250ms、Mute/Undoを検証。追加の専用テストでトリム終端/巻戻し、リボン再開、ループ、時刻固定のフェード音量、音量保存/Undoが成功。ログ `/tmp/pptx-inline-final-browser.log` のリボン非同期イベント待ちを修正後、対象再実行 `/tmp/pptx-inline-trim-browser.log` が成功。
- ヘッダーの日英900/1500px回帰成功 `/tmp/pptx-inline-compact-final.log`。確認用4174はHTTP200。4173の未保存編集は変更していない。
- 全体の完全一致は未達成。インライン音量スライダーの常時表示など外観差、動画特有の再生/媒体差替えのブラウザー回帰は今後の確認対象。PRは287に集約。CI結果は未確認。

## 2026-10-02: インライン再生ボタンの実機差分

- Bookmark core は `8de023c4` として PR287 にpush済み。下記の「未コミットcore」はこのコミットで解消。UIは引き続き検証中。
- Mac PowerPoint の専用 `converted-auto.pptx` で、ブックマークなしの状態でも前後ボタンが使えることを確認。前進1回で0:06.74→0:06.99、後退1回で0:06.99→0:06.74。これはブックマーク移動ではなく±0.25秒のシーク。並びは再生・シークバー・後退・前進・時刻・音量。バーは明るいグレー、アイコンは黒。
- インラインのスピーカークリックはミュート切替。Undo後に「Redo Media Volume」と表示されるため文書の編集として扱われる。確認用変更はUndo済み（Undo無効を確認）、復元待ちはない。
- UIレビューで、XML順と時刻順のbookmark index混在、古いplay Promiseのreject、再生後巻戻し、同一IDでのメディア差替え時URL更新を指摘。担当agentへ修正・回帰テストを依頼。未検証のUIを完成扱いしない。

## 2026-10-02: ブックマーク統合のレビュー継続

- `b7925a4c`（トリム時刻精度）はPR287の `feat/pptx-editor` へpush済み。PRはOPEN。CI成功は未確認。
- 未コミットのbookmark coreは、name+time/name/time/未使用ノードの順で全体を割り当て、先頭追加で既存メタデータを奪わない実装へ修正。参照中のbookmark削除・名前変更を変更前に拒否。検証はcore全3186 passed /109 skipped、対象54件、format/lint/typecheck/build成功。参照保護fixtureは既存timingのcondに配置し、timingが1個・ID重複なしを検証。
- 未コミットのinline preview UIはレビュー修正中。エディターごとの状態分離、選択変更時停止、正確な時刻、シークバー内のマーカー、トリム/フェード/音量、失敗表示を検証してからコミットする。ブラウザーテストは有効なWAVを使い、保存・再読込・Undo・小数時刻・選択変更を確認すること。未検証のUIを完成扱いしない。
- ヘッダー省スペース化は実装・日英900/1500px検証済み。4173の未保存編集は再読み込みしない。PowerPointの一時変更の復元待ちはない。

## 2026-10-02: トリム精度とブックマーク実測

- トリムスライダーの step=50 による既存時刻の丸めを修正。step=any で小数時刻と非50ms倍数の終端を保持し、ドラッグ・矢印キーだけ50ms移動。Home/End、逆向きフェードアウトキー、反対側端点のクランプも検証。
- 5033ms音声のブラウザー回帰で修正前の失敗を確認。修正後は500.25ms開始・250.125msフェードの保存/再読込/Undo、全ハンドルのドラッグ、再生終了・再開、日英ヘッダーが成功。ログ `/tmp/pptx-trim-precision-before.log`、`/tmp/pptx-precision-final-browser.log`。dev build/typecheck、Svelte0 errors/warningsも成功。
- PowerPoint接続使用可。専用 `/tmp/pptx-audio-across-audit/converted-auto.pptx` でAdd Bookmarkを実測。2件はname="Bookmark 1" time="0"、name="Bookmark 2" time="6747.0924"。画面時刻0:06.74。選択中は黄色丸、未選択は白丸。Add/Removeの有効状態も記録。2回Undo・保存でブックマークなしに復元しXMLも確認。保留の一時変更なし。
- Bookmark core とインラインプレビューUIは別agentが継続中。未コミット差分を消さず、coreの参照保護・並べ替えテストとUI検証を完了してから統合する。

## 2026-10-02: 統合トリムタイムライン

- 開始・終端・再生位置・フェードを共通時間軸に配置。黄色のトリム、白のフェード、青の再生位置ハンドルをドラッグ/キーボード操作できる。実音声をオフラインでデコードして波形表示。対応しないコンテナは波形不可の表示にする。
- 検証: 全5ハンドルのドラッグ、50msキー操作、キャンセル、保存・再表示・Undo、再生終端停止が成功。幅の回帰も追加。画像 `/tmp/pptx-trim-timeline.png` を目視確認。Svelte 0 errors/warnings、format/lint、依存順dev build成功。ログ `/tmp/pptx-timeline-browser-final2.log`。
- ヘッダーの日英900/1500pxブラウザー検証も成功（`/tmp/pptx-timeline-browser-final.log`）。Studio40px、保存状態同一行、重複バー除去、リボン折りたたみは実装済み。元の4173は再読み込みしていない。
- 厳密なネイティブ外観比較・全UI/操作一致は未完了。PowerPointファイル選択画面の未解決状態は下記追記を参照。次回は残るメディア開始方法・ブックマークと全体レビューを継続。PR287に統合する。

## 2026-10-02: トリム画面の再生境界

- トリム画面のプレビュー位置・フェードを再生中25ms間隔で更新し、終端を超えた際は実際のメディア再生位置も終端へ戻して停止。停止/閉じる際はタイマー解放。再再生はトリム開始位置から。音符プレースホルダーをスピーカーSVGへ変更。
- 検証: ブラウザーで4秒終端への停止・開始位置からの再再生・保存/再表示/Undo成功。依存順dev build、Svelte0 errors/warnings、format成功。ログ `/tmp/pptx-trim-preview-test.log`。
- Macのファイル選択Go to Folder画面でキー/AX/座標操作が反映されず、今回は専用資料も開けなかった。文書編集なし。ダイアログが残っている可能性あり。次回接続時は閉じてから比較を再開。トリムの統合ハンドル・波形は未実装。
- ec34e05fのCIはStatic checks成功、その他5件は確認時pending。全体一致は未達成。

## 2026-10-02: メディアのトリム・フェード

- `MediaPlayback.trim` / `fade` と再生リボン、トリム画面を追加。`trim.endMs` は再生終了位置ではなく末尾から除去する時間。未知のXML属性・子要素を保持し、保存・再表示・Undoに対応。
- Mac PowerPointで開始/末尾トリムとフェードを各50msにした実XMLを確認。専用 `converted-auto.pptx` は3回Undoしてすべて0、長さ10秒に復元・保存・閉じた。復元待ちなし。実測XMLは `/tmp/pptx-native-trim-fade.xml`。
- スライドショーでトリム境界・フェード・区間ループ・巻き戻しを反映。自然終了時のフォールバックと発表者ミラーも対応。
- 検証: core全3176 passed /109 skipped後、追加ケースを含むメディア44件成功。再生リボン6件、トリム画面1件（保存・キャンセル・再表示・Undo）、ランタイムのトリム/フェードと自然終了/巻き戻し2件成功。format/lint/types、依存順dev build、Svelte 0 errors/warnings。ログ `/tmp/pptx-trim-core-final.log`、`/tmp/pptx-trim-dialog-final2.log`、`/tmp/pptx-trim-ui.log`（トリム初回失敗はfinal2で修正確認）。
- 未完了: トリム画面の波形・統合ハンドル等の厳密な外観一致、In Click Sequence、ブックマーク、複雑なタイミング、全UI/操作一致。現在のトリム画面は個別スライダー。全体完成とは扱わない。
- PR287に統合。4173の未保存編集は触らない。ヘッダー省スペース化は日英900/1500pxで再検証済み。確認用コピーは4174。

## 2026-10-02: 音量メニューの実機照合

- 再生リボンの音量を数値入力から「小・中・大・ミュート」のメニューへ変更。詳細パネルの任意音量指定は維持。
- 専用 `converted-auto.pptx` をMac PowerPointで操作・保存して確認: Low=20000、Mediumはvol省略（スキーマ既定50000）、High=80000。Muteは既存volを保持してmute=1。チェック済みMuteの再選択でも解除されず、音量レベルを選ぶとmuteを除去する。確認後は元のHigh・ミュートなしに戻して保存し、閉じた。
- 76810d7aのCIでStatic checks、OOXML validator、Preview fidelity、Node22/26が成功。以前のLatin折返しfidelity差分は解消。Node24のブラウザーテストは確認時点で実行中。
- 検証: format/lint/typecheck、依存順dev build、Svelte 0 errors/warnings成功。音量保存・Undo・再読込・日本語・Escape・ヘッダーを含むブラウザー7件成功（`/tmp/pptx-volume-compact-final.log`）。
- ヘッダーの変更と同じPR287に統合。4173の未保存編集は触らない。全操作の完全一致は引き続き未完了。

## 2026-10-02: ヘッダー確定・ネイティブ音声の開始方法変換

- ヘッダー省スペース化を日英・900/1500pxで再検証。Studioバー40px、保存状態を同じ行、重複ツールバーなし、リボンを閉じると60px以上増える。ログ `/tmp/pptx-compact-oct2-final.log`。
- ネイティブ背景音声の開始方法を「自動」「クリック時」で相互変換。PowerPointが保存した専用seqの完全一致を条件にし、外部参照や共有アニメーションは変更前に拒否。IDとメディア設定を保持。開始遅延の変更は未対応。
- `/tmp/pptx-audio-across-audit/converted-click.pptx` と `converted-auto.pptx` をMac PowerPointで開き、修復要求なし、StartがWhen Clicked On / Automaticallyであることを確認。ループ・スライド跨ぎ・非表示も保持。両方とも編集せず、保存要求なしで閉じた。
- Mac PowerPointのpreset-shapesで、latinLnBrk省略でも図形より長い英単語が折り返されることを確認。SVGにも緊急折返しを適用。通常の残り行幅での単語分割はフラグに従う。fidelity基準値は下げていない。LibreOfficeが応答せずローカル画像比較は未完了、CIで確認する。
- 検証: core全3169 passed /109 skipped、最終音声28件、format/lint/typecheck、依存順dev build、preview/DSL/dev typecheck、Svelte 0 errors/warnings。ブラウザー7件成功後、最終変換の保存・再読込・Undoを再実行して成功。ログ `/tmp/pptx-native-convert-unit.log`、`/tmp/pptx-native-convert-target-final.log`、`/tmp/pptx-native-convert-browser-final.log`。
- 未完了: In Click Sequence、トリム/フェード/ブックマーク、共有された複雑なタイミング、全UI/操作の完全一致。PR287に統合。4173のユーザープレビューは未保存編集があるため変更・再読込しない。

## 2026-10-02: 背景音声とヘッダーの再検証

- ヘッダーはStudioバー40px、保存状態を同じ行へ統合、重複ツールバーを非表示、リボン折りたたみに対応済み。compact-editorの900/1500px・日英検証が成功。
- 音声の「バックグラウンドで再生」を追加。自動開始・999枚・ループ・非表示を一度のUndo単位で設定し、音量・巻き戻しは保持。音声のHide During Showは再生中も非表示（ブラウザーによる自動再生拒否時は再試行UIを表示）。
- Mac PowerPointの専用サンプルでPlay in Backgroundの4設定、既存Rewind ONの保持、スライドショーで音声アイコン非表示を確認。比較中の一時変更2件はUndoし、Undo無効・各設定OFFを確認して保存要求なしで閉じた。ユーザー文書の変更なし。
- ネイティブ保存の背景音声はmedia自身のdelay=indefiniteのまま、mainSeqのonBeginからafterEffect/playFrom(0.0)で開始する。最小XML fixtureを追加して読み込みに対応。複雑なコマンド開始条件は未対応。ネイティブ背景音声の開始方法・遅延変更は、コマンドを壊さないよう原子的に拒否。No StyleのIn Click Sequenceは未実装であり、手動開始で代用しない。
- 検証: format/lint/root・dev typecheck、core 3166 passed /109 skipped、依存順dev build、Svelte 0 errors/warningsが成功。編集UI・ヘッダー・音声ブラウザー9件成功。ネイティブfixtureを含む音声4件も追加検証。ログは `/tmp/pptx-background-unit.log`、`/tmp/pptx-background-browser.log`、`/tmp/pptx-background-native-final.log`。
- 全操作の完全一致は未達。Start選択肢・トリム/フェード/ブックマーク・複雑なタイミング・Latin折返しfidelity差分などが残る。PR287に統合を継続。4173のユーザープレビューには未保存編集があるため変更・再読込しない。

## 2026-10-02: 再生終了後の巻き戻し

- 音声・動画の `MediaPlayback.rewindAfterPlaying` と再生リボンの「再生が終了したら巻き戻す」を追加。ONはmediaのcTn@fill=remove、OFFはhold。入れ子のタイミングと無関係な属性を保持し、保存・Undo・再読込に対応。
- Mac PowerPointで `/tmp/pptx-audio-across-audit/deck.pptx` のRewind After Playingを切り替えて保存し、上記XML差分を実測。OFFへ戻して保存し、閉じた。ユーザー資料の変更・復元待ちはなし。
- プレビューは観客側の自然終了時だけ先頭へ戻して停止。発表者ミラーは観客側の同期に従う。OFFの動画は終端を保持。
- 旧UIの回帰失敗を確認 `/tmp/pptx-rewind-ui-before.log`。UI6/6成功（音声/動画・直接/入れ子・Undo/保存/再読込/日本語・ヘッダー）、Svelte0 errors/warnings。core全3160 passed /109 skipped、format/lint/types/build成功。ログ `/tmp/pptx-rewind-tests.log`、`/tmp/pptx-rewind-ui-final.log`。最終runtime検証は `/tmp/pptx-rewind-final-runtime.log`、依存順ビルドは `/tmp/pptx-rewind-final-build.log`。
- 未完了: ネイティブの再生位置・最終フレームの厳密な比較、Start追加選択肢、Play in Background、トリム・フェード・ブックマーク、複雑な開始条件、Latin折返しfidelity差分など。全UI・全操作の完全一致は未達成。PRは287に統合し続ける。

## 2026-10-02: 複数スライド音声の再生リボンと実機確認

- 音声の再生リボンに Play Across Slides / スライド切り替え後も再生を追加。ONで999枚、OFFで既定の1枚。既存の有限範囲はチェック状態として読み取り、他のオプションを操作しても保持。詳細の再生オプションでもslideCountを編集可能。
- Mac PowerPointの接続が復旧。以前の `pptx-native-wrap-oct2` アクセス要求2件をキャンセルし、専用の `/tmp/pptx-audio-across-audit/deck.tsx` をDSLで生成して比較。PowerPointのPlaybackタブのチェックを入れて保存すると `p:cMediaNode vol="80000" numSld="999"`、外して保存するとnumSld属性なしとなることをZIP内XMLで確認。サンプルはOFFで保存して閉じた。ユーザー文書の変更なし。
- 旧ビルドでチェックボックスが存在せず回帰テスト失敗を確認（`/tmp/pptx-across-ui-regression.log`）。format/lint、Svelte 0 errors/warnings、site117件、依存順dev build成功。ブラウザー6/6成功（直接/入れ子の音声・動画、複数スライド設定のUndo/保存/再読込/日本語/動画非表示、ヘッダー）。最終ログは `/tmp/pptx-across-ui-final.log`。
- 未完了: 実機のStartにある追加選択肢、Rewind After Playing、Play in Background、トリム・フェード・ブックマーク、入れ子の複雑な開始条件、Latin単語折返しのfidelity差分など。接続が使える間に実機比較を進める。全操作の完全一致は未達成。

## 2026-10-02: スライドをまたぐ音声再生

- 再生中の音声を `slideCount` の範囲内で保持し、範囲外・元スライドへの戻り・ショー終了・メディア更新時に破棄する。同じshapeIdを持つ次スライドのメディアとは分離し、発表者ビューの古いコマンドを世代キーで拒否。
- 旧実装でスライド移動時の停止を再現（`/tmp/pptx-across-before.log`）。音声、遅延、インライン動画、全画面動画、発表者ビューのブラウザー4件成功（`/tmp/pptx-across-browser.log`）。追加した古いコマンドの回帰も成功（`/tmp/pptx-across-final-browser.log`）。
- format/lint/root・dev typecheck/test/build成功。3158 passed / 109 skipped。依存順dev buildも成功。ヘッダー日英900/1500pxの再検証も成功。
- 未完了: 再生リボンへの複数スライド設定、動画の複数スライド挙動、途中へのジャンプ・戻り・custom showの実機比較、背景音声の発表者ビュー操作。現在の保持範囲はshowCursorと起点の差で判定。完全一致とはしない。
- 元プレビュー4173は変更せず、確認コピー4174をChromeで開いた。保存競合の通知は保持。既存のPreview fidelity差分とPowerPoint実機比較も残る。

## 2026-10-02: メディアの再生スライド数

- `MediaPlayback.slideCount` で `cMediaNode@numSld` を読み書き。省略時はOOXML既定の1枚。unsignedIntの0〜0xffffffffを受け付け、範囲外は音量など他の同時変更も適用する前に拒否。保存再読込・既定値への復元・schema境界値を検証。
- 根拠: references/ecma-376-5th/ECMA-376/OfficeOpenXML-XMLSchema-Transitional/pml.xsd:601、および https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commonmedianode 。
- format/lint/typecheck/test/build成功。3158 passed / 109 skipped。ヘッダーの既存ブラウザー検証も日英900/1500pxで再実行成功。
- 実際のスライド跨ぎ再生は未実装。`page.ts` のrender/syncMediaPlayerがスライド変更時にdisposeする。次は再生要素の保持、再生範囲終了・ショー終了時の破棄、同じshapeIdを持つ別スライドの区別、発表者ビュー同期を一体で実装する。API対応だけを再生機能完成としない。
- PowerPoint AppleScriptのバージョン取得は16.113.3で成功。一時コピー `/tmp/pptx-native-wrap-oct2/preset-shapes.pptx` のPDF書出しは2回ともタイムアウト。比較用コピーは保存せず閉じた。ユーザー資料の変更・復元待ちはなし。latinLnBrkの描画差分は未解決。
- CI 36897221430は静的検査、Node22/26、Open XML SDK成功。最終確認時Node24実行中、Preview fidelity失敗。完全一致は未達成。

## 2026-10-02: 入れ子メディアの再生設定

- 入れ子のaudio/videoを再生リボンと既存APIで認識。単純な並列グループの親遅延を合算し、総遅延の編集ではメディア自身の遅延のみ変更する。音量・ミュート・繰り返し・停止時非表示は親条件を保持。
- seq/excl/subTnLst、masterRel、参照開始条件はスライド開始時の自動再生と誤判定しない。対応できない開始変更や重複メディア対象は変更前に拒否する。これらのタイミングを実際に再生・編集する完全対応は残件。
- 非有限の音量入力も変更前に拒否。CIのNodeテストで不足していたxmllintをインストールするよう変更。既存Preview fidelity失敗（05-preset-shapes、slide1）は未解決で、基準値は変更していない。
- 回帰再現: `/tmp/pptx-nested-browser-before.log`、`/tmp/pptx-media-volume-before.log`、`/tmp/pptx-subordinate-before.log`、`/tmp/pptx-media-reference-before.log`。最終対象ユニット46件成功。
- 最終ブラウザー4/4成功（`/tmp/pptx-nested-final-browser.log`）。直接/入れ子の音声・動画についてUndo、保存、再読込、日英表示を確認。
- 最終format/lint/typecheck/test/buildと依存順dev build成功。3150 passed / 109 skipped（`/tmp/pptx-nested-gates-final.log`、`/tmp/pptx-nested-build-final.log`）。
- PowerPoint実機は再試行したがscreen capture 0×0で取得失敗。実機の変更なし、復元待ちなし。元のユーザープレビューも変更・再読み込みしていない。全体の完全一致は未達成。

## 2026-10-02: アニメーション削除時の入れ子メディア保持

- 最終format/lint/root typecheck/test/build成功。3128 passed / 109 skipped、依存順dev build成功。ログ `/tmp/pptx-clear-nested-gates.log`、`/tmp/pptx-clear-nested-tests.log`、`/tmp/pptx-clear-nested-dev-build.log`。
- `clearSlideAnimations` がルート直下以外のaudio/video timingを消す不具合を修正。メディアの祖先と開始・終了条件、iterate等を保持し、図形効果とbldLstを除去。新しいルートへの移動やIDの再割当てはしない。
- 保持した条件が削除対象のcTnを参照する場合は変更前に拒否。依存関係を保って効果を削除する完全対応は残件であり、成功扱いにしない。
- 実メディア/図形を使うchildTnLst・subTnLst、endSync・iterate、保存再読込、schema/validator、拒否時の非変更をテスト。初期回帰失敗と修正途中の消失を確認（`/tmp/pptx-clear-nested-current.log`）、最終対象47件成功（`/tmp/pptx-clear-nested-target.log`）。
- 根拠: childTnLstの開始は親に相対的（https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.childtimenodelist?view=openxml-2.19.0）。subTnLstのmasterRelは別の意味を持つため構造を維持。実機取得は今回も0×0で失敗し、Macの挙動完全一致は未検証。
- PR287は開始時HEAD34aac056、OPEN、statusCheckRollup空。全体UI/全操作の一致・既存fidelity差分は引き続き未完了。ユーザーの元プレビューや資料は変更していない。

## 2026-10-02: メディア選択時の再生リボン

- 単一のメディアを選択した場合に「再生」コンテキストタブを表示。自動/クリック開始、ループ、音量%、ミュート、動画の全画面、停止中非表示を直接設定。自動開始時は別Timingグループで秒単位の遅延を編集。既存API経由のtransactでUndo/保存対応。
- 1.001秒などの浮動小数点誤差はミリ秒変換時に丸め、入力のstep検証で桁を制約。汎用設定画面にもdelayMsを追加。日英翻訳を追加。
- 修正前は再生タブがなく音声・動画テスト2件失敗（`/tmp/pptx-media-editor-before.log`）。修正後2件＋既存compactヘッダー1件成功。保存、Undo、再読込、1001ms遅延、音量25%、ミュート、非表示、動画のみ全画面、日本語の開始操作を確認。site117件成功、Svelte0 errors/warnings、dev build成功。
- 日本語動画リボン画像 `/tmp/pptx-media-editor-video.png` を目視確認。PowerPoint実機は今回も0×0で取得失敗。Microsoftの操作説明を参照: https://support.microsoft.com/en-us/powerpoint/insert-and-play-a-video-file-from-your-computer 。ネイティブ版との見た目完全一致は未検証。
- 最終format/lint/root typecheck/test/build成功。core 3125 passed / 109 skipped。ログ `/tmp/pptx-media-editor-gates.log`、`/tmp/pptx-media-editor-tests.log`。
- 全体の完全一致は未完了。クリックシーケンス、入れ子タイミング、numSld、オンライン再生、従来fidelity差分などは残件。ユーザーの既存プレビューや編集資料は変更していない。

## 2026-10-02: 自動メディアの開始遅延

- `MediaPlayback.delayMs` を追加。イベントなしの自動開始条件の正の遅延を読み込み、保存・再読込で保持。ゼロは従来の戻り値形状を維持。setterは非負safe integerのみ許容し、非自動再生へのdelay単独指定やautoplay:false併用を変更前に拒否。
- プレゼン再生で指定時間を待機。手動再生、pauseコマンド、disposeで予約を解除。大きな遅延はブラウザーの32bit timeout制限を超えない分割予約。発表者ミラーは独自に予約せず観客側に同期。
- 修正前は1.2秒遅延を設定しても300ms時点で再生済みとなる回帰失敗を確認（`/tmp/pptx-delay-before.log`）。修正後ブラウザー音声/動画2件、発表者通常/全画面2件成功。終了後の予約解除、手動再生後のpause、再訪時のリセットも検証。
- format/lint/root・dev型検査、core全3124 passed / 109 skipped、依存順dev build成功。ログ `/tmp/pptx-delay-gates.log`、`/tmp/pptx-delay-tests.log`、`/tmp/pptx-delay-browser.log`、`/tmp/pptx-delay-presenter.log`。
- レビュー後、delay変更時に追加開始条件が消える問題を回帰失敗で再現し、既存の自動開始条件だけ更新するよう修正。最終対象24件、lint/types/core build成功。
- 全UI・全操作の一致は未完了。クリックシーケンスと親タイミング条件、numSld、オンライン再生、PowerPoint実機比較、既存fidelity差分は残件。ユーザーの編集中プレビューは変更していない。

## 2026-10-02: メディア開始条件と削除時のタイミング保持

- `evt="onClick" delay="0"` が自動再生と判定される不具合を回帰テストで再現し、イベント条件を自動再生から除外。OOXML の delay はイベント発生後の遅延であり、ゼロでもクリック待ちを意味する。根拠: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.condition?view=openxml-3.0.1 。
- 入れ子のメディア timing 削除を修正。childTnLst/subTnLst/tnLst 内の対象を削除し、空になったリストのみ除去。親の開始条件、別のルート、対象外のメディアを保持。入れ子の開始条件を自動再生 boolean へ単純に変換すると親のクリック待ちを失うため、開始条件モデルとクリックシーケンスの実装は残件。
- 検証: 修正前 helper は新規5件中4件失敗（`/tmp/pptx-media-timing-before.log`）。修正後関連50件、全体3114 passed / 109 skipped、format/lint/types/core build成功。入れ子削除はXSD検証も実行。全体ログ `/tmp/pptx-media-timing-full.log`。依存順dev buildも成功。レビューで指摘された最後のtop-level media削除について追加XSDテストを実行し、空のtimingはCT_SlideTimingの全子要素optionalにより有効と確認（新規6件通過）。ブラウザー音声・動画再生各1件成功（`/tmp/pptx-media-timing-browser.log`、`/tmp/pptx-media-timing-video.log`）、dev型検査成功。
- PowerPoint 実機は今回も画面サイズ0×0で取得失敗。ネイティブ文書・設定の変更なし。全操作・UI完全一致は未完了。

## 2026-10-02: UI再確認と発表者メディア操作

- ユーザーが指摘したヘッダーの省スペース化・文字編集の透明背景・フォント候補は既存実装済み。最新dev buildで3ブラウザーテスト成功（`/tmp/pptx-ui-current.log`）。ヘッダーは日英900/1500pxで検証。元のプレビューは再読み込みしていない。
- 確認用コピーを `/tmp/pptx-ui-review-GrRM4J/` に作成し `http://127.0.0.1:4174/` で起動、新しいChromeタブで開いた。元資料の保存済み4スライドと競合状態を保持。元は `/tmp/pptx-user-preview/`、4173。確認画像 `/tmp/pptx-current-header.png`。競合通知は消していない。
- 発表者ビューに無音の動画ミラーと再生・一時停止・位置操作を追加。観客側の状態を250msごとの小さいメッセージで同期、スライド別キーで古い操作を拒否。全画面動画は発表者の現在スライド枠内に表示する。shadow hostのlight DOMには表示されない不具合を実ブラウザーで再現し、shadow tree内へoverlayを置いて修正。
- 最新の発表者通常/全画面2件成功（`/tmp/pptx-presenter-final.log`）。無音、双方向操作、シーク、遷移と終了時の破棄、全画面枠寸法を検証。両モードの画像を目視確認。既存ショー/動画と初期発表者テスト9件、音声/配信2件、dev単体19件も成功。format/lint/dev型検査/build成功。Mac PowerPoint実機との完全一致は未確認。
- 残件: メディアのクリックシーケンス、numSld、ナレーション、オンライン再生、ネイティブ同等の再生UIなど。従来のPreview fidelity差分は未解決。全体完了ではない。

## 2026-10-02: 動画の全画面再生

- 保存済み fullScreen=true の video を、ショー中の再生開始時に viewport 全体へ表示。縦横比を維持し、同一 video 要素で一時停止・再開する。終了後は元の位置のポスターと再生ボタンへ戻り、hideWhenStopped はポスターも非表示。開始時・終了時・破棄時のフォーカスを管理する。
- Microsoft の Mac 向け説明は、再生開始時に画面全体へ拡大することを明記: https://support.microsoft.com/en-us/powerpoint/set-a-video-to-play-full-screen 。終了・Escape・複数全画面動画の競合時の詳細は実機未確認。現状の Escape はショー全体を終了する。
- 旧 dist で開始操作が存在しない失敗を再現（/tmp/pptx-fullscreen-old.log）。既存メディア/ショー7件、dev unit19件、format/lint/dev型検査/build成功。新規全画面動画2件は手動再生・pause/resume・同一要素再利用・終了とEscape破棄・autoplay/hideWhenStopped・フォーカス復帰を検証。全画面と復帰後のスクリーンショットを目視し、ポスター上の不要な標準ボタン表示を除去した。
- PowerPoint は今回も画面サイズ0×0で取得失敗。ネイティブ文書・設定は変更していない。残件は発表者動画同期、クリックシーケンスのメディア開始、numSld、ナレーション、オンラインメディア、再生UIの一致など。
- 45490f9c の CI: Static/Node22/Node26/OOXML成功、Node24は確認時進行中。Preview fidelity は従来と同じ 05-preset-shapes slide1 の 0.6161 vs baseline0.7731 で失敗（/tmp/pptx-fullscreen-existing-fidelity.log）。全体完了ではない。

## 2026-10-02: 埋め込みメディアの実再生

- dev の build manifest に audio/video/online のメディア情報を追加。HTTP state は埋め込みバイナリを含まず、SHA-256 URL に置換し、同じデータを共有。専用 endpoint は MIME / immutable cache / Range 206・416 を返す。
- audience のショー中のみ SVG のポスター画像を HTML audio/video に置換。元画像の座標と親グループ変換を維持。autoplay / loop / volume / mute、終了時の hideWhenStopped、スライド移動・ショー終了時の停止と復元に対応。更新内容が同じなら再生要素を維持する。
- 実 WAV の時間進行、保存属性、移動時停止、再訪時先頭復帰、更新時維持、メディアへフォーカス中の Escape を検証。ブラウザー生成 VP8 WebM を Space キーで再生し、実デコード寸法・終了時非表示を検証。配信の重複排除・部分/末尾 Range・416 も検証。最終3件成功（/tmp/pptx-media-final.log）。既存ショーと合わせた7件も成功（/tmp/pptx-media-browser.log）。dev unit19件、format/lint/dev型検査/build成功。
- 未完了: fullScreen 属性、発表者ビューの動画同期・遠隔再生操作、PowerPoint と同じ再生コントロール、クリックシーケンス内でのメディア開始、スライドをまたぐ再生（numSld）、ナレーション識別、オンラインメディア、ブラウザー非対応 codec。今回をメディア完全一致とは扱わない。現状はブラウザー標準コントロールで再生し、失敗時に再試行を表示する。
- PowerPoint の読み取り接続を再試行したが、画面サイズ0×0で失敗。文書・設定の変更なし。pause/waiting を showWhenStopped の終了と同一視する変更は、実機の裏付けなく行っていない。
- PR #287 は OPEN、head feat/pptx-editor。bcfea96d の Static/Node22/Node26/OOXML は成功、Node24/Preview fidelity は確認時進行中。全UI・全操作一致は引き続き未完了。

## 2026-10-02: ショー設定の抑止チェックと kiosk restart

- 日本語の「ナレーション/アニメーションを表示しない」が正の OOXML 属性へ直結していた逆転を再現・修正。英語も Show without ... に統一。チェックは属性の否定として保存し、キャンセル・Undo・保存後の再読込を検証した。
- kiosk の正の restart ミリ秒を再生開始から計測し、現在のショー先頭へ戻す。自動ループで再計測し、終了時にタイマーを解除。ブラウザーのタイマー上限を超える期間も分割して待機する。
- restart=0 の実機での意味、リンクで入れ子にした custom show の期限継承は未確認。現状は 0 をスケジュールせず、リンク中も全体の期限を継承し、その時点のショー先頭へ戻る。PowerPoint 完全一致と扱わない。
- 発表者ビューでボタンにフォーカスがあると Esc が無視される不具合を新テストで再現して修正。初期の kiosk テストタイムアウトは再起動失敗の証拠ではなく、この終了操作によるものだった。
- 検証: core 3106 passed / 109 skipped、site 117 passed、Svelte 0 errors/warnings。依存順 build、format/lint/root/dev 型検査成功。設定・custom show・ヘッダー6件成功、最終の再生/発表者/kiosk 9件成功（/tmp/pptx-kiosk-final.log）。
- 実機取得は引き続き画面0×0で失敗。今回ネイティブ操作・文書変更はなく復元待ちはない。
- メディア監査: core は埋め込み audio/video と再生属性を扱うが、preview はポスターのみ。dev の manifest/bytes 配信、video/audio プレイヤー、発表者ビュー同期、ナレーション識別・再生、Playback リボンは未実装。showNarration の保存対応を音声再生完了と扱わない。
- 描画 CI の長い Latin 単語の折り返し差分は LibreOffice 基準。latinLnBrk=false を無視する変更は PowerPoint の裏付けがないため採用していない。fc23c22e の Static/Node22/Node26/OOXML は成功、Preview fidelity は 05-preset-shapes slide1 fg-SSIM 0.6161 vs baseline0.7731 で失敗（/tmp/pptx-fidelity-fc23.log）、Node24 は確認時進行中。
- ユーザーの元プレビューは再読み込みしていない。ヘッダー最適化は既に PR #287 の654fa448に含まれる。全UI・全操作一致は未達成。

## 2026-10-01: ショー内の相対リンクと全画面終了競合

- next/prev/first/last のリンクを SVG 生成時の固定スライド番号から実行時アクションへ変更。再生中の custom show/range の順序・重複位置・非表示を反映する。明示的なスライドリンクは従来通り。通常プレビューでも非表示を飛ばす。
- 全画面終了直後に発表者ビューを開始すると、遅延した fullscreenchange が新しい再生を止める不具合を再現し修正。新テストでは停止状態と位置を診断して原因を特定した。
- 旧distでは相対リンクの意味が失われていることを新規テストで確認。最終ブラウザー12/12成功（/tmp/pptx-nav-final-browser.log、session77640正常終了）。通常/発表者、C→A→C→Bの重複順、非表示、終了、既存custom show/履歴を検証。
- 全体テスト3106 passed / 109 skipped、format/lint/root types/dev types成功、依存順buildと最終dev rebuild成功。ZIP保存時刻差で失敗する既存atomicityテストは展開エントリ比較に修正し5/5再検証。レンダラーの旧固定リンクテストは実行時リンクと保存再読込の検証に統合したため総数は1減。
- 実機PowerPointは今回も画面0×0で取得不能。操作・復元待ちはなし。全操作/UI完全一致は未達成。次はCIの描画差分（10546dc3 run36877616297: 05-preset-shapes fg-SSIM 0.6161 vs baseline0.7731）とkiosk restart/ナレーション等を追う。CI比較はLibreOffice基準なので、PowerPointの裏付けなしに閾値緩和しない。
- 元のユーザープレビューは再読み込みしていない。引き続きPR #287へ一本化。

## 2026-10-01: カスタムショーへのリンク

- customShow のクリックアクション（ID と returnToShow）を core、図形/文字/セルのリンク編集、SVG、通常再生と発表者ビューへ接続。OOXML 保存・再読込に対応。
- リンク先ショーの順序・重複を保持し、入れ子から呼び出し元の位置とアニメーション進捗に戻る。タイマーによる終了、再開時の状態リセットも実装。
- core 3107 passed / 109 skipped、format/lint/typecheck、Svelte 0 errors/warnings、依存順 build 成功。関連ブラウザー20件のうち19件が初回成功。残る1件は再読込後の消えたUndo履歴を使うテスト手順を修正し、単独再実行成功（/tmp/pptx-custom-ui-retest.log）。
- Mac PowerPoint は今回も画面0×0で取得失敗。変更・復元待ちはなし。Esc/終了リンクで returnToShow に従う挙動は Microsoft ShowAndReturn の仕様に基づくが実機未検証。次/前/先頭/最後のプリセットリンクは現在グローバルなスライド番号へ解決され、カスタムショー内の順序への対応が残る。
- 全操作の完全一致は未達成。kiosk restart、ナレーション等の残件も継続。ユーザーの元プレビューは再読み込みしていない。
- 直前10546dc3のCIはStatic、Node22/26、OOXML成功。Node24は確認時実行中、Preview fidelityは失敗（今回のログ原因は未確認）。新しいpushのCIも要確認。

## 2026-10-01: 最後に表示したスライドと終了リンク

- `lastSlideViewed` / `endShow` を OOXML のクリックアクション、図形・選択文字・セルのリンク編集、SVG、通常再生・発表者ビューへ接続。文字リンクの重複パーサーを共通化した。
- 訪問履歴はスライド番号だけでなく custom show 内の位置も保持。同一スライドの連続表示を区別し、終了・再開で履歴をリセットする。
- core 3105 passed / 109 skipped、format/lint/typecheck、Svelte 0 errors/warnings、依存順 build 成功。再生・発表者・kiosk のブラウザー9件成功。リンク編集6件も成功（`/tmp/pptx-actions-links-final.log`、session89542正常終了）。初回の3失敗はプリセット追加後のテスト期待値 lastSlide を endShow に直して解消した。
- 旧ビルドで新アクションのリンクが描画されないことを回帰テストで確認。実機 PowerPoint での動作比較は未実施。custom show のリンク起動/return、kiosk restart、ナレーション等は未完了。
- ヘッダー最適化は直前の654fa448で PR #287 に反映済み。元のユーザープレビューは再読み込みしていない。全 UI の完全一致は未達成。

## 2026-10-01: ヘッダー領域の最適化

- 既存の40px Studioヘッダー、保存状態の同一行表示に加え、リボン右端に折りたたみ/展開ボタンを追加。タブ選択・タブの矢印キー操作で再展開。狭い画面ではタブを横スクロール可能にした。
- 検証専用プレビュー1500×900で上部領域は展開時約180px、折りたたみ時約102px。画像 `/tmp/pptx-header-expanded.png` と `/tmp/pptx-header-collapsed.png` を目視確認。元の編集中プレビューは変更・再読み込みしていない。
- format/lint、Svelte型検査（0 errors/warnings）、site既存117テスト、dev build成功。compact-editorブラウザー検証は日英・900/1500pxで成功（session35135終了）。途中の失敗はテストセレクターがフォント欄にも一致/ラッパー階層を誤認したためで、修正済み。
- PowerPoint全操作の完全一致は引き続き未完了。以下の再生機能等の残件を継続する。

## 2026-10-01: browse スクロールバーと kiosk 操作

- browse/showScrollbar を再生画面右端の縦スクロールバーに接続。ドラッグ、トラッククリック、Arrow/Page/Home/End、現在位置同期に対応。スライドの重複順序を保持し hidden を除外。バーの幅は表示領域から確保する。UI はブラウザーで撮影・確認したが Mac 実機との一致は未検証。
- kiosk では loop=false の入力でも自動再生をループさせ、空白クリックと通常の移動キーを無効化。明示的なスライドリンクと Escape は維持。旧ビルドで最終スライド停止を再現した。restart の再起動間隔、ナレーション、custom show のリンク起動/return は引き続き確認・実装が必要。
- 根拠: https://support.microsoft.com/en-us/powerpoint/training/create-a-self-running-presentation と https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.kioskslidemode 。restart について SDK の規範文だけで現在の Office の動作を断定しない。
- PowerPoint 実機は今回も画面サイズ0×0で取得失敗。ネイティブ操作や変更はしていない。ユーザーの元プレビューは再読み込みしていない。
- 最終 format/lint/dev typecheck/build 成功。最終統合ブラウザー16/16成功（/tmp/pptx-browse-kiosk-final-browser.log、session93916正常終了）。途中14/16の失敗はテスト側の fullscreen 終了待ちと編集中変数参照を修正。
- 056e82a7 の CI run36873233002 は確認時 Static 成功、他チェック実行中。全体の完全一致はまだ未達成。

## 2026-10-01: 再生対象の削除とウィンドウ表示

- 再生対象の Custom Show を削除したとき、同じ Undo トランザクションで再生対象を all に戻すよう修正。他の設定は保持。旧ビルドで参照が残る回帰を再現し、保存した PPTX・Undo でのショー/選択 ID 復元・無関係なショー削除での選択維持を検証。
- browse モードの Present が fullscreen API を呼ぶ不具合も旧ビルドで再現し修正。スクロールバーの表示設定はまだ再生 UI へ未反映。Microsoft の ShowScrollbar 資料ではスライド間移動用とされるため、stage の overflow による近似は採用していない。次はスライド移動バーの実機動作を確認して実装する。
- 根拠: https://learn.microsoft.com/en-us/previous-versions/office/office-12/ff763072(v=office.12) と https://support.microsoft.com/en-us/powerpoint/training/create-a-self-running-presentation 。ウィンドウ表示の対応だけで Mac 実機との完全一致とは扱わない。
- format/lint、Svelte 0 errors / 0 warnings、site 117 tests、dev build/typecheck 成功。core の変更なし。統合ブラウザー16/16成功（Custom Shows、presentation、browse3件、設定dialog、再生6件）。ログ `/tmp/pptx-show-followup-browser.log`、session53924は正常終了。
- PR #287 の 9b734b5a CI run36872039905 は確認時実行中。前の42f5b21dは後続pushでcancelled。全体の完全一致は未完成。ユーザーの元プレビューを再読み込みしない。

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
