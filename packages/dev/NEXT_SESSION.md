## 2026-10-02: 連続改行の空行を実機と比較

- Mac PowerPoint で Before / Shift+Enter / Shift+Enter / After を入力し、空行を作る後方の改行だけを 80pt に変更。Before と After の間隔が広がることを確認（中央配置のため Before は上、After は下へ移動）。先頭空行と同様に空行の改行書式が高さへ影響する。一時変更を Undo disabled まで取り消し、Outline title / Ordinary text box を復元して保存済み。
- 連続改行と、文字のない a:r が改行の直前にあるケースをブラウザー回帰へ追加。図形・表とも編集開始の位置差は 2px 未満。既存の先頭・途中改行と入力書式のテストを含め 8 件成功。
- SVG / foreignObject の先頭空行サイズ比較を 2 個の連続改行にも拡張。今回はテストと比較記録のみで、本体実装の追加変更なし。
- 残件: 末尾改行、空段落の箇条書き、表スタイル等。完全一致は未完了。参照ファイルは復元済み、4173 と .pnpm-store を保持し、PR #287 に集約。

## 2026-10-02: 先頭改行の高さと編集開始位置

- Mac PowerPoint の先頭 Shift+Enter のみを 40pt → 80pt に変更すると、後続 After が下へ移動することを画面で確認。文字の直後の改行とは挙動が異なる。接続復旧後に Font Size / Typing を Undo disabled まで取り消し、Outline title / Ordinary text box に戻して保存済み。
- 先頭改行のサイズを通常表示が無視することで、図形・表とも編集開始時に約 63px ずれるケースをブラウザーで再現。描画でも a:br の実効書式を解決し、空行の高さに反映。文字のある行は改行のサイズで広げない。SVG の空行にも改行のフォントメトリクスを使用する。
- 大きな改行を X に置き換えた際、80pt と通常の行高を保持し、BeforeXAfter になることを図形・表で検証。
- 検証: format/lint/core typecheck/build、preview/DSL/dev typecheck、依存パッケージを含む dev build 成功。core 3300 成功 / 109 skip、関連 browser 12 成功。先頭改行の修正前の約 63px のずれは、修正後に 2px 未満。
- 全操作一致は未完了。連続改行・末尾改行、空文字 run、空段落の箇条書き、表スタイル等の残件を継続。4173 と .pnpm-store を保持。PR #287 に集約。

## 2026-10-02: 改行だけが大きい場合の編集時レイアウトシフト

- Mac PowerPoint の reference.pptx で Before / Shift+Enter / After を入力し、改行のみ選択して 40pt → 80pt に変更。前後の文字・行の画面位置は不変。保存 XML で a:br/rPr sz=8900 と normAutofit fontScale=90000 を確認（UI は 80pt）。リボンのサイズ欄は AX click ではフォーカスが移らず、Font メニューを開いた直後の Tab でサイズ欄へ移動できた。
- 文字の直後の改行 span に line-height:0 と vertical-align:top を指定し、書式を保持したまま行ボックスへの影響をなくした。段落先頭や連続改行の空行には適用していない。このケースの実機比較は図形で実施、表は共通の編集表示経路の回帰を検証。
- ブラウザーで修正前に図形・表の編集開始時に約 53px の位置ずれを再現。修正後は 2px 未満、改行のコピーで 80pt を保持。継承書式・空段落を含む関連 10 件成功。
- 検証: format/lint/core typecheck/build、core 3298 成功 / 109 skip、site 135 成功、Svelte 0 errors/warnings、preview build、DSL/dev typecheck、dev build 成功。
- 一時変更を Undo disabled まで取り消し、Outline title / Ordinary text box に戻して保存済み。4173 と .pnpm-store は保持。全操作一致は未完了、PR #287 に集約。

## 2026-10-02: 明示改行の継承書式

- getShapeRunFormatEffective / getTableCellRunFormatEffective に { breakIndex } を追加。a:br の直接 rPr と段落・リスト等の既定書式を既存 resolver で合成。通常 run / field のインデックスは別々のまま維持。
- 選択、編集 HTML、書式付きコピーで改行の継承フォント・サイズ・太字が失われる不具合を修正。図形の選択書式で変更前の失敗を再現し、図形・表、範囲外、保存再読込、コピーを検証。ブラウザーで両方の改行 span が Courier New / 28pt / bold になることも確認。
- 根拠: Microsoft SDK Drawing.Break remarks (ISO/IEC 29500 §21.1.2.2.1): https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.break?view=openxml-3.0.1 。改行の rPr は挿入文字の書式を保持する。行高にどう影響するかはこの資料だけでは確定しないため、描画側を推測で変更していない。
- 検証: core 3298 成功 / 109 skip、site 135 件、Svelte check 0 errors/warnings、関連 browser 6 件成功。format/lint/core typecheck/build、preview build、DSL typecheck、dev typecheck/build 成功。
- 改行の描画側は依然 fmt:null / 既定サイズ。行高・前後の文字サイズ・改行直後の入力は Mac 実機比較が必要。空文字 run、空段落の箇条書き、表の共有辺/スタイル等も残る。全操作一致は未完了。参照文書に変更なし。4173 と .pnpm-store を保持。PR #287 に集約。

## 2026-10-02: 図形の空段落と編集開始時の位置

- getShapeRunFormatEffective の null 指定で段落末尾の実効書式を取得。図形の空段落も表示・編集 HTML・カーソルの書式取得に同じ継承を使用する。
- endParaRPr の 18pt / 36pt が同じ高さになる不具合を SVG / foreignObject で変更前に再現し、保存再読込も検証。ブラウザーで図形の空行の行間が normal になり約 5px ずれる追加原因を確認し、空段落の既定行間を既存 LINE_HEIGHT に合わせた。明示した段落行間は優先する。
- 関連 browser 6 件成功。図形・表の空行後の文字位置差 2px 未満、36pt・Courier New・太字・色、Font ダイアログを検証。core 3296 成功 / 109 skip、site 135 件、Svelte check 0 errors/warnings、format/lint/typecheck/build と preview・DSL・dev の型検査/build 成功。最終行間修正後は関連 core 5 件・browser 6 件、format/lint/preview 型検査、preview/dev build を再実行。
- PowerPoint の画面取得は成功。reference.pptx は Outline title / Ordinary text box、保存済み、Undo disabled。一時変更なし。この空段落ケースの Mac 実機比較自体は未実施。
- 残件: 空文字 run・明示改行の書式、空段落の箇条書き、フィールド独自 pPr、表の共有辺/スタイル等。全操作一致は未完了。4173 と .pnpm-store を保持し、PR #287 のみに集約。

## 2026-10-02: 図形フィールドの実効文字書式

- 図形の a:fld でも既存 getShapeRunFormatEffective に { fieldIndex } を指定できるようにし、通常 run と同じ段落・リスト・プレースホルダー・テーマ継承を適用。数値指定は従来どおり a:r のみを数える。
- 図形の日付フィールドが段落の Courier New / 28pt / bold / 色を失うことを SVG と foreignObject の通常文字との比較で再現し修正。直接指定した italic との合成、範囲選択の書式、数値 run 指定の範囲エラーを検証。
- 編集 HTML、選択書式、コピー経路も同じ resolver を使用。ブラウザーで図形フィールドの編集開始、書式付きコピー、Font ダイアログまで検証。
- 検証: core 3294 成功 / 109 skip、format/lint/typecheck/build、preview・DSL・dev の型検査・build 成功。Svelte check 0 errors/warnings、site 135 件、関連 browser 5 件成功。
- ネイティブ文書に変更なし。このフィールドケースの Mac 実機比較は未実施。フィールド独自 pPr とリンク、空文字 run / 改行のみの段落、空段落の箇条書き、表の共有辺/スタイル等は未確認。全操作一致は未完了。4173 と .pnpm-store を保持。PR #287 のみに集約。

## 2026-10-02: 段落末尾書式を既存文字から分離

- endParaRPr を既存の最終 run / field に継承していた共有 resolver を修正。null（段落末尾の入力位置）のみ適用する。根拠は ECMA-376 §21.1.2.2.3 と Microsoft SDK EndParagraphRunProperties remarks（新しい文字挿入用）。
- 図形と通常文字・フィールド混在の表で、変更前に 48pt・太字が既存文字へ漏れることを再現。保存再読込と入力位置の書式保持を検証。
- 既存の選択書式テストが endParaRPr を本文の太字として用いていたため、本文 rPr で太字を指定する fixture に訂正。前回の copyTextRange selector 拡張に追従していなかったテスト callback の型エラーも修正。
- 検証: core 3292 成功 / 109 skip、format/lint/typecheck/build 成功。preview・DSL・dev の型検査・build 成功。関連 browser 4 件成功。
- 今回ネイティブ参照文書に変更なし。このケース自体の Mac 実機比較は未実施。フィールド独自 pPr、図形フィールドの実効書式、空文字 run / 改行のみの段落、全操作一致は未完了。4173 と .pnpm-store を保持。PR #287 のみに集約。

## 2026-10-02: 表内フィールドの実効文字書式

- 表内の a:fld が通常文字と違って段落既定のフォント・サイズ・太字を継承しない不具合を SVG / foreignObject の比較テストで再現し修正。
- 既存 getTableCellRunFormatEffective に { fieldIndex } 指定を追加。数値は引き続き a:r のみを数える。直接書式優先、保存・再読込、範囲外の検証を追加。
- 編集 HTML、選択範囲の書式取得、コピー処理も同じ実効書式を使用。ブラウザーで表示修正後もコピーにフォントが落ちるケースを検出し、その経路も修正した。
- 検証: core 3290 成功 / 109 skip、format/lint/typecheck/build、preview・DSL・dev の型検査・build 成功。site 135 件成功、Svelte check 0 errors/warnings。browser 4 件成功（通常文字・フィールド・表スタイル・空段落）。
- このターンはネイティブ参照文書に変更なし。直前の再接続で画面取得成功、Outline title / Ordinary text box と保存済み・Undo disabled を確認。
- 未確認: フィールド独自 pPr、通常 run と field と改行の混在時の endParaRPr 継承、図形内フィールドの実効書式、空文字 run / 改行のみの段落。今回のフィールドケースの Mac 実機比較は未実施。全操作一致は未完了。
- ユーザーの 4173 と .pnpm-store は保持。PR #287 のみに集約。

## 2026-10-02: 表の空段落の文字サイズと編集時の位置

- `getTableCellRunFormatEffective` の runIndex に null を許容し、段落末尾の実効書式を既存の段落・リスト・表スタイル継承で解決。通常の数値インデックスの範囲検証は維持。
- 空の `<a:p>` の endParaRPr を SVG の行高、foreignObject の文字サイズ、編集 HTML、空段落のカーソル位置でのフォント取得に反映。18pt と 36pt の空行が同じ高さになる不具合を変更前に再現。
- ブラウザーで編集開始時に約 6px ずれる追加原因を確認。表の foreignObject コンテナーだけ残っていた line-height:1.2 を既存の文字描画・編集と同じ LINE_HEIGHT に統一。新規ブラウザー回帰で空行の後の文字位置差が 2px 未満、36pt・Courier New・太字・色の保持、Font ダイアログのサイズ・フォントを検証。
- 検証: core 3287 成功 / 109 skip、format/lint/typecheck/build、preview・DSL・dev の型検査と build 成功。最終変更後の関連 core 26 件と browser 4 件成功。site 単体 135 件成功、Svelte check 0 errors/warnings。
- 今回はネイティブ文書に変更なし。この空段落ケース自体の Mac 実機比較は未実施。フィールドの実効書式、空文字 run や改行のみの段落、空段落の箇条書き表示は引き続き確認が必要。表以外も含む全操作一致は未完了。
- 4173 と .pnpm-store は保持。PR #287 のみに集約。

## 2026-10-02: 横結合セルの最終列を Mac 実機に合わせる

- reference.pptx の 3×3 表で第 1 行を Select Row > Merge Cells により横結合し、Header Row を外して Last Column の on/off を比較。on では結合セル全体が青、off では帯の淡色になることを確認。
- lastCol および右隅の列判定を gridSpan の終端へ変更。変更前の失敗を再現し、右端に届かない結合セル、on/off、保存再読込を回帰テストで検証。
- 検証: core 3284 成功 / 109 skip、format/lint/typecheck/build 成功。preview・DSL・dev の型検査と build も成功。
- 一時変更 5 操作を Undo して Undo disabled / Redo Table を確認し保存済み。接続再確認でも画面を取得でき、Outline title / Ordinary text box と Saved to my Mac を確認。未復元変更なし。
- 残件: 結合セルの帯・共有辺・隅の優先順位、他の組み込み GUID、非単色塗り、tblBg。全操作一致は未完了。4173 と .pnpm-store は保持。

## 2026-10-02: 縦結合セルの集計行を Mac 実機に合わせる

- Mac PowerPoint の reference.pptx で 3×3 表の第 1 列を Table Layout > Table > Select Column > Merge Cells により縦結合。Header Row を外し、Total Row の on/off を比較。on では結合セル全体が accent1 の青、off では帯の淡色になることを画面で確認した。結合セルの開始行ではなく終端が最終行に達することが集計行の条件。
- `tableStylePartsForCell` の lastRow および下隅の行判定を rowSpan の終端へ変更。lastCol は未確認のため変更していない。回帰テストは変更前に淡色の誤表示で失敗、変更後は on/off と保存再読込に成功。
- 検証: core 3283 成功 / 109 skip、format/lint/typecheck/build 成功。preview・DSL・dev の型検査と build も成功。
- 一時変更は表挿入・結合・Header Row・Total Row の on/off の計 5 操作を Undo し、Undo disabled / Redo Table を確認して保存済み。未復元変更なし。
- 残件: 横結合の lastCol、結合セルの帯・共有辺、他の組み込み GUID、非単色塗り、tblBg。全操作一致は未完了。ユーザーの 4173 と .pnpm-store は保持。

## 2026-10-02: No Style, Table Grid の組み込み定義

- GUID `{5940675A-B579-460E-94D1-54222C63F5DA}` のみの表で、塗りなし・tx1 色の 1pt 外周/内部罫線を解決。定義は Microsoft SDK TableStyle remarks の ISO/IEC 29500 例が根拠。空の条件領域は省略している。
- 6 個のスタイルオプション全 64 組み合わせ、全セルの四辺、対角線なし、保存再読込を回帰テストで確認。追加前は fill=inherit の失敗を再現した。
- 検証: core 3282 成功 / 109 skip、format/lint/typecheck/build 成功。preview・DSL・dev の型検査・build、Svelte check 0 errors/warnings。
- この GUID の Mac 実機比較は未実施。他の組み込み GUID、結合・共有辺、非単色塗り、tblBg は未完了。全操作一致は未完了。
- ネイティブ文書には今回変更を加えていない。ユーザーの 4173 と .pnpm-store は保持。

## 2026-10-02: 帯スタイルの罫線

- band1H/band2H/band1V/band2V の明示した辺を、表内部でもそのセルの辺として解決。辺が未指定の場合は insideH/insideV を維持する。明示辺と内部罫線の共存を回帰テストに追加。
- 根拠: LibreOffice `oox/source/drawingml/table/tablecell.cxx` の applyTableStylePart と pushToXCell（https://github.com/LibreOffice/core/blob/master/oox/source/drawingml/table/tablecell.cxx）。これは他実装との照合であり、PowerPoint 実機の帯罫線一致は未検証。完全一致の証拠とはしない。
- PowerPoint 画面取得成功。reference.pptx は Outline title / Ordinary text box、保存済み、Undo 無効。結合操作の試行で生じた表挿入・Total Row・サイズ変更は全て取り消し済み。結合セルの条件領域がアンカー基準か端基準かは未確定。
- 検証: core 3281 成功 / 109 skip、format/lint/typecheck/build 成功。preview・DSL・dev の型検査と build、site Svelte check 成功。
- 未完了項目は下記の表外観記録を参照。全操作一致は未完了。

## 2026-10-02: 埋め込み表スタイルの外観

- `getTableCellAppearanceEffective` で tcStyle の単色塗り、明示 noFill、外周・内部罫線を解決し、プレビューに統合。推測アクセント塗り、強制白背景、未指定の灰色グリッドを除去。
- 実機の `tableStyles.xml` を `test/fixtures/native-table-style.xml` として保存。dml-main.xsd による xmllint 検証成功。見出し・帯の塗り、見出し下罫線、XML の保存保持を回帰テストに追加。
- fillRef/lnRef のテーマ行列参照、背景塗り参照、phClr と色変換を解決。既定 Medium Style 2 – Accent 1 は GUID のみでも実機定義から解決する。他の組み込み GUID は未対応。表の寸法はキャッシュし、行列追加・削除時に無効化する回帰を追加。
- 未完了: gradient/pattern/image の実描画、罫線競合の完全な優先順位、結合セルをまたぐ部分的な共有辺、結合領域の lastRow/lastCol スタイル選択、帯領域の辺解決、表背景 tblBg。
- 最終 core 検証: 3280 成功 / 109 skip、format/lint/typecheck/build 成功。preview・DSL・dev の型検査と build、Svelte 0 errors/warnings。最終ビルドの表文字編集 browser 2 件成功。site/dev の単体 154 件成功。
- PowerPoint の reference.pptx は画面取得・保存済み・Undo 無効を再確認。未復元変更なし。ユーザーの 4173 と .pnpm-store は保持。全 PowerPoint 操作一致は未完了。

## 2026-10-02: 埋め込み表スタイルの文字書式

- 既存 `getTableCellRunFormatEffective` に `tableStyles.xml` の `tcTxStyle` 継承を統合。wholeTbl・帯・行列端・四隅の順序、直接フォント・テーマフォント・色・太字・斜体を扱う。セル側の明示書式を優先する。
- 通常表示（SVG / foreignObject）と編集・内部コピーの回帰テストを追加。行・列の書式重複、明示 off、テーマフォントも確認。core3269成功/109skip、site135成功、dev19成功、関連browser4成功。format/lint/root・preview・DSL・dev型検査、core/preview/DSL/dev build、Svelte 0 errors/warnings。
- PowerPoint 接続が復旧し、reference.pptx の画面と保存状態を再確認。一時挿入した表を保存した検証用コピーは `/tmp/pptx-outline-audit/native-table-style.pptx`。Mac が既定 GUID のスタイル定義を tableStyles.xml に保存することを確認した。参照元は表を Undo して保存済み、文字は Outline title / Ordinary text box、未復元変更なし。
- 残件: 定義のない組み込み GUID、tcStyle の塗り・罫線、空段落 / field の文字書式。今回の文字継承は完全な表スタイル対応を意味しない。全 PowerPoint 操作一致も未完了。
- 4173 のユーザー編集と .pnpm-store は保持。PR は引き続き #287 のみ。

## 2026-10-02: 表の段落既定文字書式

- `getTableCellRunFormatEffective` を追加。セル内のrun、段落既定、アウトラインレベル既定とテーマを共通の読取り経路で解決する。通常描画・inline編集・Fontダイアログ・コピーで同じ実効書式を使う。
- 修正前は段落で指定したCourier Newが編集開始時にCalibriへ変わることをブラウザーで再現。修正後はフォント、28pt、太字、色を保持し、内部clipboardとFontダイアログにも反映されることを確認。
- 検証: core3266成功/109skip、site135成功、dev19成功、最終ビルドの関連browser5件成功。format/lint/root・preview・DSL・dev型検査、core/preview/DSL/dev build、Svelte 0 errors/warnings。
- 残件: `tableStyles.xml` の `tcTxStyle` と条件領域の継承、組み込みGUIDのスタイル定義、空段落/fieldの実効書式。新APIの現時点の対応範囲はセル内の文字書式とテーマであり、完全な表スタイル対応ではない。
- 次の実装は既存APIを拡張する。TableCellDataには既にCELL_ROW/CELL_COL/CELL_TABLEがあるため、行列引数の追加は不要。スタイルpartをリレーションから取得し、wholeTbl/帯/端/角の優先順位を仕様と実機で確認する。GUIDのみで定義がない組み込みスタイルも残件として扱う。
- PowerPointは起動検出されるがcgWindowNotFoundで画面取得不可。今回ネイティブ文書の変更は行っていない。以前の記録ではreference.pptxはUndo disabled・復元済みだが、現在画面での再確認はできていない。
- 全操作一致は未完了。4173のユーザー編集と.pnpm-storeは引き続き保持する。

## 2026-10-02: 下線色の独立編集

- Mac Font ダイアログで、下線なし時の色選択無効化と、下線あり時の「Automatic」を実機確認。ダイアログを Cancel し、reference.pptx の Undo disabled を確認。参照文書の未復元変更なし。
- `TextFormat.underlineColor` を追加。明示色は `a:uFill/a:solidFill`、null は `a:uFillTx`。継承停止、テーマ色、PPTX保存再読込、未知の非solid下線塗りの保持、rPr子要素順序をテスト。
- Font ダイアログに下線色と Automatic を追加。選択範囲、書式、選択外、Undo、保存XMLをブラウザーで検証。内部/HTML clipboardと表セルのテーマ色解決も対応。
- SVG rasterizer が text-decoration-color を無視するため、明示色の単線下線は測定済み幅の線で描画。native 下線との重複なし。画像上の文字/strikeと下線の色分離もテスト。
- 検証: core3259成功/109skip、site135成功、関連ブラウザーテスト成功。最終SVG調整後のpreview/layout101件も成功。format/lint/root・DSL型検査、core・preview・editor build、Svelte 0 errors/warnings。
- 残件: 下線色の濃淡テーマパレット、gradient/pattern塗りの編集・描画、Mac全フォントでの下線寸法一致。normalizeHeight実描画、表セルの完全な継承書式、アウトラインの特殊下線、全操作一致も未完了。今回のテスト成功はこれらの完了を意味しない。

## 2026-10-02: 全下線スタイルの編集表示

- 通常表示とinline編集で共通の `textUnderlineStyle` を使用。太線、長破線、鎖線、二重波線、単語のみの下線を編集開始後も維持する。外部HTML clipboardは従来のCSS書式を維持。
- 本番エディターで全18設定の編集開始前後を検証。文字矩形の差は各軸2px未満、下線の相対寸法とパターン一致、strikeは実線。修正前にheavyがauto厚へ変化する失敗を再現済み。空白のみの通常下線も維持。
- 表セルの明示的なscheme色はスライドのカラーマップとテーマで解決して編集HTMLへ渡す。表スタイル継承と色変換の完全な解決は引き続き残件。
- アウトラインはUI前景色を継承するため、SVG背景の固定色を導入していない。特殊下線の完全な表示は残件。Mac全フォントでの線幅・波形の一致、normalizeHeight描画、全操作一致も未完了。
- 検証: core3250成功/109skip、site134成功、関連browser7件成功。format/lint/root・DSL型検査、core・preview・editor build、Svelte 0 errors/warnings。
- PowerPoint接続と画面取得は復旧。reference.pptxのUndo無効、Outline title表示を確認。今回ネイティブ文書は変更していない。

## 2026-10-02: 編集開始時とHTML貼り付けの下線保持

- `textClipboardHtml`が下線を単線へ丸め、inline編集開始時にdbl/wavy等が変わることを本番エディタのブラウザー回帰で再現して修正。CSSで表現できるdouble/dotted/dashed/wavyを出力・再取り込みする。pattern付き下線とstrikeは外側spanの実線strike、内側uの下線へ分離し、親のstrikeをパーサーで保持。
- dbl/wavyの編集開始でglyphのx/y/width/height差が2px未満、wavyとstrike併用でstrikeがsolidであることを検証。HTML clipboardを含むブラウザー5件成功。
- normalizeHeightを含む内部clipboard metadataが拒否されていた。boolean検証を追加し、true/falseともcopy→parse→pasteで他の書式とともに保持する回帰を追加。
- 品質確認: core3250件/109skip、site133件、format/lint/typecheck/core・editor build、Svelte 0 errors/warnings。
- 残件: 編集HTMLではheavyの太さ、long dash/dash-dot/double-wave/words-onlyなどの完全な表示対応が未完了。previewのSVG背景パターンと編集HTMLのCSS点線/破線は寸法の一致まで未確認。normalizeHeightは保存・clipboard対応であり、文字高さの描画は未実装。全操作一致の完了ではない。
- PR #287へ継続反映。ユーザーの4173の未保存編集と.pnpm-storeには触れない。

## 2026-10-02: 下線描画と表セルのホーム Font 操作

- 下線のDrawingML値をSVGへ渡す際の単線への丸めを撤去。二重線・太線・点線・長破線・鎖線・二重波線を明示的なSVG線/パスで描画。単語のみの下線は計測済みtoken幅で空白を除く。通常単線のフォント固有下線位置は従来のtext-decorationを保持。
- HTML側はCSS下線とSVG背景パターンで描き分ける。Chromiumでは太い破線のtext-decoration-thicknessが表示に反映されないため、点線/破線系も背景パターンを利用。strikeは別spanで単線を維持。
- ブラウザーの18設定で画像が区別でき、文字の位置・幅・高さが不変なことを確認。resvgでも18設定とstrike併用を実画像で検証。関連単体109件成功。線幅・波形寸法のMac完全一致を証明したものではない。
- Mac PowerPointの下線選択メニューを取得できた。メニュー確認のみで文書変更なし。reference.pptxのUndo disabled、未復元変更なし。
- 表セル選択時のホームFont操作を有効化。太字・サイズ・大文字小文字変更をセルへ適用し、ダイアログのセル書式取得を共通化。ホーム太字とダイアログ文字間隔、選択外保持、Undoをブラウザーで確認。表スタイル等の継承書式表示は未完了。
- 品質確認: core3248件/109skip、site133件、format/lint/typecheck/core・preview・dev build/DSL typecheck、Svelte check成功。
- 残件: inline編集中の特殊下線表示、下線色、Equalize character heightの実描画、表セルの完全な継承書式読取り、全操作の実機比較。全操作一致は未完了。

## 2026-10-02: 下線選択・XML boolean・small caps

- Font ダイアログから全18種類のDrawingML下線値を選択可能。wavyHeavyの選択・保存をブラウザー回帰で確認。描画側は依然一部スタイルを単線へ簡略化しており、全下線の見た目一致は未完了。
- OOXMLの `b="false"` / `i="false"` をオンとして読む不具合を修正。1/true/0/falseと、継承元オンに対する明示オフの保存再読込を検証。
- SVGのsmall capsは小文字由来の文字を小さな大文字として描画・計測し、元の大文字と行メトリクスを保持。縮小率0.8は近似であり、Mac実機の全フォント一致は未確認。
- format/lint/core typecheck、core/preview/editor build、DSL typecheck、Svelte 0 errors/warnings、core3236件成功/109skip。結合文字・折返しの追加修正後はsmall caps/text layout/SVG modesの66件成功。FontDialogブラウザー回帰成功。
- PowerPoint画面取得は復帰。reference.pptxはOutline title / 44pt、Undo disabledを実機確認。一時変更の復元待ちはなし。下線メニューを開くと画面取得0×0になる事象があったが、通常画面は取得できている。
- 残件: 下線色、下線各種の実描画、Equalize character heightの実描画、表セルの継承書式表示など。同一PR #287を継続し、4173の未保存編集と.pnpm-storeは触らない。全操作一致は未完了。

## 2026-10-02: Font ダイアログ

- Home の Font ボタンと Cmd/Ctrl+T で Font / Character Spacing の2タブを開く。Latin/Asian フォント、スタイル、サイズ、色、下線、二重取り消し線、上付き/下付きと位置、caps、文字高さ、文字間隔、カーニングを設定可能。
- 変更した項目だけ適用し、未編集の混在書式と Complex Script フォントを保持。Escape 後に再表示できなくなる問題も修正。
- 選択文字・表セル・アウトラインの保存OOXML、選択外保持、Cancel/Escape、Undoをブラウザー検証。ダイアログ内フォント一覧の選択、既存フォント選択、日英ホーム幅を含むブラウザー計5件成功。site133件、format/lint/core typecheck、Svelte 0 errors/warnings、editor build成功。coreは前回3228件成功から変更なし。
- 実機の Font ダイアログ閲覧中に画面取得0×0が再発。今回は文書変更なし。以前の一時変更は取り消し済み。実機の設定画面は開いたままの可能性がある。
- 残件: 下線色、全下線スタイルの選択、small caps のSVG字形サイズ、Equalize character heightの実描画、表セルの継承書式表示など。Macとの全操作一致は未完了。
- 同一PR #287を継続。4173の未保存編集には触らず、確認用4175を使用。.pnpm-storeも変更しない。

## 2026-10-02: Equalize character height の OOXML 対応

- `TextFormat.normalizeHeight` を追加。`normalizeH` の true/false/1/0 を読み、明示オン・オフを保存。実効書式の継承と解除、範囲書式リセットも対応。
- Mac の reference.pptx で Equalize character height をオンにして保存すると `normalizeH="1"` を確認。Cmd+Zと保存後に属性が消え、Undo disabled。未復元の一時変更なし。
- core 全体3228件成功/109skip、format/lint/typecheck/build成功。字形の高さを実際に揃えるプレビュー描画はまだ未実装。保存対応を視覚的な完全一致と混同しないこと。
- Font ダイアログは font_dialog エージェントが作業中。変更した項目だけ適用し、混在する未編集書式を保持すること。Mac の2タブと各チェックボックス・offset・別Asian fontに合わせる。下線色API/プレビュー描画は残件。

## 2026-10-02: カーニング描画と文字サイズショートカット

- OOXMLのkern閾値をSVG/HTML描画、fontkit/ブラウザ計測に反映。inline HTMLもサイズと閾値からfont-kerningを設定。SVGの結合判定・計測キャッシュにも反映する。閾値未満・境界・0・省略を検証。
- 実ブラウザで本番browserTextMeasurerを使いAVの有効→無効→有効の幅を検証。HTML clipboard/inline CSS回帰4件成功。これは実機との全字形・全フォント一致を証明するものではない。
- 全体単体テスト3221件成功/109skip、site133件成功、変更後preview101件成功。format/lint/core typecheck/core build/preview build/DSL typecheck/editor build成功、Svelte 0 errors/warnings。
- FontダイアログUI、SVGの小文字を小さく描くsmall caps、Equalize character heightは残件。カーニング描画は対応したが、設定UIと実機比較の拡大を引き続き行う。全操作一致は未完了。

## 2026-10-02: カーニング無効の実機保存形式

- Mac PowerPoint の reference.pptx でタイトル全文44ptを選択し、Font > Character Spacing > Use kerning for fonts を解除して保存すると、該当 a:rPr に `kern="0"` が出力されることを確認。元はrun属性なしで12pt以上を継承。ゼロは無効として描画する根拠になる。
- 検証後 Cmd+Z と保存を実行し、Undo disabledを確認。変更前バックアップは `/tmp/pptx-outline-audit/reference-before-kerning.pptx`。復元待ちなし。
- 複数図形の⌘⇧. / ⌘⇧,、選択外保持、1回のUndo/Redo、保存OOXMLを確認するブラウザー回帰が成功（font-size-multiselect.test.mjs）。表セルは単独拡大18→20pt、矩形4セルの⌘⇧,による縮小18→16pt、選択外保持、Undo後の全セル書式復元を確認。Outlineを含めた最終ブラウザー回帰5件成功（`/tmp/font-shortcuts-final-browser.log`）。全操作一致は未完了。

## 2026-10-02: コピー・貼り付け時の大文字書式保持

- HTML clipboardでall caps / small caps / 明示的解除を読み書きする。CSSのtext-transformとfont-variant-capsは別々に継承し、子の片方の解除で親のもう片方の指定が消えないようにした。
- Chromium回帰4件成功。入れ子・往復・shape/tableの貼り付けとUndo・all caps編集後の追加入力の保存を確認。site133件成功、Svelte 0 errors/warnings、format/lint/core typecheck成功。
- 確認用4175はHTTP200。4173の未保存編集には触れていない。

## 2026-10-02: Mac文字操作の追加実測

- 実機の選択文字で⌘⇧.は44→48pt、⌘⇧,は44→40pt。各操作をUndoし44ptへ復元済み。
- `hELLO wORLD. nEXT tEST` 全選択でSentence caseは `Hello world. Next test`、Capitalize Each Wordは `Hello World. Next Test`。句読点・Unicode一般の全一致を証明するものではない。
- ⌘TでFontダイアログが開く。Font/Character Spacingの2タブ。Character SpacingにはNormal、By（0–1000pt）、Use kerning for fonts、閾値（1–1000pt）がある。参照ではカーニング有効・12pt以上。閲覧後Cancel済み。
- 参照文書は `Outline title`、44pt、Undo disabledへ復元済み。画面取得も成功。未復元の変更はない。
- Fontダイアログ、カーニング描画・計測は未対応。SVG側small capsも単なる大文字化に留まり、実機同等の字形サイズを未検証。これらを全操作一致の残件として扱う。

## 2026-10-02: Homeの文字種変換

- HomeにChange Case（sentence/lower/upper/title/toggle）を追加。選択文字、caretの現在単語、図形全体、アウトラインから操作可能。setShapeText/setTableCellTextの既存APIへ `{ case: ... }` を追加し、元run/fieldのa:tだけを変更して書式・hyperlink・未知XMLを保持。Unicode展開（ß/İ）、Greek sigmaの文脈、改行位置を検証。
- Homeのcompact切替境界を2000pxへ調整。日英900/1500/1601/1900/2100pxで横スクロールなしのブラウザー回帰成功。
- 検証: core3220件成功/109skip、site133件成功、format/lint/typecheck/core build/editor build成功、Svelte 0 errors/warnings。アウトラインの選択範囲・caret単語・Undo/Redo・保存ブラウザー回帰成功。
- Canvasブラウザー回帰も成功。Unicode選択範囲、周辺文字・italic書式保持、caret現在単語、図形全体、Undo/Redo、保存再読込を確認。今回のブラウザー回帰はHome/Outline/Canvasの計3件成功。
- Mac参照reference.pptxへの一時入力はUndo済み。Outline titleへ復元、Undo disabledを確認。reference-before-font.pptxとバイト一致。画面0×0の失敗が断続したが、最後の再取得では正常なスクリーンショットを確認。復元待ちなし。
- 実機で確認済みのChange Case仕様はメニュー5種類とcaret時の現在単語UPPERCASE。sentence/titleの句読点・Unicode細部まで実機一致を確認したわけではない。Fontダイアログ、文字サイズショートカット、カーニング描画など残件あり。全操作一致は未完了。
- 同一PR #287を継続。確認用4175 HTTP200。4173の未保存編集と.pnpm-storeには触らない。

## 2026-10-02: 文字範囲編集の未知XML保持

- setShapeText/setTableCellTextの範囲置換で、境界runの変更しない部分から未知属性・子要素が消える問題を修正。元runを複製しa:tのみ更新する。部分fieldは従来どおりliteral run化し、field専用id/type/pPrを除去、未知属性・rPr等は保持。範囲書式変更も共通化。
- 回帰: shape/tableの保存再読込、実際の外部hyperlink relationship、mc:Ignorableな未知属性、rPr/extLst、fieldのpPr除去を検証。core3210件成功/109skip、format/lint/typecheck、core/editor build成功。ログ `/tmp/partial-run-{gates,tests,build,editor-build}.log`。
- ブラウザー8件成功（文字・表セルclipboard 2件、図形/表セルのtop/center/bottom余白・位置6件）。旧ヘッダー構造を参照していたtext-body-editingテストを現行の折りたたみパネル操作へ更新。clipboardログ `/tmp/partial-run-clipboard.log`。
- Mac Change CaseをcaretのみでUPPERCASEにすると、現在単語だけが変換された（Outline title → OUTLINE title）。Undo済み・Undo disabled確認、reference.pptxとreference-before-font.pptxも一致。復元待ちなし。
- Change Case UIは未実装。現在のreplayTextEdits + TextFormat spansでは置換された部分の未知run XML/field情報を保持できないため、そのまま実装しない。既存setShapeText/setTableCellTextの正規経路で文字変換を扱う設計が次の課題。Unicode文脈（Greek sigma）と長さ変化（ß→SS）、混在書式、paragraph properties保持、caret単語境界、outline/tableも必要。
- 同一PR #287を継続。4173の未保存編集と.pnpm-storeには触らない。Mac版との全操作一致は未完了。

## 2026-10-02: 空段落の文字サイズ変更

- 図形全体のIncrease/Decrease Font Sizeで複数空段落の書式が残る不具合を修正。表セル・ゼロ長runにも対応。通常runや未知XMLを作り直さず、既存setShapeTextFormat/setTableCellTextFormatのparagraphEndオプションで段落末を更新する。
- ブラウザー回帰は旧bundleで20/44ptが変化せず失敗、修正版で24/48pt・Undo/Redo・保存再読込成功。既存の混在サイズ・キャレット・図形選択テストも成功。
- 検証: core3207件、site129件、ブラウザー5件（空段落・既存文字サイズ・アウトライン3件）成功。最終ログ `/tmp/empty-font-browser-final.log`。確認プレビュー4175 HTTP200。format/lint/typecheck、core/editor build、Svelte 0 errors/warnings。途中のゼロ長runが後続文字のサイズ判定をずらす不具合も回帰で再現・修正。
- NativeでHomeのChange Caseメニューを確認: Sentence case. / lowercase / UPPERCASE / Capitalize Each Word / tOGGLE cASE。FontダイアログはLatin/Asian font、style/size、color/underline、strike/double strike、superscript/subscript/offset、small/all caps、equalize height。今回は文書の変更なし、Cancel済み。
- Change Case未実装。次はcopyTextRange + TextEdit/replayTextEdits、OutlineSelectionModel.copy/replaceの書式spanを利用可能。ただしUnicode変換でUTF-16長が変わる場合のspan再計算、フィールド等の保持、nativeのキャレット時動作を確認してから実装する。
- 完全一致は未完了。同一PR #287、プレビュー4175を継続。4173の未保存編集と.pnpm-storeには触らない。

## 2026-10-02: Homeの文字サイズ拡大・縮小

- HomeにIncrease/Decrease Font Sizeを追加。選択範囲内の混在サイズをそれぞれ増減し、図形選択・キャレット入力・アウトライン・表セル編集から適用できる。
- Mac参照で10→10.5、13→14/12、44→48/40、96→115、98→118、100→120/96、7.5→7、7→6、6→7、混在20+44→24+48を確認。96ptより上は丸めた20%刻み、8pt未満は整数刻み。参照文書への一時変更は全て取り消し済み、ディスク上もreference-before-font.pptxと一致。
- 検証: site125件、ブラウザー7件成功（選択範囲・図形全体・キャレット連続増減・アウトライン複数項目・Undo/Redo・PPTX再読込、Home日英900/1500/1900px、編集位置）。format/lint/typecheck、Svelte 0 errors/warnings、editor build成功。最終ログ `/tmp/font-step-final-browser.log`。
- 残件: 複数段落内の空段落のendParaRPrは図形全体の増減で更新されない（coreの範囲書式変更が段落終端を変更しない）。空段落のキャレットから次に入力する文字の増減は対応。文字サイズショートカット、Change Case、Fontダイアログ、カーニング描画なども未完了。全操作一致とは扱わない。
- 同一PR #287を継続。確認プレビュー4175はHTTP200。4173の未保存編集と.pnpm-storeは触らない。

## 2026-10-02: Homeの文字間隔メニュー

- HomeにCharacter Spacingを追加。Very Tight=-3pt、Tight=-1.5pt、Normal=0、Loose=3pt、Very Loose=6pt。More SpacingでExpanded/Condensedの数値を指定。混在選択の未変更値は保持し、既存kernも変更しない。
- Mac参照 `/tmp/pptx-outline-audit/reference.pptx` で全プリセットを確認。一時変更は全てUndo済み。ディスク上も `/tmp/pptx-outline-audit/reference-before-font.pptx` と一致、復元待ちなし。この参照文書への一時変更と取り消しはユーザー明示許可済み。
- カーニングはOOXMLで読み書きできるがpreview/inline描画が未対応のため、今回のメニューには追加していない。後続では閾値を含む描画・編集・計測を揃えてからUIを追加する。文字拡大・縮小、Change Case、Fontダイアログの完全一致も未完了。
- 検証: 文字間隔（選択範囲・混在未変更・混在からNormal・既存kern保持・Cancel・Undo/Redo・再読込・キーボード）、Home日英900/1500/1900px、編集位置関連3件のブラウザー計5件成功。site121件、format/lint/typecheck、Svelte 0 errors/warnings、editor build成功。詳細ダイアログ画像 `/tmp/pptx-character-spacing-dialog.png`、最終回帰ログ `/tmp/character-spacing-final-browser.log`。プレビュー4175 HTTP200。
- 同一PR #287を継続。4173の未保存編集と.pnpm-storeは触らない。Mac版との全操作完全一致は未達成。

## 2026-10-02: 編集開始時の文字間隔保持

- 正負の文字間隔が編集開始時に消えて文字幅が変わる不具合を再現（Wide:65.57px→53.77px）。HTMLへの書式変換でletter-spacingを保持し、編集ビューでズーム倍率を適用する修正。
- Mac参照 `/tmp/pptx-outline-audit/reference.pptx` への一時変更と取り消しをユーザーが明示許可。実機でIncrease Font Size44→48pt、Character Spacing > Loose=Expanded3ptを確認。両変更はUndo済み（Undo disabled）。ディスク上の参照は `/tmp/pptx-outline-audit/reference-before-font.pptx` とバイト一致、復元待ちなし。
- 検証: ブラウザー7件成功（正負文字間隔100%/200%、画面上のRange座標とstage不変、clipboard、既存編集、Home）。site121件、format/lint、Svelte 0 errors/warnings、editor build成功。確認用4175 HTTP200。
- Homeの文字拡大・縮小、文字間隔メニュー、Change Caseの追加は未完了。全体完全一致を達成したとは扱わない。同一PR #287、4173の未保存編集を触らない。

## 2026-10-02: 小さい文字と混在サイズの編集位置

- 10pt文字の編集開始時に約6px下へ移動する不具合をブラウザー実測で再現。段落の既定18ptによる行高への影響を除き、各runへ表示用既定サイズを適用。空段落のcaret用サイズは維持。フォント代替もpreviewへ統一。
- 新規text-edit-metrics.test.mjsで10/12/24/44pt、混在サイズ、2段落目の実文字Range座標・幅・高さを比較。関連ブラウザー10件、site121件、format/lint、Svelteチェック、editor build成功。
- ユーザーのスクロール対象はホームと確定。ホーム修正は前コミット86f07712でPR #287へpush済み。確認用4175はHTTP200。4173の未保存編集は触らない。完全なMac版一致は未達成。

## 2026-10-02: 文字編集の位置ずれ・ホームリボン

- 編集開始時の書式バーをキャンバス上の折りたたみパネルへ移し、スライド全体の移動を解消。focusでスクロールしない。文字内部の単一クリックで編集開始しクリック位置へcaret配置、ダブルクリックは単語選択。枠はドラッグ可能。
- サイズ未指定文字が編集時に14pxへ縮小する不具合を修正。previewと同じplaceholder別既定サイズ・テーマフォント・行高を使用。実文字Rangeのx/y/幅/高さが編集前後2px以内、stageと隣接図形不変、ドラッグと保存をブラウザー検証。
- ホームは狭い画面でグループメニューへ折りたたむ。フォント欄は常時、段落は1100px以上で表示。日英900/1500/1900pxで横スクロールなしを確認。右の書式ペインなど全メニューのスクロール廃止は対象外。
- 検証: 関連ブラウザー11件、site121件、format/lint/typecheck、Svelte 0 errors/warnings、core/editor build成功。coreは前回3204件成功/109skipからソース変更なし。
- 同一PR #287。確認用プレビュー4175、未保存編集のある4173と.pnpm-storeは触らない。Mac実機との全操作完全一致は未完了。今回native参照文書への変更なし。

## 2026-10-02: 動画 Crop の数値位置・サイズ

- VideoペインにPicture position（幅・高さ・X/Y offset）とCrop position（幅・高さ・左・上）、Resetを追加。画像と切り抜き枠を独立編集し、Resetは画像全体へ枠を拡張する。
- nativeで横・縦の全項目とResetの保存XMLを比較。Crop height4cm/top3cmはt=1708/b=31625、枠y=1080000/h=1440000。参照trim-reset.pptxはUndo/保存後、`/tmp/pptx-video-ribbon-reset-baseline.xml`とslide XMLがバイト一致。復元待ちなし。
- ブラウザーで8項目、負offset、範囲外入力のエラーと入力復元、Undo/Redo、保存再読込、動画・ポスター・再生設定保持を確認。画像 `/tmp/pptx-video-crop-pane.png`。core3204件成功/109skip、format/lint/typecheck/core・editor build、Svelteチェック成功。
- 未検証: 回転・グループ内のnative Crop座標、ドラッグハンドル・リボンCropメニュー、動画スタイルギャラリーなど。完全一致は未達成。同一PR #287を使用。4173の未保存編集と.pnpm-storeは触らない。

## 2026-10-02: 動画リボン Reset

- Video Format > Adjust に Reset を追加。色補正・塗り・線・効果・3D書式を除去し矩形へ戻す。動画本体・ポスター・crop・寸法・再生設定は保持し、1回のUndo/Redoに対応。色補正のみを戻すペインのResetとは別操作。
- PowerPoint実機でRotated WhiteとBeveled Oval Blackのリセット、cropの保持を保存XMLで確認。証拠 `/tmp/pptx-video-ribbon-reset-{styled,result,oval,oval-result,cropped,crop-result}.xml`。参照資料 `/private/tmp/pptx-poster-audit/trim-reset.pptx` はbaselineとバイト一致、復元待ちなし。
- API `resetShapeVideoFormatting` は動画のみ対象。汎用コマンドも音声・オンライン動画・非メディア・ロック中は無効。未知の拡張XMLと保存再読込をユニット検証。
- core3197件成功/109skip、format/lint/typecheck/core・editor build成功。最終ブラウザー3件成功（動画Reset2件・ヘッダー1件）。後続の汎用コマンド有効条件を含むfocused15件成功、Svelte0 errors/warnings。画像 `/tmp/pptx-video-ribbon-reset.png`。
- ヘッダー要望は `24b84ffa` で対応済み。同一PR287を継続。確認プレビュー4175 HTTP200、4173の未保存編集には触らない。`.pnpm-store/`は触らない。
- 次の動画関連の差分: native Cropの画像幅/高さ/offsetとcrop枠位置、動画スタイルギャラリー、配置ボタンの構成、任意画像の画素一致。全体のPowerPoint完全一致は未達成。

## 2026-10-02: 動画修整スライダーのライブ反映

- Videoペインの明るさ・コントラストをドラッグ中にも反映。確定は1回の履歴として保存し、pointercancel時は取り消す。
- 修正前はmouseup前のSVG transferが0のままでブラウザー回帰失敗。修正後はライブ反映・保存・1回のUndo/Redoが成功。
- format/lint/typecheck、Svelte 0 errors/warnings、site121件、editor build成功。coreは直前の3195件成功から変更なし。
- PR #287を継続使用。ヘッダー省スペース化は反映済み。確認用プレビュー4175、4173の未保存編集には触れない。全体の完全一致は未達成。

## 2026-10-02: Washoutの白潰れ修正

- 前項の描画式の差を修正。LibreOffice公式実装 `vcl/source/bitmap/bitmap.cxx` のMSO互換処理に基づき、明るさをコントラストの前後に半分ずつ適用。輝度補正のみsRGB指定。ポスター・再生動画・修整候補・Washout候補へ適用。
- 動画ブラウザー回帰は旧式の暗部279（白へクリップ）で失敗、新式217で成功。+20/+20修整の保存・Undoも成功。Svelte 0 errors/warnings、site121件成功。
- 静止画の画素回帰（暗部217±1・白255）成功。core3195件、format/lint/typecheck、core/editor build成功。
- 任意画像の画素完全一致は引き続き未証明。実機資料の新しい一時変更なし。PR #287を継続使用。

## 2026-10-02: 動画テーマ色の濃淡とサムネイル

- More Variationsにテーマ色の濃淡5段を追加（基本10色＋濃淡50色＋標準10色）。schemeClrとlumMod/lumOffを保持し、動画用tint/satModを適用。実機Accent1の淡色80%・濃色25%の保存XMLと一致。カスタムテーマの濃淡規則は未検証。
- グレースケール・セピア・テーマduotone・白黒閾値のサムネイルをSVGに変更。白黒3候補の表示差、保存、Undo/Redo、再読込をブラウザーで検証。修整→色のリボン順序も実機に合わせた。
- 検証: format/lint/typecheck、Svelte 0 errors/warnings、site121件、editor build、動画recolor回帰成功。画像 `/tmp/pptx-video-recolor-pane.png` を確認。ヘッダー日英900/1500pxも再検証成功。
- 未解決: 既存poster renderer / MediaInlinePreviewのbrightness/contrast式はWashout (.7,-.7)でslope=.3/intercept=1.05となり全白になる。実機では暗部#262626が約#D9D9D9、白は白のまま。正確な式は複数画素で実機比較が必要。Washoutサムネイルは従来のCSS近似を維持した。画素完全一致は未達成。
- 実機操作はWashout適用後Undo/保存済み。trim-reset.pptxのslide XMLが `/tmp/pptx-video-reset-baseline.xml` と一致、復元待ちなし。PRは#287。4173の未保存編集には触れない。

## 2026-10-02: 動画の色変更ギャラリー

- リボンColorと動画ペインから21プリセット、基本テーマ色・標準色を選択可能。保存、Undo/Redo、再読込、ライブ動画への反映を実装。詳細ペインの修整と色変更プリセットには別々のアクセシブル名を設定。
- Mac実機で保存XMLを比較。Washoutはlum bright=70000/contrast=-70000、Darkはtint45000/satMod400000、Lightはshade45000/satMod135000。セピア、白黒、基本テーマ色も確認。参照資料はUndo後に保存し、baseline XMLとの完全一致を確認。復元待ちなし。
- 検証: core3194件、site121件成功。format/lint/typecheck、Svelte 0 errors/warnings、core/editor build成功。ヘッダー日英900/1500pxと色変更のブラウザー回帰成功。
- 確認用プレビュー http://127.0.0.1:4175 をChromeで表示。4173の未保存編集は触らない。引き続きPR #287に集約。
- 未完了: More Variationsのテーマ濃淡5段、サムネイルの画素一致、リボンReset、スライダー操作中の連続反映、ネイティブCrop配置。全体の完全一致は未達成。

## 2026-10-02: 動画ペインの色リセット

- VideoタブのResetを実装。grayscl/duotone/biLevel/lumのみ解除し、表紙・動画・図形書式・透明度・未知の拡張を保持。Undoは1操作。リボンResetは図形書式まで解除する別動作で未実装。
- 実機でグレースケール/セピア/白黒と明るさ/コントラストを比較。参照資料 `/private/tmp/pptx-poster-audit/trim-reset.pptx` は全一時操作をUndoして保存、slide XMLが `/tmp/pptx-video-reset-baseline.xml` と完全一致。復元待ちなし。
- ヘッダーの省スペース化は反映済み。確認用プレビュー http://127.0.0.1:4175 。4173の未保存編集には触れない。PRは引き続き #287 に集約。
- 検証: 動画Reset・修整・リボンのブラウザー3件成功。非既定の塗り/線/透明度を持つResetケースも再検証済み。site121件、Svelte 0 errors/warnings、format/lint/typecheck、core/editor build成功。全core検証で検出した新APIのmanifest登録漏れも修正。
- 次はRecolor、リボンReset、ドラッグ中の連続プレビュー、Cropの実機配置。完全一致は未達成。

## 2026-10-02: 動画書式の詳細ペイン

- 検証: 動画修整・動画リボンのブラウザー2件、ヘッダー省スペース表示1件、site121件成功。Svelte 0 errors/warnings、format/lint、editor build成功。幅の狭いペインではみ出す数値欄も回帰テストで再現後修正し、画像 `/tmp/pptx-video-format-pane.png` を確認。確認用プレビュー4175の稼働を再確認。
- Mac実機でFormat Videoの4タブと、Video内のBrightness/Contrastプリセット・スライダー・数値欄を再確認。専用Videoタブと修整オプションからの導線を実装。図形へ選択を移した際のタブ復帰も対応。
- 回帰テストは変更前にFormat Videoタブリスト不在で失敗することを確認。保存値、Undo、スライダーと数値欄の同期、タブのキーボード操作、図形への選択切替を検証対象に追加。
- 未完了: Recolor、Resetの意味、ドラッグ中の連続プレビュー、Cropの実機配置。Cropは実機でPicture position (Width/Height/Offset X/Offset Y) とCrop position (Width/Height/Left/Top)を確認。既存の4辺百分率UIとは異なる。
- 今回実機は読み取りとCrop展開のみ。文書内容変更なし。完全一致は未達成。

## 2026-10-02: 動画の修整ギャラリー

- 検証: 修正前はCorrectionsボタン不在で回帰失敗。修正後は動画修整（保存値・原子的Undo・Normalリセット・ライブ動画filter・キーボード・日本語）と既存動画リボンの2件成功。site121件、Svelte 0 errors/warnings、format/lint/typecheck、editor build成功。画像 `/tmp/pptx-video-corrections-gallery.png` を目視確認。

- Mac実機のVideo Format > Correctionsは5×5の候補（列: 明るさ -40/-20/0/+20/+40%、行: コントラスト同値）。共通の風景サムネイルと選択枠、Movie Correction Options...を持つ。
- +20/+20を一時適用して保存し、`a:blip/a:lum bright="20000" contrast="20000"`を確認。Undoして保存、lumが消えたこととUndo無効を確認済み。参照資料 `/tmp/pptx-poster-audit/trim-reset.pptx` に未復元の内容変更なし。
- 修整ギャラリーを既存の明るさ・コントラストAPIへ接続。両値を1つのUndo操作で保存。編集中のHTML動画にも表紙SVGと同じ補正を適用。
- 詳細ペインのネイティブUI（Videoタブ、スライダー等）、色/スタイル/Crop/独立Reset、補正後の画素レベルの実機一致は引き続き未検証。完全一致は未達成。

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
