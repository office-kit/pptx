# Mac PowerPoint parity integration

The final review target is PR #287 (`feat/pptx-editor`). No additional PR is required.

The earlier implementation remains preserved on `feat/mac-powerpoint-parity` at `ef67cf0`. It and #287 developed different UI, persistence, animation and rendering implementations after their common base. A trial merge produced 39 conflicted files and was aborted without discarding either branch. Transfer compatible features individually, retaining #287 behavior and testing each integration. The earlier branch is not fully merged.

## Integrated: presentation guides and grid settings

- Public APIs read/write drawing guides, stored guide visibility, grid spacing and snapping. Preserve related view parts, surviving guide metadata and unrelated view settings. Extended guide positions use native master units; the public API uses EMU.
- Registered the four mutations in the existing editor command catalogue, with English/Japanese labels and structured fields for guides and grid spacing. The canvas integration below uses these document settings.
- Six API tests and six capability/localization coverage tests pass. Root TypeScript and targeted lint/format checks pass.

## Outstanding

Reconcile the remaining operations from the earlier branch. Continue native visual/interaction comparison. All-operation Mac PowerPoint UI parity remains incomplete. The older branch contains the detailed comparison history in its version of this file.

Current unresolved areas (2026-10-03; the sections below retain the comparison history):

| Area                     | Remaining work                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| ------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| WordArt                  | Gallery: Shape Format ▸ WordArt Quick Styles and Insert ▸ WordArt show the 20 native presets as swatches in native order; all 20 apply the captured payload (replacing, not merging): run colors keep their scheme token and tint transforms (`<a:schemeClr val="accent2"><a:lumMod/><a:lumOff/>`), and the two bevels write the captured `<a:scene3d>`/`<a:sp3d>` under `<a:bodyPr>`. A unit test compares every preset's run, paragraph end and body 3-D with its native capture. Deviations: `b`/`spc` are written as `0` instead of removed; labels are the captured Office 2007-theme names, not theme-dependent; a preset without a bevel removes an earlier preset's bevel, and Clear WordArt removes outline, effects and 3-D but keeps the fill — neither is confirmed against a native capture. The canvas paints the gradient and pattern presets' glyph fills, both in the slide and while editing text: one gradient spans the whole text block (every line, not each line separately), and the pattern uses the shape-fill tiles. Not yet compared against native screenshots: path gradients are elliptical approximations, `lin scaled` is not distinguished, and pattern tiles start at the slide origin rather than the text. The two bevels (and any `<a:bevelT>` on a text body) are shaded on the canvas, in the slide and while editing, shape text and table cells alike: an SVG lighting filter over the glyphs, sized from the bevel's `w`/`h`, lit from the rig's direction and revolution, with a coarse per-preset edge profile and material highlight. This is a 2-D approximation not yet compared against native screenshots: one distant light stands in for the rig's lights, the camera is always drawn head-on, extrusion and contour are not drawn, and bevels inherited from the layout's body properties are ignored. Remaining: complete reflection rendering, selected-range application, native swatch artwork. |
| Outline                  | Show Formatting persistence, Thesaurus/Translate and native comparison of cross-slide selection and drag gestures (metrics and menu order: see "Views").                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| Ruler and tabs           | Complex wrapping, live-drag snapping, how an indent drag applies to mixed paragraphs, and the ruler's white text-area band. Mixed-marker display, the shape-selected ruler, rotated/vertical mapping and locale decimal separators were verified against Mac PowerPoint on 2026-10-07 (see "Ruler and texture: native verification").                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| Menus and views          | Home ribbon, Format Shape pane rows, object/thumbnail menus, notes pane and status bar now follow the native geometry measured on 2026-10-07 (see "Native geometry audit"), as do the Insert, Draw and Design tabs (see "Insert, Draw and Design ribbons"), the Transitions, Animations, Slide Show, Record, Review and View tabs (see "Show and review tabs") and the Shape Format / Picture Format / Table Design / Table Layout tabs (see "Contextual tabs: geometry audit"). The Effects sections, Shape Options / Text Options and collapsed sections followed on 2026-10-07 (see "Format Shape pane: Effects and Text Options"). The text-editing, table-cell, picture and slide-background context menus followed on 2026-10-07 (see "Text, table-cell, picture and slide-background menus"). The menu bar and its keyboard shortcuts followed on 2026-10-07 (see "Menu bar and keyboard shortcuts"). Still open: additional native views.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| Image/background effects | Texture gallery present with generated approximations (not Microsoft’s bitmaps; native visual comparison pending), remaining image-effect rendering and native original/rendered-image correction handling. The session "last texture" default was verified on 2026-10-07; whether the ribbon's Shape Fill ▸ Texture also sets it is assumed.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| Native verification      | Zoom, Home ribbon, Format pane and context-menu geometry were measured through the accessibility API on 2026-10-07. Recheck other tabs, galleries, panes and interaction cases. Keep bridge failures separate from observed PowerPoint behavior.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Release review           | Verify subsequent changes against ordinary build artifacts and CI, and complete the whole-PR maintainer review before claiming readiness.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |

Historical statements that a feature is missing can be superseded by a later section;
for example, outline collapse and ruler tab gestures are now implemented. This list
does not define a percentage denominator for all PowerPoint operations.

## Native geometry audit (2026-10-07)

Measured side by side with Mac PowerPoint 16 in a 1512 × 900 pt window (dark appearance, the system setting at the time; geometry does not depend on it), on a deck PowerPoint created itself and closed without saving. Native positions come from the accessibility API (System Events / AXUIElement), so they are exact points; the editor's come from the DOM at the same CSS size. Context menus were opened with one posted right-click while PowerPoint was briefly frontmost and dismissed by reactivating the previous app; nothing was typed. `native-geometry.test.mjs` asserts the matched metrics.

- **Home ribbon.** Native: tabs 26 pt; a 72 pt command row 7 pt below them; groups separated by a rule with 10 pt each side; large buttons 50 pt wide filling the row (Reset 38, Add-ins 45, Designer 51); small buttons 26 × 26; menu buttons with an arrow 38 × 26; two-row groups (Font, Paragraph) on a 32 pt pitch starting 4 pt down; font box 174 pt and size box 56 pt; Add-ins and Designer are separate groups. Editor before: an 82 pt row, 28 × 26 small buttons, 20–24 pt tall font/size boxes, 30 pt font-size steps, Insert's labels at two heights and Add-ins/Designer in one group, so at 1512 pt the groups ended 104 pt further right. Now matched; the Home groups end at 1501 pt (native 1471) and the collapse thresholds are unchanged. The second Font row is now in PowerPoint's order — Bold … Subscript, Character Spacing ▾, Change Case ▾, a rule, Text Highlight Color ▾, Font Color ▾ — and the editor's extra Font dialog (A…) button is gone (Cmd+T and Character Spacing ▸ More Spacing… open the dialog; PowerPoint uses its Format menu). Arrange and Quick Styles put ▾ beside the icon as natively. Left: Bullets/Numbering stay plain 26 pt toggles (native 38 pt split buttons with galleries), icons and label fonts are the editor's own, and the title bar's Home button and the window's native title bar are not reproduced.
- **Window layout.** Native: thumbnail pane 249 pt + 5 pt splitter; notes pane one line (39 pt + 5 pt splitter); status bar 28 pt. Editor before: 200 pt thumbnails, a 120 pt notes pane, a 26 pt status bar. Now 249 / 44 / 28.
- **Zoom.** Native 100% shows one slide point per screen point (a 13.33 in slide is 960 pt wide), and Fit leaves about 22 pt around the slide, so this window fits at 120%. The editor's 100% was 96 px per inch (the same slide was 1280 px), so Fit read 76%. Zoom percentages now mean what they mean in Mac PowerPoint (100% = 960 px for that slide) and Fit leaves 22 pt, reading 119% here; the 10–400% range, the ±10 steps and the Zoom dialog presets are unchanged in meaning. Status bar: Notes, Comments, four 36 pt view buttons, 14 pt −/+, a 102 pt slider, a 54 pt percentage and a 38 pt Fit button, 13 pt from the right edge; the editor matches within 1–13 pt (its Accessibility item and wider language label move Notes/Comments slightly left). Text is 12 px instead of 11.
- **Format Shape pane.** Native: 300 pt wide, 27 pt title row, Shape Options / Text Options switch, 44 × 42 pt category tabs, 25 pt section headers with a chevron, all sections collapsed on first open; options as radios on a 20 pt pitch, a rule, then label-left / control-right rows on a 30 pt pitch with 26 pt controls ending 17 pt from the pane edge (color buttons 39 pt, sliders 113 pt with a 66 pt box, pop-ups 112 pt, Size/Position boxes 82 pt). Editor before: a 41 pt title, 38 × 34 tabs, labels stacked above their controls, 19–23 pt controls of varying widths, a No outline button and no line-type radios. Now: 27 pt title, 42 × 42 tabs, chevron headers, one-line Color / Transparency / Width rows, 26 pt controls, 112 pt pop-ups, 82 pt Size/Position boxes, and the Line section starts with No line / Solid line / Gradient line (since enabled, with the Gradient fill controls for the line) and hides its controls for No line, as natively. Left at the time: Shape Options / Text Options, the Effects sections, collapsed sections, Sketched style, 39 pt arrow buttons and the extra flip checkboxes — all addressed in the next section.
- **Context menus.** Native items are 24 pt and separators 11 pt. The object menu is Cut, Copy, Paste | Edit Text, Edit Points | Reorder Objects, Reorder Overlapping Objects | Group ▸ (Group, Regroup | Ungroup), Bring to Front ▸, Send to Back ▸, Lock | Hyperlink... | Save as Picture... | Translate... | View Alt Text..., Set as Default Shape Style, Size and Position..., Format Shape... | Action Settings... | New Comment. The thumbnail menu is Cut, Copy, Paste, Select All | New Slide, Duplicate Slide, Delete Slide, Add Section | Format Background... | Hide Slide | Zoom..., Slide Show | New Comment — with no Layout ▸ or Reset Slide, unlike the 2026-10-06 notes below. Both menus now match, with separators as separate 11 pt rules; Lock/Unlock, Reorder Overlapping Objects, Hyperlink..., Action Settings..., Select All, Zoom... and Slide Show work, and Edit Points, Reorder Objects, Save as Picture..., Translate... and Set as Default Shape Style are disabled with the reason as their tooltip. Japanese labels follow PowerPoint's own `ja.lproj` strings where the editor had none. Left: the text-editing menu (Cut, Copy, Paste | Exit Edit Text, Font..., Paragraph..., Bullets ▸, Numbering ▸ | Thesaurus..., Translate... | Format Text Effects..., Format Shape... | Lock | Hyperlink... | New Comment) — inline editing still shows the browser's menu, because the editor's menu takes focus from the text editor; accelerators show the editor's own shortcuts (⌘D, not PowerPoint's ⇧⌘D).

## Contextual tabs: geometry audit (2026-10-07)

Shape Format, Picture Format, Table Design and Table Layout, measured the same way as the Home tab above (accessibility API, 1512 × 900 and 1200 × 900 pt windows; the Picture Format capture at 1200 pt shows the Home tab, so that width is inferred from the others). They share the Home tab's conventions: a 72 pt row, a rule with 10 pt each side between groups, large buttons filling the row (50 pt with a ▾ beside the icon, 38–42 pt without), 22 pt rows with an icon and label, 36 × 22 icon menus, 26 pt small buttons on a 32 pt pitch, 18 × 58 pt gallery arrows and 78 × 24 pt spin boxes. Captions wrap onto at most two balanced lines ("Insert Row" / "Above"); Japanese captions break after a particle or at a katakana word ("図形の" / "塗りつぶし", "代替" / "テキスト"). Below 1300 pt each tab switches to its compact layout. When labels need a few more points than native (Japanese captions, or the wider system font on Linux), the in-ribbon galleries narrow first, as PowerPoint's do, so the tabs keep their full layout at 1512 pt; only an overflow the galleries cannot absorb takes the compact step earlier. `contextual-ribbon-geometry.test.mjs` asserts group order and starts (within 9 pt) and control sizes at both widths, plus the Japanese labels; off macOS, where the font is not San Francisco, labelled controls may be up to 20% wider and group starts may drift by 8%.

- **Shape Format.** Native: shape strip (6 × 3 cells of 20 × 18 pt in a 120 pt frame) with Text Box ▾ / Edit Shape ▾ / Merge Shapes ▾ rows; Quick Styles strip (three 58 pt swatches in 174 pt), Shape Fill (large), Shape Outline ▾ and Shape Effects ▾ (38 × 26); the WordArt strip, Text Fill, Text Outline ▾, Text Effects ▾; Alt Text; Arrange expanded as Bring Forward ▾, Send Backward ▾, Selection Pane, Reorder Objects ▾, Align ▾ and Group ▾ / Rotate ▾; Height, Width and Lock Aspect Ratio; Format Pane. At 1200 pt Insert Shapes is Shapes plus three icon menus and Arrange is one button. Editor before: a Shapes button and labelled small rows, 40 × 30 swatches, no WordArt strip, the collapsed Arrange menu at every width, spin boxes with a "cm" suffix and no aspect lock. Now matched: groups start within 1–3 pt of native at 1512 pt and the tab ends at 1498 (native 1496). The WordArt strip shows the first three presets and its › opens the twenty-preset gallery (WordArt Quick Styles); the shape strip's arrows page through the gallery's shapes rather than opening it; Text Box ▾ offers Draw Horizontal Text Box and Vertical Text Box (`eaVert`). Left: spin boxes show the number without "cm"; Reorder Objects ▾ only offers Reorder Overlapping Objects; Merge Shapes stays disabled.
- **Picture Format.** Native: Remove Background; Adjust (Corrections large, Color ▾ / Artistic Effects ▾ / Transparency ▾ rows, then Compress Pictures, Picture Quality ▾, Change Picture ▾ and Reset Picture ▾ icons); Picture Styles (212 pt strip, Picture Border ▾ / Picture Effects ▾ / Picture Layout ▾ rows); Alt Text; Arrange; Size (Crop ▾, "Height:" / "Width:" boxes, lock); Format Pane; Animate as Background. The editor showed Shape Format for pictures. Now a picture shows this tab (videos keep Video Format, audio keeps Shape Format and Playback) with groups within 2–8 pt of native. Corrections and Color reuse the Video Format galleries; Transparency offers PowerPoint's 0–95% presets; Change Picture ▸ From a File..., Reset Picture (corrections, recolor and transparency), Picture Border, Picture Effects (Shadow, Glow, No Effects) and Crop ▸ Crop / Crop to Shape work. Disabled with the reason: Remove Background, Artistic Effects, Compress Pictures, Picture Quality, the picture style strip, Picture Layout, Reset Picture & Size and Animate as Background. At 1200 pt the Color/Artistic Effects/Transparency and Picture Border/Effects/Layout rows lose their labels and Arrange collapses.
- **Table Design.** Native: Table Style Options (Header Row, Total Row, Banded Rows / First Column, Last Column, Banded Columns); Table Styles (518 pt strip of 76 × 56 pt swatches, Shading ▾ / Borders ▾ / Effects ▾ rows, icon-only at 1200 pt); WordArt Styles (Quick Styles large, Text Fill / Text Outline / Text Effects icons); Draw Borders (Pen Style and Pen Weight pop-ups 120 × 22, Pen Color ▾ 92 × 22, Draw Table, Eraser). Editor before: Table Styles (an ID prompt), Shading and Borders only. Now all are shown in that order, groups within 6 pt of native. The check boxes write `a:tblPr` flags; the strip shows the gallery row holding the table's style and marks it, and › opens the whole gallery (see "Table Styles gallery" below); Borders ▾ draws the pen on the chosen edges of the selection (all, outside, inside, each side, inside horizontal/vertical, diagonals) and No Border removes the cells' explicit borders; WordArt styles, Text Fill/Outline/Effects format the selected cells. Left: No Border cannot write PowerPoint's explicit `a:noFill` edge, so a style's own borders still show; Effects ▸ Cell Bevel, Draw Table and Eraser are disabled; Shading has no No Fill entry.
- **Table Layout.** Native: Table (Select ▾, View Gridlines); Rows & Columns (Delete ▾, a rule, Insert Row Above 60, Insert Row Below 60, Insert Column Left 68, Insert Column Right 74 — the last three as 22 pt rows at 1200 pt); Merge (Merge Cells, Split Cells); Cell Size (Row and Column boxes, Distribute Rows 113 / Distribute Columns 131 pt, icon-only at 1200 pt); Alignment (six 26 pt toggles, Text Direction ▾, Cell Margins ▾); Table Size; Arrange (Bring Forward ▾, Send Backward ▾, then Align ▾ / Group ▾ / Rotate ▾ as labelled 22 pt rows); Format Pane. The tab was called Layout and had Delete/Insert Rows/Columns, Merge Cells, Height, Width and Alignment buttons that opened prompts. Now matched in order and size (groups within 6 pt at both widths). Select ▾ selects the cell, column, row or table; Delete ▾ removes the selected columns, rows or the table; inserts go above/below/left/right of the selection; Row/Column boxes and Distribute act on the selected rows or columns (all of them when the table is selected) and Table Size scales them, keeping the frame the size of the grid; Text Direction offers Horizontal, Rotate all text 90°/270° and Stacked; Cell Margins offers Normal, None, Narrow and Wide (Custom Margins... opens the Format pane). Left: View Gridlines is shown pressed and disabled (the editor always draws gridlines); Split Cells only splits merged cells (PowerPoint's Split Cells dialog divides any cell); inserting and deleting rows or columns still requires splitting merged cells first.
- **Japanese.** Labels follow PowerPoint's `mso40ui` `tbbtn.strings`: 図の形式, テーブル デザイン / テーブル レイアウト, 上に行を挿入, 高さを揃える, セル内の配置, タイトル行, 縞模様 (行), 背景としてアニメーションを付ける and so on. Format Pane changed from 書式設定ウィンドウ to 書式ウィンドウ and Reorder Objects from オブジェクトの順序変更 to オブジェクトの並べ替え to match.

## Table Styles gallery (2026-10-07)

PowerPoint's 74 built-in table styles, identified by GUID, are now applied, written and drawn.

- **Definitions.** `test/fixtures/native/builtin-table-styles.xml` holds PowerPoint for Windows' own `a:tblStyle` serialization of every style, in gallery order (captured through its Table Styles gallery by the Dewiride PPTX Studio project, Apache-2.0). It was checked against Mac PowerPoint 16.113: the 74 GUIDs in its binary appear in the same order, the 14 style templates embedded in the app match byte for byte (one writes an empty `fontRef` with an end tag instead of `/>`), and a deck Mac PowerPoint saved with nine of the styles matches exactly. `setTableStyleId` writes a style's definition into `ppt/tableStyles.xml` when the deck lacks it, as PowerPoint does.
- **Rendering.** Every part of a table style is drawn: whole table, banded rows and columns, first/last column, header and total rows, and the four corner cells, in PowerPoint's order (MS-OI29500 §2.1.1265), with the cells' own formatting on top; translucent band fills; and the Themed Styles' background (the theme's style-matrix fill under the table). Against Mac PowerPoint 16.113's exports of a deck with all 74 styles (header row and banded rows on, Office 2007 theme), the sampled cell colors of 62 styles match exactly and 66 within 4 levels per channel. Left: the Themed Styles' background is the theme's gradient, whose saturated stops (`satMod` 300–350%) come out lighter than PowerPoint's for orange accents (Accent 6 differs by up to 53 levels, other accents by up to 19); the background's `effectRef` shadow is not drawn.
- **Gallery.** Native: the strip shows one row of seven 76 × 56 pt swatches; › opens the gallery with Best Match for Document (No Style, No Grid; Themed Style 1; No Style, Table Grid; Themed Style 2), Light, Medium and Dark, seven styles to a row, then Clear Table. Swatches are a 5 × 5 table drawn by the preview renderer in the slide's theme with the table's six style options, so they follow the check boxes. Names follow PowerPoint's own strings in English and Japanese (e.g. 中間スタイル 2 - アクセント 1, スタイル (淡色) 1, 濃色スタイル 2 - アクセント 1/アクセント 2, 表のクリア). Clear Table applies No Style, No Grid and removes the cells' own fills and borders. Left: the Best Match group lists the same 14 styles for every theme; whether PowerPoint's Clear Table also resets cell fills and borders was not re-measured against PowerPoint.

## Insert, Draw and Design ribbons (2026-10-07)

Measured like the Home ribbon (accessibility API, 1512 × 900 and 1200 × 900 pt windows) and asserted by `native-ribbon-geometry.test.mjs`. Native buttons overhang their group by 1 pt, so between two groups' buttons there are 19 pt (9 + rule + 9), not 21; these three tabs use 9 pt, which puts every command within 0–4 pt of PowerPoint (the Home tab still uses 10). Large buttons are at least 38 pt wide, or 50 pt with ▾ beside the icon, and a caption of more than one word takes two lines broken where they balance ("Header & / Footer", "Ink to / Text"; Text Box stays on one line). Japanese captions break after の/と/を/に or where the script changes ("新しい / スライド", "インクを / テキストに変換"); Japanese labels follow PowerPoint's `ja.lproj` strings, and the Insert ▸ Object caption is now オブジェクト (it was 対象, shared with the Animation pane, which now reads オブジェクト too).

- **Insert.** Native groups: New Slide | Table | Pictures, Screenshot | Cameo | Shapes, Icons, 3D Models, SmartArt, Chart | Zoom, Link, Action | Comment | Text Box, Header & Footer, WordArt, Date & Time, Slide Number, Object | Equation, Symbol | Video, Audio, ending at 1330 pt; at 1200 pt 3D Models, SmartArt and Chart become 22 pt rows (icon, name, ▾) and Date & Time, Slide Number and Object 24 × 22 icons. Editor before: generic 52 × 66 buttons with no ▾ and no 1200 pt layout. Now a dedicated InsertRibbon with the same groups, widths and 1200 pt layout (it switches below 1350 pt, or while the row overflows). ▾ menus: New Slide (layouts, Duplicate Selected Slides), Table (a 10 × 8 grid that inserts that size with header and banded rows, Insert Table..., Draw Table disabled), Pictures (Photo Browser..., Picture from File..., Stock Images..., Online Pictures...; only Picture from File... works), Screenshot (Available Windows: the browser's own picker; Screen Clipping disabled), Chart (PowerPoint's 15 categories; Column, Line, Pie, Bar, Area and Radar open the chart dialog on that type, the others are disabled), Text Box (Draw Horizontal / Vertical Text Box; vertical writes `vert="vert"`), Video (Movie Browser..., Movie from File..., Online Movie...) and Audio (Audio Browser..., Audio from File..., Record Audio...). Shapes, WordArt, 3D Models, SmartArt, Zoom and Equation show ▾ as natively (galleries or disabled). Left: the Insert menus other than Design's were not captured natively, so their item lists come from PowerPoint's string tables rather than a measured menu; Japanese at 1200 pt still scrolls about 70 pt (PowerPoint's Japanese collapse was not captured); New Slide and Text Box are split buttons (click inserts, ▾ opens the menu) where PowerPoint draws the text box with the mouse.
- **Draw.** Native: Draw 38, Eraser ▾ 50, Lasso Select 38 | a 180 × 60 pt pen gallery 18 pt inside its group, Add ▾ 50 | Ink to Text / Shape / Math, 38 each | Draw with Trackpad 62; nothing collapses at 1200 pt (the gallery narrows to 168 pt). Editor before: 52 pt buttons, a free-width gallery and an "Add Pen" button. Now matched; Eraser ▾ offers Stroke Eraser (what the editor's eraser does) with Point and Segment Eraser disabled, and Add ▾ offers Add Pen, Add Pencil and Add Highlighter (pencil and highlighter keep the default gallery's width and opacity). Left: the gallery stays 180 pt at 1200 pt; the pen icons are the editor's own.
- **Design.** Native: the Themes gallery is 18 pt Previous/Next around whole 95 pt slots (85 × 48 pt thumbnails) and takes the width the other groups leave (ten slots at 1512 pt, seven at 1200 pt); then Variants ▾, Colors ▾, Fonts ▾, Background Styles ▾ | Layout ▾, Slide Size ▾ | Design Suggestions (a 69 pt toggle). Menus: Variants is a row of four 97 × 60 variants; Colors lists Theme Colors (the deck's) then All Colors (24 sets: Office, Office 2013 - 2022, Office 2007 - 2010, Grayscale, Blue Warm, Blue, Blue II, Blue Green, Green, Green Yellow, Yellow, Yellow Orange, Orange, Orange Red, Red Orange, Red, Red Violet, Violet, Violet II, Median, Paper, Marquee, Slipstream, Aspect) on a 27 pt pitch, then Customize Colors... and Reset Slide Theme Colors; Fonts lists Theme Fonts then All Fonts (26 pairs) as 62 pt "Aa" rows; Background Styles is a 4 × 3 grid of 84 × 66 pt cells with Format Background... and Reset Slide Background; Slide Size has Standard (4:3) / Widescreen (16:9) radios and Page Setup...; Layout is a 5-column gallery of 100 × 77 pt layouts. Editor before: a fixed six-theme page with ⌃/⌄, ▾ in the captions, 8 color sets and 8 font pairs. Now matched, with the color and font values read from Office's own `Theme Colors/*.xml` and `Theme Fonts/*.xml`; Office 2007 - 2010 fonts are Calibri / Calibri as PowerPoint lists them (they were Cambria / Calibri). Reset Slide Theme Colors is disabled (the library writes no per-slide theme override). Left: Variants stays disabled — Office Theme has four variants natively, which recolor the theme and switch to dark backgrounds, and the library has no variant writer; the gallery still holds the editor's seven color/font presets rather than Office's .thmx themes, so most slots are empty; the editor keeps Customize Fonts... at the end of Fonts (PowerPoint's Fonts menu has no command after the list); the Layout menu has no theme-name heading.

## Format Shape pane: Effects and Text Options (2026-10-07)

Compared with the accessibility dumps and screenshots of Mac PowerPoint 16 captured the same day (pane, preset popovers and Text Options, 1512 × 900 pt window); PowerPoint was not driven for this change. `format-pane-effects.test.mjs` covers the behaviour.

- **Shape Options / Text Options.** Native: a 278 × 26 pt segmented switch 14 pt under the title, 134 × 24 pt segments; Text Options has Text Fill & Outline, Text Effects and Textbox categories (44 × 42 pt), with Text Fill (No fill, Solid fill, Gradient fill, Picture or texture fill, Pattern fill; Color, Transparency) and Text Outline (No line, Solid line, Gradient line) sections. Now matched and applied to every run of the selected shapes. No fill (`<a:noFill/>`), Picture or texture fill (the default texture, then Insert..., Clipboard or Texture; the picture is stretched over the text as PowerPoint writes it) and Gradient line (the Gradient fill controls: type, direction, angle, stops, Rotate with shape, plus Width) are enabled. Gradient and Pattern text fill apply a default fill; their stop/pattern editors are not in the Text Fill section yet. Not offered for text pictures: Transparency, Tile picture as texture and the offsets (the library writes only the stretched picture). The default gradient line (accent 1, light to dark, as for Gradient fill) is assumed, not captured natively.
- **Collapsed sections.** Native: every section starts collapsed and keeps what was opened. Now the Format Shape pane's sections (Fill, Line, the six Effects sections, Size, Position, Text Box and the Text Options sections) start collapsed and are remembered for the session; Size and Position... opens Size and Position. Format Background, Format Picture's crop and Format Video sections are unchanged.
- **Effects.** Native order Shadow, Reflection, Glow, Soft Edges, 3-D Format, 3-D Rotation on a 25 pt header pitch. Shadow: Presets and Color (39 × 26 pt), Transparency, Size, Blur, Angle, Distance (113 pt sliders, 66 pt boxes); Reflection: Presets, Transparency, Size, Blur, Distance; Glow: Presets, Color, Size, Transparency; Soft Edges: Presets, Size. 3-D Format: captioned 59 × 60 pt Top bevel / Bottom bevel galleries with Width and Height, Depth and Contour color buttons with Size, Material and Lighting galleries (Lighting with Angle), Reset (61 × 26 pt). 3-D Rotation: a 59 × 60 pt Presets gallery, X / Y / Z Rotation and Perspective boxes with two 32 pt nudge buttons each, Keep text flat, Distance from ground, Reset. All implemented in that order and layout; preset popovers have PowerPoint's headings (No Shadow + None, Outer, Inner, Perspective; Reflection Variations; Glow Variations; Soft Edge Variations; Bevel; Standard / Special Effect / Translucent; Neutral / Warm / Cool / Special; Parallel / Perspective / Oblique) and names. Shape effects use new core setters (`setShapeInnerShadow`, `setShapeReflection`, `setShapeSoftEdge`, `null` removal on `setShapeShadow` / `setShapeGlow`, shadow scale/skew, `setShape3D`); Text Effects writes the runs' effects and the text body's 3-D. Each edit is one undo step and a multiple selection changes in one transaction.
  - Native-verified values: Offset: Bottom Right reads 60% / 100% / 4 pt / 45° / 3 pt, Tight Reflection: Touching 48% / 35% / 0 pt, a 5 pt glow 5 pt / 60%; preset names and gallery groupings come from the popover dumps.
  - Assumed, not verified natively: the XML of the Inner, Perspective and remaining Outer shadow presets, the reflection blur (6350 EMU; native shows 0 pt), Half/Full reflection sizes (55% / 90%), glow `satMod` 175%, the 3-D nudge steps (10°, field of view 5°) and directions, the X / Y / Z mapping to camera `lon` / `lat` / `rev`, and the Japanese names of the reflection, glow, soft-edge, bevel, material, lighting and rotation presets. The Japanese pane labels, shadow preset names (オフセット: 右下, 内側: 左上, 透視投影: 左上), arrow labels (始点矢印の種類 …), 3-D nudge buttons (時計回り, 狭角, 広角) and the Textbox tab (テキストボックス) were checked against `native-ref/ja`. In Japanese, native right-column controls sit 17 pt further left (x = 1439 instead of 1456) because the labels are wider; the editor keeps the English column. Glow preset names omit the theme color name PowerPoint inserts ("Glow: 5 point; Dark Teal, Accent color 1" natively). Preset tiles are CSS approximations (shadow, reflection, glow, soft edge) or text (bevel, material, lighting, rotation), not PowerPoint's images; X / Y / Z show 0° for presets that imply a rotation without writing one.
  - Soft Edges under Text Effects is disabled, as in Mac PowerPoint (the reason says PowerPoint offers none for text, although the schema allows a run `<a:softEdge>`). Keep text flat writes `<a:bodyPr><a:flatTx/>` for shapes and text and is disabled for pictures and connectors, as natively; the preview does not render 3-D rotation, so it has no visible effect there.
- **Line section.** Native: Width, Sketched style (39 pt menu: None, Curved, Freehand, Scribble), Compound type, Dash type, Cap type, Join type, Begin/End Arrow type and size as 39 × 26 pt gallery buttons (40 pt tiles, three per row; sizes named Arrow L Size 1–9 / Arrow R Size 1–9). Now Sketched style is a 39 pt menu button with PowerPoint's four 139 × 29 pt choices (line icons with a check mark; Japanese なし, 曲線, フリーハンド, フリーハンド — Mac PowerPoint's Japanese build names Freehand and Scribble alike) that writes `ask:lineSketchStyleProps` (MS-ODRAWXML §2.38, ext `{C807C97D-BFC1-408E-A445-0C87EB9F89A2}`, as Office writes it), disabled for connectors as natively; and the arrow buttons and gallery tiles match. Unverified: PowerPoint replaces the geometry with its generated hand-drawn path when it applies a sketch, while the library keeps the crisp preset and only writes the props, so whether PowerPoint redraws such a shape sketched (rather than crisp until edited) was not checked natively. The preview draws a seeded wobble as an approximation of each style, not PowerPoint's path. Left: Compound type and Dash type are still 112 pt pop-ups (native 39 pt galleries).
- **Position.** Native has no flip controls in the pane; the editor's flip checkboxes are removed (Arrange ▸ Rotate keeps Flip Vertical / Flip Horizontal).

## Views (2026-10-07)

Measured from AX dumps and window screenshots of Mac PowerPoint 16 in a 1512 × 900 pt window (English and Japanese) for every view on the View tab. `native-views-geometry.test.mjs` asserts the matched geometry and structure in both languages: exact on macOS, with documented tolerances for font-dependent values elsewhere.

- **Status bar.** The switcher has exactly PowerPoint's four segments (Normal 37 pt, Slide Sorter, Reading View, Slide Show 36 pt; 147 pt in all). Normal is selected in Normal and Outline View, Slide Sorter in the sorter, and no segment in Notes Page or the three master views. The left text follows the view (Slide N of M, Notes N of M, or the master's name with "Currently in … View."); the Notes and Comments buttons appear only in Normal and Outline View.
- **Slide Sorter.** It opens at 80%, as natively, and Fit returns to 80%. Thumbnails are 250 pt × zoom (200 pt at 80%) in cells 6 pt wider with a 39 pt gutter, centred, giving 206 × 147 pt cells on a 245 pt pitch, six to a row at 1512 pt, with the first at (32, 142). The 24 pt row gap is assumed: the native capture has only one row.
- **Outline View.** Titles start 36 pt in. Body bullets step 11.5 pt per level, and the text sits 6 pt after the bullet. Show Formatting is on by default and draws text at a third of its slide size. The outline text menu follows PowerPoint's order, including Thesaurus… and Translate… (disabled), Show Formatting and Hyperlink… ⌘K.
- **Notes Page.** The portrait 7.5 × 10 in page is fitted with 22 pt around it (96% at 1512 × 900). The slide image and notes body sit where the default notes master places them.
- **Slide Master.** The Slide Master tab replaces Design, Slide Show and Record and is selected on entry. The pane shows each master (a 193 pt image in a 119 pt cell) followed by its layouts (124 pt images indented to x = 90, 75 pt cells on an 87 pt pitch), with the current slide's layout selected. Rename, Themes, Colors, Fonts, Background Styles (Format Background… and Reset Background on the selected layout) and Slide Size use the existing core APIs. A layout's placeholders can be dragged or moved with the arrow keys, which writes the layout's placeholder bounds.
  - Shown disabled with a reason, because the library cannot do them yet: Insert Slide Master, Insert Layout, Delete, Preserve, Master Layout, Insert Placeholder, the Title and Footers checkboxes, Hide Background Graphics, and renaming a master.
  - Placeholder outlines and prompt text use Office's default master geometry and theme fonts. Decorative master and layout shapes are not drawn in this view.
- **Handout Master and Notes Master.** The pages and ribbons follow PowerPoint's default masters: header, date, footer and page number; six handout slide frames; and the notes slide image above five body levels. They are view-only, because the library cannot read or write either master. Orientation, the placeholder checkboxes and the handout's slides-per-page options are shown disabled with the reason.
- **Reading View.** In the standalone editor, Reading View is a full-window mode with PowerPoint's 32 pt title bar ("Slide Show - [name]") and the slide letterboxed below it. Click, the arrow keys and Space move between slides, and hidden slides are skipped. Esc returns to the slide shown last. When a host viewer is attached, Reading View still opens the host's presentation viewer.
- **Shortcuts.** The View menu (see "Menu bar and keyboard shortcuts") drives the views: ⌘1–⌘5 switch to Normal, Slide Sorter, Notes Page, Outline View and Reading View, and View ▸ Master ⌥⌘1–⌥⌘3 open the slide, handout and notes masters. Each view item is checked while its view is shown, as in PowerPoint.
- **Japanese.** View names, status texts, master prompts and ribbon labels use Mac PowerPoint's Japanese wording.

## Show and review tabs (2026-10-07)

Measured like the Home ribbon above (AX dumps of Mac PowerPoint 16 at 1512 and 1200 pt windows, plus the no-effect state and the two Effect Options menus); `show-tabs-geometry.test.mjs` asserts the matched metrics in English and Japanese. All six tabs now use the 72 pt row, PowerPoint's AX button widths (e.g. Transitions Preview 46, Apply To All 38; Slide Show Play from Start 54, Play from Current Slide 73; View Normal 43, Slide Sorter 38), its groups and rules, and ▾ on its menu buttons. Native group captions are absent, as in PowerPoint; the groups carry them as accessible names only.

- **Transitions.** Native: Preview | a framed Transition Styles gallery (18 pt arrow columns, 92 × 56 pt tiles: 10 at 1512 pt, 6 at 1200 pt) with Effect Options beside it | one Timing group (Duration 78 × 24, Sound 102 × 26, On Mouse Click, After + 78 × 24, Apply To All). With no effect, Preview and Effect Options are disabled and Duration stays enabled. Editor before: a 9-tile wrapping gallery, Effect Options opening the transition dialog, timing split over three groups, Duration disabled without an effect. Now matched: the gallery lists all 49 native transitions in order (those the library cannot write — Morph, Reveal, Flash and the other PowerPoint 2010+ effects — disabled with a reason), pages with ‹ ›, keeps the applied tile in view and sizes itself to what fits (measured, so longer Japanese labels take tiles away instead of overflowing). Effect Options is PowerPoint's 58 pt-row menu for the applied effect: Push's four directions, Wipe / Cover / Uncover's eight (Wipe's diagonals are `p:strips`), Split's four, Random Bars / Blinds / Comb / Checkerboard orientations, Shape (Circle, Diamond, Plus; In / Out disabled), Clock (Clockwise, Wedge; Counterclockwise disabled), Fade (Smoothly, Through Black) and Zoom (In, Out, Zoom and Rotate). Tiles write PowerPoint's default option (Push From Bottom, Split Vertical Out, Random Bars / Blinds Vertical, Clock Clockwise). Left: the tile art is the editor's own, the gallery has no expand (▾) button, and Duration shows the speed's duration rather than each effect's own default (PowerPoint: Push 1.00, none 2.00).
- **Animations.** Native: Preview ▾ | Entrance gallery | Emphasis gallery (five 64 pt tiles each) | Exit Effects ▾ | Path Animation ▾ | Effect Options ▾, Animation Pane, Trigger ▾, Animation Painter | Start (102 × 26) and Duration (78 × 24). At 1200 pt the Emphasis gallery collapses into a 55 pt Emphasis Effects ▾ button and Entrance shows six tiles. With no effect on the selection, Effect Options, Trigger, Animation Painter, Start and Duration are disabled. Editor before: one mixed gallery with None, no Emphasis gallery or collapse, Effect Options opening a dialog. Now matched, with PowerPoint's 29 entrance and 24 emphasis effects (the library writes Appear, Fly In, Fade, Zoom and Spin; the rest are disabled with a reason). Effect Options is PowerPoint's Direction (eight, the four diagonals disabled; fly effects only) and Sequence (As One Object, By Paragraph; All at Once disabled) menu. The None tile is gone, as natively; effects are removed from the Animation Pane. Left: Preview has no ▾ menu (AutoPreview), Exit Effects is a menu of the editor's four exit presets rather than PowerPoint's gallery, and Preview stays disabled on a slide without effects.
- **Slide Show.** Native: Play from Start, Play from Current Slide, Presenter View, Custom Show ▾ | Rehearse with Coach | Set Up Slide Show, Hide Slide, Rehearse Timings, Record ▾ and the four options in two columns | Always Use Subtitles, Subtitle Settings ▾. Editor before: Custom Show and Record in groups of their own, Custom Show a plain button. Now matched; Custom Show ▾ opens a menu with Custom Slide Show..., Record ▾ offers From Beginning / From Current Slide. Left: the Custom Show menu does not list the deck's shows (the editor cannot play one).
- **Record, Review, View.** Native widths and grouping applied (Review: a rule between Next and Show Comments inside the Comments group; View: Ruler / Gridlines / Guides on a 19 pt pitch beside Notes). These tabs keep their layout from 1512 down to 1200 pt, as natively. Left: Check Accessibility, Show Comments and Hide Ink are plain buttons (native split buttons with a ▾ menu), and Language opens the editor's language menu rather than PowerPoint's dialog.
- **Japanese.** Gallery and menu items carry Mac PowerPoint's Japanese names. Japanese button labels wrap per character, so they get at least six characters a line at 10 px and the buttons widen; the Animations row collapses its Emphasis gallery whenever the expanded row would not fit.

## Menu bar and keyboard shortcuts (2026-10-07)

Compared with the AX dump of Mac PowerPoint 16.113.3's menu bar (every menu and submenu, English and Japanese UI, with nothing selected, a shape selected and text selected). A web page cannot own the macOS menu bar, so the title bar carries the menus in place of the old Edit and View buttons. `core/menubar-native.ts` is generated from the capture: order, separators, submenus, labels and shortcut glyphs are PowerPoint's in both languages.

- **Menus**: File, Edit, View, Insert, Format, Arrange, Tools, Slide Show, Window and Help. The application menu ("PowerPoint": About, Settings…, Services, Hide, Quit) is left out: those belong to the browser and the operating system, and the editor's language picker stays in the title bar. Lists that name the capture machine's state are not reproduced: Open Recent keeps only More... ⇧⌘O, and Subtitle Settings ▸ Spoken Language / Subtitle Language / Microphone have no items. Window lists this deck with PowerPoint's • mark. U+200B in Japanese labels is stripped.
- **Japanese differences**, from the capture: Pick Up Object Style is ⌥⌘C (⇧⌘C in English); the Window menu has a different order; Slide Show has no Rehearse with Coach. Japanese items carry their English item's id, so one command table serves both languages.
- **States**: Undo and Repeat name the edit ("Undo Move", "Can't Repeat"); Format ▸ Apply To Defaults becomes Apply Object Style when something is selected; View ▸ Enter Full Screen becomes Exit Full Screen. Every item in the capture's enabled-state table (Cut, Copy, Duplicate, Symbol..., Action Settings..., Hyperlink..., Font..., Paragraph..., Columns..., Alignment ▸, More Options..., Pick Up / Apply Object Style, Format Object..., the stacking order, Rotate or Flip ▸ and Align or Distribute ▸ items) is enabled and checked as PowerPoint's with nothing, a shape or text selected (Center ✓ for centred text, Align to Slide ✓). While text is being edited, the menus open without taking focus, so the caret and selection stay. A walk of every menu in the three states against the capture found no item enabled where PowerPoint disables it except View ▸ Enter Full Screen (disabled in the capture, which PowerPoint had in the background); the editor disables the unsupported items listed below and, with nothing selected, Insert ▸ Date and Time... and Slide Number (they insert a field into a selected text box).
- **Commands**: items run the editor's existing commands (the same ones as the ribbon and context menus). Notable mappings: File ▸ Save saves to the project (or downloads outside the development preview) and Save As... downloads; Page Setup... opens Slide Size; Properties... edits the core properties; View ▸ Show Slides toggles the thumbnail pane; Format ▸ More Options... opens Text Options ▸ Text Box; Help ▸ PowerPoint Help opens the command search. Disabled items name the reason as a tooltip. Not available: New from Template, Save as Template, Export, Print, Close, the OneDrive/SharePoint items (Move, Rename, Version History, Share ▸), Restrict Permissions, Passwords, Always Open Read-Only, Paste Special, Clear, Replace Fonts, Find Next/Previous (use the Find dialog), Rename Section (Home ▸ Section), Handout and Notes Masters, Message Bar, Markup, Slides From ▸, SmartArt ▸, 3D Models ▸, Zoom ▸, Action Buttons ▸, Icons, Equation, Object, Cameo, the online media sources, Bullets and Numbering..., Animation Painter (Animations tab), Theme Colors... (Design ▸ Variants), Compress Pictures, Spelling, Thesaurus, Translate, Set Proofing Language (Review ▸ Language), AutoCorrect, Macros, Add-ins, Rehearse with Coach, subtitles, the Window menu's window management, and the system items (AutoFill, Start Dictation, Emoji & Symbols).
- **Keyboard**: every captured shortcut is bound through the same table the menus display (`core/menubar-shortcuts.ts`) and runs its item whenever the item is enabled. Keys match by physical key, so Option's alternate characters do not matter. Control stands in for Command unless PowerPoint gives the Control chord its own item (⌃F Advanced Find, ⌃H Replace, ⌃⌘V Paste Special, ⌃⌥⌘G Guides, ⌃⇧⇥ Show Slides), which keeps the editor's Ctrl+Z, Ctrl+T … working. While text is edited on the slide, ⌘ and ⌃ menu shortcuts act on the selected text (⌘L/⌘E/⌘R alignment, ⌘T, ⌘K, ⌥⌘M, the object-style keys, ⌥⇧⌘V Paste and Match Formatting); undo, the clipboard, Select All and ⌘↩ (which ends editing) stay with the text. In other text fields only document-wide commands run (File, the views, Slide Show, Find, New Slide, Help). A disabled item still swallows its ⌘/⌃ key, so ⌘D on text does not bookmark the page.
- **Changed keys** (to PowerPoint's): ⌘G is Find Next, so Group is only ⌥⌘G and Ungroup ⌥⇧⌘G; ⌘K is Insert ▸ Hyperlink, and the command search moved to ⌘? (Help ▸ PowerPoint Help); ⌥⌘C / ⌥⌘V no longer paint formatting in English — Pick Up / Apply Object Style are ⇧⌘C (⌥⌘C in Japanese) / ⇧⌘V; pasting text without formatting is ⌥⇧⌘V (was ⇧⌘V); ⌘H no longer opens Replace (⌃H does). ⌘0 still fits the slide (the status bar names it).
- **Keys a page cannot have**: the fn/Globe chords (View ▸ Enter Full Screen 🌐F; Window ▸ Fill 🌐⌃F, Center 🌐⌃C, Return to Previous Size 🌐⌃R) and the dictation/emoji keys never reach a web page. The browser or macOS keeps ⌘N, ⇧⌘N (Chrome's incognito window), ⌘W, ⌘Q, ⌘M, ⌘H, ⌘T (Chrome's new tab) and ⌃⇧⇥ (previous tab) in most browsers, and Chrome may keep ⌘1–⌘5 for tab switching outside the embedded preview; File ▸ New Presentation and Insert ▸ New Slide still run from the menus, and their keys work wherever the browser delivers them.
- Tests: `menubar.test.mjs` asserts all 33 menus and submenus (labels, separators, submenu markers, shortcut glyphs) in English and Japanese against the capture, the enabled and checked states of the capture's state table in the three selection states, and that ⌘D, ⌥⌘G, ⌥⇧⌘G, ⌘T, ⌘K, ⇧⌘1, ⇧⌘C (en) / ⌥⌘C (ja), ⌘L, ⌘E, ⌘4, ⌘1, ⇧⌘N, ⌥⌘R and ⌘? run their commands (and ⌘G and the other language's pick-up key do not). `site/test/menubar-shortcuts.test.mjs` covers glyph parsing, key matching, the Control fallback and the en/ja pairing.

## Text, table-cell, picture and slide-background menus (2026-10-07)

Compared with the native AX dumps and screenshots of Mac PowerPoint 16.113.3 (English UI; the trailing Continuity Camera block and Services are macOS's, not PowerPoint's, and are not reproduced). Labels were compared after stripping U+200B and normalising U+2011. Japanese context menus were not captured, so their labels follow PowerPoint's Japanese ribbon wording and are unverified. `context-menus-parity.test.mjs` asserts each menu's items, separators, submenus, shortcut hints and disabled states in both languages.

- **Text being edited** (right-click in the inline editor): Cut, Copy, Paste | Exit Edit Text, Font... ⌘T, Paragraph... ⌥⌘M, Bullets ▸, Numbering ▸ | Thesaurus... ⌃⌥⌘R, Translate... | Format Text Effects..., Format Shape... ⇧⌘1 | Lock | Hyperlink... ⌘K | New Comment ⇧⌘M. This replaces the browser's menu. The menu opens without taking focus, so the caret and selection stay in the text; commands act on the selected range (Font..., Paragraph..., the list galleries and Hyperlink... on it, Cut/Copy/Paste through the async Clipboard API with the same conversion as ⌘X/⌘C/⌘V, Cut only after the clipboard write succeeds), and Escape closes the menu without ending the edit. Format Text Effects... opens the Format pane on Text Options ▸ Text Effects. Cut, Copy and Hyperlink... are disabled with no selection (native Cut/Copy likewise). Bullets ▸ and Numbering ▸ show PowerPoint's galleries — five 90 pt tiles per row: None, filled round, hollow round, filled square, hollow square, star, arrow and check bullets; None, 1. 1) I. A. a) a. i. — followed by Bullets and Numbering.... Left: the gallery bullets are written as Unicode characters (PowerPoint uses Wingdings glyphs), Bullets and Numbering..., Thesaurus... and Translate... are disabled (no dialog / online service), and the menu has no keyboard navigation while the text keeps focus (any other key closes it and reaches the text).
- **Table cells** (a caret in a cell or selected cells): Cut, Copy, Paste | Font... ⌘T, Paragraph... ⌥⌘M, Bullets ▸, Numbering ▸ | Insert ▸ (Insert Columns to the Left/Right, Insert Rows Above/Below), Delete ▸ (Delete Columns, Delete Rows, Delete Table), Select ▸ (Select Table, Select Column, Select Row) | Merge Cells, Split Cells... | Thesaurus... ⌃⌥⌘R, Translate... | Format Text Effects..., Format Shape... ⇧⌘1 | Lock | Hyperlink... ⌘K | New Comment ⇧⌘M. This replaces the editor's earlier Clear cell text / Select all cells / Select table / Size and Position... items (Delete and ⌘A still clear and select). Insert and Delete act on as many rows or columns as are selected, keep the selected cells selected, and deleting every row or column deletes the table; row and column commands end text editing first. Left: Format Text Effects... is disabled for cells (the pane's Text Options cover shape text only), Split Cells... only unmerges merged cells (PowerPoint's dialog splits any cell into N), and row/column commands are disabled while the table has merged cells.
- **Pictures**: Cut, Copy, Paste | Change Picture ▸ (From a File..., From Stock Images..., From Online Sources..., From Brand Images..., From Icons..., From Clipboard...) | Reorder Objects, Reorder Overlapping Objects | Group ▸, Bring to Front ▸, Send to Back ▸, Lock | Hyperlink... ⌘K | Edit Picture, Save as Picture... | View Alt Text..., Crop ⇧C, Auto Crop, Size and Position..., Format Picture... ⇧⌘1 | Action Settings... | New Comment ⇧⌘M. From a File... replaces the picture and Crop opens the crop dialog; the online sources, From Clipboard..., Edit Picture, Save as Picture... and Auto Crop are disabled with the reason. Submenus now show their shortcuts (Group ⌥⌘G, Regroup ⌥⌘J, Ungroup ⌥⇧⌘G, Bring to Front ⇧⌘F, Bring Forward ⌥⇧⌘F, Send to Back ⇧⌘B, Send Backward ⌥⇧⌘B), and the object menu gains the captured Hyperlink... / Format Shape... / New Comment hints.
- **Slide background**: Cut, Copy (disabled), Paste, Paste Special... ⌃⌘V | New Slide ⇧⌘N, Duplicate Slide ⇧⌘D, Delete Slide | Hide Slide | Ruler, Grid and Guides ▸ (Add Vertical Guide, Add Horizontal Guide, Delete | Smart Guides, Guides ⌃⌥⌘G, Gridlines ⌘' | Snap to Grid | Grid Options...), Zoom... | Format Background... | Slide Show ⇧⌘↩ | New Comment ⇧⌘M. Layout ▸, Reset Slide and Select all left this menu (as natively; they remain on the Home tab and ⌘A). Ruler, Hide Slide and the Grid and Guides toggles show their state. Left: Paste Special... is disabled (the browser exposes no clipboard formats) and Delete under Grid and Guides is disabled (right-click a guide to delete it).

## Chart and table contextual tabs (2026-10-06)

- A selected chart shows PowerPoint's Chart Design and Format tabs (Format is the Shape Format tab renamed, as natively). Chart Design carries Add Chart Element, Quick Layout, Change Colors, Chart Styles, Switch Row/Column, Select Data, Edit Data and Change Chart Type; the commands the chart dialog covers open it, and Quick Layout, Change Colors, Chart Styles and Switch Row/Column are disabled with the reason.
- A selected table, or cells being edited, shows Table Design (Table Styles, Shading, Borders) and Layout (Delete/Insert Rows and Columns, Merge Cells, Height, Width, Alignment) instead of Shape Format.

## Context menus (2026-10-06)

- Right-clicking an object offers PowerPoint's menu: Cut, Copy, Paste, Edit Text, Group ▸, Bring to Front ▸, Send to Back ▸, Link…, Edit Alt Text…, Size and Position…, Format Shape… and New Comment. Duplicate, Delete and Copy/Paste formatting left the menu (⌘D, Delete and ⌘⌥C/⌘⌥V still work), as PowerPoint has none of them there.
- Right-clicking a thumbnail offers PowerPoint's menu: Cut, Copy, Paste, New Slide, Duplicate Slide, Delete Slide, Add Section, Layout ▸, Reset Slide, Format Background…, New Comment and Hide Slide (Move slide up/down are gone — dragging reorders thumbnails — and the outline keeps Move Up/Move Down as natively).
- Right-clicking the slide itself now offers PowerPoint's items after Paste: Layout ▸ (checked current layout), Reset Slide, Grid and Guides ▸ (Add Vertical/Horizontal Guide, Grid Options…), Format Background… and New Comment.

## Shape Format ribbon (2026-10-06)

- PowerPoint's contextual Shape Format tab groups: Insert Shapes (Shapes, Edit Shape ▾, Text Box, Merge Shapes ▾), Shape Styles (gallery, Shape Fill ▾, Shape Outline ▾, Shape Effects ▾), WordArt Styles (Quick Styles ▾, Text Fill ▾, Text Outline ▾, Text Effects ▾), Accessibility (Alt Text), Arrange, Size (Height, Width) and Format Pane. The tab now renders these groups in that order instead of the generic fill/outline/effects/size/arrange buttons.
- Edit Shape ▸ Change Shape uses `setShapePreset` with the presets the library writes; Edit Points and Merge Shapes are disabled (no point editor yet; merging needs boolean geometry). WordArt Quick Styles opens the 20-preset gallery (5 × 4 "A" swatches previewed with CSS from the deck theme, Clear WordArt below) that Insert ▸ WordArt also uses; presets are derived from `test/fixtures/native/wordart-*-shape.xml` and a unit test compares each written run against its fixture; Text Effects applies Shadow, Reflection and Glow presets to every run. Arrange is the same menu as Home's.

## Window chrome and Home ribbon (2026-10-05)

Compared against Mac PowerPoint 16 (dark appearance, `review.pptx` built from `packages/dsl/examples/review.tsx`) with window widths 1512, 1200, 1000 and 800 px.

- Home ribbon: PowerPoint shows uncaptioned clusters — Paste with Cut/Copy/Format Painter; New Slide ▾, Layout ▾, Reset, Section ▾; Font (two rows); Paragraph (two rows); Picture, Shapes, Text Box; Arrange, Quick Styles, Shape Fill/Shape Outline; Add-ins and Designer. Implemented all. Paste ▾ offers Paste, Keep Text Only (the system clipboard's plain text, at the text cursor or as a text box) and Paste Special… (disabled: the browser does not expose clipboard formats). Shapes (Home, Insert and Shape Format) opens PowerPoint's gallery — Lines, Rectangles, Basic Shapes, Block Arrows, Equation Shapes, Stars and Banners, limited to presets the library writes — and the chosen shape is drawn by dragging on the slide (Shift keeps it square) or dropped one inch square on a click. Section ▾ adds, renames and removes `p14:sectionLst` sections, and the thumbnail pane shows each section's name above its first slide. Columns ▾ (One/Two/Three, More Columns…) sits on Paragraph's first row; Text Direction ▾, Align Text ▾ and Convert to SmartArt follow the alignment buttons on the second row, as in PowerPoint. Add-ins, Designer and Convert to SmartArt are shown disabled (no Office add-in host, Microsoft 365 design service or SmartArt writer). Shape Fill / Shape Outline show only a ▾ and Font Color is an "A" over a bar of the current color, like PowerPoint, and New Slide / Quick Styles labels wrap onto two lines, so every group stays expanded at 1500 px as natively. Escape that closes a collapsed group or ribbon menu no longer also clears the slide selection.
- Collapse order observed natively and implemented: 1200 px keeps everything but shrinks Slides/Insert to small icons and collapses Drawing; 1000 px collapses Slides, Paragraph, Insert and Drawing; 800 px also collapses Font. Clipboard never collapses. A collapsed group closes after a command is chosen.
- Tabs: native order is Home, Insert, Draw, Design, Transitions, Animations, Slide Show, Record, Review, View. Implemented in that order, including Draw and Record. The active tab is bold with an underline bar and no background, and no tab shows group captions.
- Insert: native New Slide, Table, Pictures, Screenshot, Cameo, Shapes, Icons, 3D Models, SmartArt, Chart, Zoom, Link, Action, Comment, Text Box, Header & Footer, WordArt, Date & Time, Slide Number, Object, Equation, Symbol, Video, Audio. All are shown in that order with the native names. Screenshot captures a shared screen or window (the browser's screen picker), Header & Footer fills the date, slide number and footer slots (a plain box along the bottom edge when neither the layout nor the master has one), WordArt adds PowerPoint's "Your text here" box, Symbol inserts at the text cursor from a character panel (the Mac command opens the system Character Viewer) and Video/Audio embed a file. Cameo, Icons, 3D Models, SmartArt, Zoom, Object and Equation are disabled with the reason as their tip: the library does not write those parts yet.
- Design: native theme gallery, Variants, Colors, Fonts, Background Styles, Layout, Slide Size, Design Suggestions. All are shown in that order. The Themes gallery applies Office's built-in color sets with their font pairs to the deck theme (`setPresentationTheme` + `setPresentationFonts`, keeping masters and layouts) and marks the current one. Colors ▾ and Fonts ▾ list Office's color sets and font pairs with Customize Colors…/Customize Fonts…, Layout ▾ applies a layout to the selected slides, and Slide Size ▾ offers Standard (4:3), Widescreen (16:9) and Page Setup… (scaling with Ensure Fit). Variants (the built-in themes have none) and Design Suggestions (needs the Microsoft 365 service) are disabled. Layout editing itself is on the Slide Master tab.
- Transitions: native Preview, effect gallery, Effect Options, Duration, Sound, On Mouse Click, After and Apply To All. Implemented the gallery (None, Fade, Push, Wipe, Split, Cut, Random Bars, Cover, Uncover — Morph, Reveal and Shape are not written by the library), Effect Options (opens the transition dialog), Preview (plays the effect on the editing canvas), Duration (written as PowerPoint's `p14:dur` inside `mc:AlternateContent` with a speed fallback), Sound ([No Sound], [Stop Previous Sound] and Other Sound… for a WAV file; PowerPoint's built-in sound library is not bundled), On Mouse Click, After and Apply To All. The slide show does not play transition sounds yet.
- Animations: native Preview, entrance/emphasis gallery, Exit Effects, Path Animation, Effect Options, Animation Pane, Trigger, Animation Painter, Start and Duration. All are shown in that order. The gallery holds the library's entrance and emphasis presets (colored stars as natively) and Exit Effects opens the exit presets; Preview plays the slide's effects over the editor and Animation Painter gives the next selected object the selected one's effect. Path Animation and Trigger are disabled: the library does not write motion paths or triggers yet.
- Slide Show: native Play from Start, Play from Current Slide, Presenter View, Custom Show, Rehearse with Coach, Set Up Slide Show, Hide Slide, Rehearse Timings, Record, Keep Slides Updated, Play Narrations, Use Timings, Show Media Controls, Always Use Subtitles and Subtitle Settings. All are shown in that order. Rehearse Timings presents with a slide/total timer and, on exit, asks whether to keep the times as each slide's After timing (and turns on Use Timings); Record plays the show (no narration capture); Use Timings, Play Narrations and Show Media Controls (`p14:showMediaCtrls`) write the show properties. Rehearse with Coach, Keep Slides Updated and the subtitle commands need Microsoft 365 services and are disabled.
- Review: native Spelling, Thesaurus, Check Accessibility, Translate, Language, Mark All as Read, Show Changes, New Comment, Delete, Previous, Next, Show Comments, Always Open Read-Only, Restrict Permission and Hide Ink. All are shown in that order. Check Accessibility opens the issue list, Language sets the selected text's proofing language (`setShapeTextLanguage`), Delete removes the comments on the slide or in the presentation, Previous/Next go to the neighbouring slide with comments and open the pane, Show Comments toggles it and Hide Ink hides every pen stroke. Spelling (the browser checks as you type), Thesaurus, Translate, Mark All as Read, Show Changes, Always Open Read-Only and Restrict Permission are disabled with the reason. The comments pane is docked on the right like PowerPoint's and is non-modal: the ribbon, thumbnails and canvas stay usable, and it follows the selected slide. It keeps the editor's review-then-Apply drafts (one undo step for all slides) rather than PowerPoint's per-comment Post; Apply refuses only if the comments themselves changed meanwhile. Slide Sorter and Notes Page float it over the view.
- View: native Normal, Outline View, Slide Sorter, Notes Page, Reading View, Slide Master, Handout Master, Notes Master, Ruler/Gridlines/Guides, Notes, Zoom, Fit to Window, Macros. All are shown in that order. Notes Page shows the slide above its editable notes on a portrait page, Reading View opens the preview's rendered viewer (thumbnails, zoom, Present) with Normal and Esc returning to the editor, and Slide Master opens a Slide Master tab (Rename, Master Layout, Colors, Fonts, Background Styles, Format/Reset Background, Slide Size, Close Master) acting on the current slide's layout; the canvas still shows the slide rather than the layout. Insert Slide Master, Insert Layout, Delete, Handout Master, Notes Master and Macros are disabled with the reason. Thumbnails and Grid Options are in the View menu.
- Find/Replace are in the Mac Edit menu, not on the Home ribbon; an Edit menu now carries Undo, Redo, Cut, Copy, Paste, Select All, Find and Replace.
- Thumbnail pane: PowerPoint has no add/duplicate/delete/move buttons above the thumbnails; removed (ribbon, context menu, ⌘D, Delete and Option-arrow reordering remain).
- Status bar: native shows "Slide 1 of 3", proofing language, Accessibility; Notes, Comments, Normal/Slide Sorter/Reading View/Slide Show, zoom −/slider/+, percentage and Fit icon on a neutral background. All implemented. The language is the system language and region; Accessibility lists pictures, charts, tables and groups without alternative text and slides without a title (PowerPoint's checker reports more rules). Reading View opens the preview's rendered viewer in the window; its status bar has Normal (Esc also returns), unlike PowerPoint's Reading View, which is a windowed slide show.
- Draw: native Draw, Eraser, Lasso Select, a pen gallery (black pen, red pen, pencil, highlighter), Add, Ink to Text, Ink to Shape, Ink to Math and Draw with Trackpad. Implemented all controls. Strokes are saved as open freeforms named "Ink N" (`setShapeCustomGeometry`) rather than `p:contentPart` InkML, so they stay editable everywhere. The eraser removes whole strokes it crosses, Lasso Select picks shapes wholly inside the loop, Add creates a pen from a color, and Ink to Shape recognizes lines, triangles, rectangles and ellipses. Ink to Text and Ink to Math are disabled: the browser has no handwriting recognizer.
- Record: native Cameo, From Beginning, From Current Slide, Clear Recording, Reset to Cameo and Learn More. Implemented the same buttons; From Beginning/From Current Slide play the show (no narration or camera capture), Cameo, Clear Recording and Reset to Cameo are disabled, and Learn More opens Microsoft's recording help.
- Title bar: native uses neutral window chrome with quick-access Save/Undo/Redo, centered "name — Saved to my Mac" and a Search box. The editor bar now uses neutral chrome and presents the command palette as Search. The Quick Access Toolbar is icons — Save, Undo, Redo and ⋯ — with New, Open and Download under ⋯. An AutoSave switch leads the title bar in the dev preview; turning it off pauses saving to the project until ⌘S or until it is turned back on. The tab row ends with Comments (toggles the comments pane, like the status bar's) and an accent Share ▾ whose Send a Copy downloads the `.pptx`; Share with People and Copy Link are shown disabled because they need OneDrive or SharePoint. In the dev preview an **Agents** toggle (sparkle icon, pressed while open) precedes Comments, where PowerPoint puts task-pane toggles such as Copilot; it opens the Agents task pane docked at the right, closed by default, with a title row and × like the Comments pane and the editor's colors in light and dark. The preview has no header of its own.
- Still different: the editor hides notes by default, while PowerPoint shows "Click to add notes". The editor chrome follows the system's dark appearance like Mac PowerPoint (the slide keeps its own colors). New Slide ▾, Layout ▾ and Design ▸ Layout ▾ are galleries of layout thumbnails (background plus title and content placeholders drawn as prompt boxes, footers omitted) like PowerPoint's.

## Current UI migration

- Application grid/drawing/smart-guide visibility is stored independently of document undo/save state and synchronized on storage events.
- Grid Options stages changes until OK; Cancel also discards a pending Set as Default. New decks inherit committed spacing/snap defaults. Existing non-square/custom grid spacing is retained when only display options change.
- Canvas renders grid dots and stored drawing guides. Guide dragging commits once on release, supports cancellation and undo. Keyboard guide movement/deletion is available.
- Shape movement snaps to independent X/Y grid intervals in slide coordinates, including transformed groups. Smart guide display can be disabled.
- View menu exposes Grid and Guides, Ribbon, zoom actions and browser-supported fullscreen. This is a partial View menu, not full Mac menu parity.
- Guide context menus support adding horizontal/vertical guides, deleting and changing color; all edits participate in document history. Custom grid spacing accepts centimeters and retains EMU precision.
- Remaining here: native snapping priority/resize/drawing behavior, further View modes, native visual comparison and migration of the other operation surfaces. Full operation parity is still incomplete.

## Normal and Slide Sorter views

- View menu and status controls switch between Normal and Slide Sorter; Command-1/2 follows the previously observed Mac menu shortcuts.
- Sorter reuses slide selection, range selection, duplication, deletion, context menu and drag reordering. Arrow keys navigate its rendered rows/columns. Double-click or Return returns to Normal on the selected slide.
- Normal and sorter maintain separate zoom values. Manual normal zoom survives switching; Fit to Window resumes automatic fit. View changes do not modify document history or saved content.
- Browser tests exercise grid dialog cancellation, saved spacing/snapping, guide drag/undo/add/color, two-window display preference propagation and precise custom spacing; sorter tests cover range selection, delete/undo, navigation, independent zoom and shortcuts. Full Mac view appearance and complete menu/ribbon parity are still outstanding.

## Guide snapping and Zoom follow-up

- Visible drawing guides participate in object movement snapping in slide coordinates, including transformed groups, independently of the Smart Guides preference. Grid snapping currently takes precedence; native priority still needs comparison.
- Canvas context menus can add guides after all guides have been deleted. Fine grids render at integer multiples of their actual spacing rather than an unrelated minimum pitch.
- View > Zoom > Zoom... stages presets/custom percentage/Fit until OK. Cancel preserves the view. Menu keyboard events do not delete the underlying selection.
- An earlier native comparison attempt returned `Sky Computer Use native pipe closed before response`. No native document was edited. Existing screenshots/comparison records guide implementation; complete visual and behavioral equality remains unverified.

## Inline notes pane

- Replaced the notes modal with an editable pane below the slide and a Notes status control. The existing notes command opens/focuses this pane. The separator supports pointer and keyboard resizing.
- Typing bursts save without an Apply button; blur, hiding the pane and switching views commit the draft to its original slide. Undo/redo flush pending input before restoring history. Japanese composition is not committed mid-composition.
- Browser coverage checks slide switching, sorter switching, focused autosave, hide/show, resizing, bilingual content and persisted undo/redo. Existing notes/transition coverage was adapted and passes. Native pixel comparison and rich-text notes formatting remain outstanding.

## Zoom controls and View ribbon

- Status percentage opens the Zoom dialog. Added the two-segment zoom slider (10–100–400%), keyboard percentage changes, Home/End and 10-point increment/decrement controls from the earlier implementation. Native button comparison on 2026-10-03 confirmed 120 → 130, custom 123 → 130, and 130 → 120, matching the controller's next/previous multiple-of-ten behavior. Slider keyboard stepping remains unverified.
- Added a View ribbon surface for Normal/Slide Sorter, thumbnail visibility, drawing guides, Grid Options, Zoom and Fit to Window. View state does not change document history. This is an incremental subset, not the complete native View ribbon.
- Ribbon tabs expose tablist/tab/tabpanel semantics with arrow/Home/End navigation. Existing browser locators now address tabs by their semantic role.
- Zoom browser tests cover endpoints, keyboard adjustment, view-specific zoom and dialog cancellation; rich-text tests verify scaled text and selection preservation.

## Thumbnail pane

- Normal view thumbnails resize by dragging the divider or using its arrow/Home/End keys. Escape cancels a drag. The chosen width survives hiding/showing thumbnails and switching views without changing the presentation.
- Removed the decorative canvas dot pattern so grid markings only appear when enabled. Native size limits and exact divider appearance still need direct comparison.

- View menus and submenus stay within the viewport, opening to the left when the right edge has insufficient room. A browser regression test reproduces the prior clipping and covers wide/narrow windows.

## Selection Pane

- Home opens a Selection Pane with reverse stacking order, expandable groups, sibling range/toggle selection, inline name editing, and individual/all-object visibility controls. Changes use document transactions, save, and undo; closing the pane restores properties.
- Group names and visibility now use the group's own nonvisual properties. Hidden objects and hidden group descendants are omitted from preview rendering and canvas hit targets without deleting their content.
- Drag an object name above or below a sibling to change its stacking order, including within expanded groups. Insertion feedback is shown; cross-group drops are rejected. Browser coverage checks saved order and undo for top-level and grouped objects.
- Dragging near the list edges scrolls it; leaving the list, dropping, cancelling or closing the pane stops scrolling. A 45-object browser test checks an offscreen reorder, saved stacking order and undo.
- Multi-object dragging and native pane geometry remain outstanding. Exact appearance, scroll speed and native keyboard comparison remain unverified.

## Native comparison resumed

- Computer Use now connects to the installed Mac PowerPoint. Inspected Home > Arrange and its Align/Rotate submenus, plus the Selection Pane, in a new disposable presentation saved under `/tmp`.
- The native Selection Pane has a single Hide All/Show All eye toggle, a separate Lock All/Unlock All toggle, and Bring Forward/Send Backward footer buttons. The visibility toggle and footer now match those operations, including disabled controls without a selection.
- Home > Arrange now opens a menu for stacking order, grouping, alignment, rotation/flipping and Selection Pane. Object alignment uses geometry rather than the previous paragraph-alignment command. Submenus support keyboard navigation and stay within the viewport.
- Native Lock All writes DrawingML `spLocks` attributes (`noGrp`, `noRot`, `noMove`, `noResize`, `noEditPoints`, `noAdjustHandles`, `noChangeArrowheads`, `noChangeShapeType`). Unlock All removes those attributes. The Selection Pane now supports individual/all-object locks and undo, with geometry restrictions in canvas gestures, keyboard movement, alignment, rotation and grouping.
- Still missing here: native reorder-overlapping view and complete native geometry/styling. Full-operation parity remains incomplete.
- The Arrange menu and properties panel now share their alignment reference. A browser regression first reproduced the mismatch, then verified changing the reference in either surface updates the other, including after reopening the Selection Pane. Both Selection Pane tests pass; editor build and Svelte checks pass.
- Further native comparison: a single selected title placeholder exposes distribution commands but does not acquire changed coordinates after vertical distribution. With two selected placeholders and Align Selected Objects, both distribution commands are disabled. Slide-relative multi-object distribution is still unverified. Native editing paused when the user switched to another document.
- Arrange > Rotate > More Rotation Options now opens the existing properties panel, scrolls to Rotation and focuses its numeric input. Browser coverage verifies entering 37 degrees, saved PPTX rotation, Undo and reopening the Selection Pane. This connects the operation; the existing properties panel still does not reproduce the native Format Shape pane's layout.

## Multiple-object stacking

- Bring to Front, Send to Back, Bring Forward and Send Backward now consume the complete sibling selection, retaining its existing stacking order regardless of click order. The four canonical core APIs accept one shape or an array; each operation rewrites and commits the containing shape tree once. Mixed parent/slide selections fail before mutation.
- Native comparison on the disposable `pptx-native-selection-lock-audit.pptx` confirmed a contiguous pair moves backward together past one unselected sibling. Duplicate/move changes were undone. The temporary slide-alignment reference was restored to Align Selected Objects.
- Regression coverage includes noncontiguous/reversed selections, clamping, duplicate handles, nested groups/extensions, mixed parents/slides, save/load and one-step undo/redo. Browser coverage verifies Selection Pane footer selection, persisted order and undo.
- Core suite: 2,840 passed / 109 skipped; editor unit tests: 57 passed; Selection Pane browser tests: 2 passed. Root TypeScript, Svelte checks and builds pass. Full native operation/appearance parity is still incomplete.

## Mac arrangement shortcuts

- Integrated the native key equivalents recorded in the earlier branch: Option-Command-G / Option-Shift-Command-G, Shift-Command-F/B and Option-Shift-Command-F/B. Physical key codes avoid Option-generated characters. (Superseded: Command-G is now Find Next, as in PowerPoint; see "Menu bar and keyboard shortcuts".)
- Open Arrange menus pass these commands through after dismissal. Text fields, composition and dialogs retain their own input handling. Shift-Command-F no longer opens Find.
- Browser regression first failed on the unhandled backward shortcut. Both Selection Pane browser tests now pass, including multi-object backward/front moves, menu dismissal without Find, grouping and ungrouping through Mac shortcuts. Svelte check reports no errors/warnings and editor build passes.

## Regroup

- Home > Arrange > Regroup and Option-Command-J restore a dissolved group from a selected former member, preserving current geometry and stacking order. Grouping continues to use the canonical `groupShapes` API; only the editor session remembers former membership.
- Undo/redo snapshots include that session history. New/Open clear it; deleted members are pruned so later reused shape IDs cannot restore unrelated objects. Multiple dissolved groups and nested sibling scopes remain separate.
- Reconfirmed Option-Command-G, Option-Shift-Command-G and Option-Command-J in native Mac PowerPoint on the disposable reference. A single selected former member restored both children into a freshly named group; all three temporary operations were undone.
- Editor document tests: 34 passed; full editor unit suite before the final additional nested-group test: 59 passed; Selection Pane browser tests: 2 passed, including menu regroup, undo and the Mac regroup shortcut. Svelte check has no errors/warnings; editor build passes. Native visual parity and the outstanding operation gaps above remain incomplete.

## Slide-relative distribution

- Reproduced native horizontal distribution with three ordinary rectangles in a DSL-authored 16:9 deck. Saved native x coordinates were 1,676,400 / 4,267,200 / 7,772,400 EMU for widths 914,400 / 1,828,800 / 2,743,200. PowerPoint includes both slide margins in equal spacing. One selected rectangle centers (x = 5,638,800 EMU); this resolves the earlier inconclusive placeholder observation.
- Both Arrange surfaces now honor the shared slide/selection reference. Slide distribution moves all selected objects and supports one or two objects; selection distribution keeps outside objects fixed and requires three. Geometry changes retain the existing transformed-group handling.
- Native temporary geometry changes were undone and saved; Align Selected Objects was restored. Regression coverage checks both axes with one, two and three objects and single-step undo/redo.

- Final distribution validation: 36 document tests pass, including literal native saved coordinates and the single-object implicit slide reference. The full editor suite passed 61 tests before this last regression was added; both Selection Pane browser tests pass with saved two-object distribution and undo. Root TypeScript, Svelte check and editor build pass.

## Native overlapping-object view inspection

- Home > Arrange > Reorder Overlapping Objects opens a full-window black view with separated translucent blue slide planes, perspective projection and numbered stacking positions; Cancel and OK are at the lower right. Selecting a plane highlights it. Dragging a plane changes its position; Cancel discards staged ordering. Confirmed on three selected ordinary rectangles, including when their bounds do not overlap. The temporary order was cancelled and the main editor returned with no new undo entry.
- Reconfirmed with two selected rectangles: a single selection disables the command; dragging and Left/Right keys change the staged order; Return commits and one Undo restores the previous order. Temporary native changes were undone. AX button actions did not reliably close this native window, so keyboard confirmation/cancellation was used.
- The web editor now opens a full-window perspective layer preview for selected siblings. It supports pointer dragging, arrow/Home/End keys, Escape/Cancel and Return/OK. Cancel does not change the document; confirmation is one undoable transaction, retaining unselected sibling slots. Nested previews retain ancestor transforms. Exact native animation, noncontiguous-selection semantics and nested-object native comparison still need verification.

## Object locking — native comparison

- Mac Lock All sets group flags `noGrp`, `noUngrp`, `noRot`, `noMove`, `noResize` and separately locks descendants. Locked objects remain selectable, text-editable, deletable and reorderable; Arrange alignment/grouping is disabled. Native temporary edits were undone and the reference has no pending changes.
- Editor locks propagate from groups to descendant geometry, including Regroup members outside the selection. Locked resize handles display a diagonal mark and rotation handles are hidden.
- Public `isShapeLocked` reports an object's own combined movement/resize lock. `setShapeLocked` accepts one shape or a batch, preserving other constraints/extensions and committing once per slide. Group descendants retain independent locks. Pictures, connectors and graphic frames use schema-specific lock elements; those object kinds have schema/round-trip coverage but still need native interaction comparison.
- Targeted core tests: 16 passed, including XML schema validation. Selection Pane browser tests: 2 passed, including locked dragging, keyboard movement, saved locks, text-editor entry, Lock All/Unlock All and undo.

- Final lock validation: 2,845 core tests passed / 109 skipped, 65 editor tests passed, 3 browser tests passed. Root format/lint/TypeScript, Svelte diagnostics and core/editor builds pass. Crop aspect-ratio controls are disabled for locked geometry; general cropping remains available.

- Reorder validation: 2,846 core tests passed / 109 skipped; full editor suite 66 passed plus the final nested-scope regression (41 document tests passed); all three Selection Pane/reorder browser tests passed. Format/lint/TypeScript, Svelte diagnostics and builds passed.

- Numeric property fields now also respect inherited object locks: position, size, rotation, aspect-ratio and flip controls are disabled. A browser regression failed before the fix and passes after it; Svelte diagnostics and the editor build pass. A separate Selection Pane drag test hit a preview-server startup timeout during the combined run; the changed lock scenario passed independently.

## Size and Position controls

- Reconnected to Mac PowerPoint and inspected its Size & Properties pane. Size and Position are separate disclosure sections; dimensions and coordinates use cm. Size orders Height, Width, Rotation, Scale Height, Scale Width and Lock aspect ratio. The editor now follows those sections, labels, units and dimension/rotation input limits.
- Native scaling from 150% to 200% produces twice the original dimension, not three times. Selecting a different object and returning resets that baseline. Percentage controls use the initial selected dimensions, including per-object dimensions in a mixed selection. Temporary native resize edits were undone; native Undo is disabled again.
- Six browser tests passed for geometry, inherited placeholders, mixed selections, fill/line, selection locks and stacking. A subsequent geometry run passed all three tests after adding excessive/zero scale validation. Svelte diagnostics report no errors or warnings; the editor build passes.
- This is a partial migration of the existing properties panel. Native pane tabs, picture-specific scaling, and full Fill/Line/Effects/Text Box layouts remain outstanding. Full UI and operation parity is not complete.

- Added independent horizontal/vertical Position origin selectors. Native comparison confirmed that switching to Center preserves the entered value while moving the object by half the slide dimension; Undo restores the position but retains the origin preference. Browser tests cover both axes, saved coordinates, unchanged displayed values and Undo. All three geometry browser scenarios pass. Nested-object origin semantics still need native comparison.

## Text Box controls

- Native Text Box controls were inspected through the restored Mac connection. Added a collapsible section for vertical alignment, five text directions, autofit, centimeter margins and wrapping, plus a staged Columns dialog.
- Columns accepts 1–16 columns and 0–40.64 cm spacing; margins accept 0–55.88 cm, matching native field limits. Cancel/Escape discard the dialog draft. Keyboard input stays within the modal.
- Browser coverage checks mixed selection, unchanged margins, invalid input, persisted direction/autofit/wrapping, staged columns, Cancel/Enter/Escape, bilingual labels and undo.
- Full format-pane layout and exact text-layout comparison remain outstanding. This is partial operation parity.

## Inherited text layout

- Native comparison with a layout-authored text body confirmed that the placeholder shows Shrink text on overflow, three columns and 0.25 cm spacing without local overrides. The read-only comparison ended with Cancel and no undo entry.
- Effective body properties now resolve autofit and column settings through layout/master placeholders. Rendering, editable text, overflow auditing and the Text Box/Columns controls share these values. Explicit one-column and no-autofit settings override inheritance; clearing local column settings restores inheritance.
- Unchanged Columns confirmation preserves inherited XML. Invalid column arguments are validated before mutation, preserving existing settings on failure. Browser coverage checks displayed inherited settings, unchanged confirmation, a saved one-column override, undo, inline columns and explicit no-autofit.
- Validation: full core run passed 2,859 tests with 109 skips after extending the XML property-test timeout; the remaining DSL assertion was corrected to use the effective getter's two-argument signature and its five-test file passed. All 67 editor unit tests and four Text Box/rich-text browser scenarios pass. Format, lint, TypeScript, Svelte diagnostics and core/preview/editor builds pass.

## Centered text anchors

- Native Top/Middle/Bottom Centered preserve each paragraph's alignment and original wrapping width. Confirmed with a left-aligned `ABC` paragraph and a right-aligned `A` paragraph: they span the original inner frame. Temporary text/alignment changes were undone, leaving native Undo disabled.
- All six anchor choices now map to `anchor` and `anchorCtr`. Effective centering inherits through layout/master; explicit false overrides inherited true. Existing `setShapeTextAnchor` callers preserve centering unless they supply the new `centered` option.
- SVG, HTML and editable text use shared layout-derived translation, including bullets, paragraph alignment and column positions, rather than shrinking the paragraph frame. Text measurement still uses the preview's configured measurer; exact font-metric fidelity remains part of the broader parity work.

- Validation: 2,856 core tests passed in the full run; one byte-for-byte ZIP comparison failed on timestamp metadata and its 11-test file passed on isolated rerun. All 61 focused core tests, 67 editor unit tests, and three browser scenarios passed; the Text Box scenario passed again after the final rotation-coordinate adjustment. Format, lint, TypeScript, Svelte diagnostics and affected builds pass. Clearing an editing host now ignores Chromium's caret-only BR instead of inserting a trailing newline.

## Leaving text editing

- Escape now commits pending text and exits editing, matching the native observation. A browser regression reproduced the lost final input before the fix and passes after it, including reopening the saved text. Both rich-text browser scenarios pass; Svelte diagnostics, format/lint and editor build pass.

## Default column spacing

- Native Mac PowerPoint displays 0 cm for a three-column placeholder with no local or inherited `spcCol`. The comparison dialog was cancelled and Undo remained disabled.
- HTML, SVG, autofit and overflow auditing now use zero spacing when the effective column gap is absent. Regression coverage compares omitted spacing with explicit zero across these paths; all 13 focused tests pass.

## Fill and Line sections

- Native Format Shape shows separate Fill and Line disclosures. Solid Line orders Width, Compound type, Dash type, Cap type and Join type, alongside transparency, sketch and arrow controls. Native test changes to fill/line were undone.
- Split the editor's combined paint section into independent disclosures. Added direct compound, cap and join fields using existing mutation commands and per-selection readback, including mixed group-child selections.
- The native radio selectors, sketch and gradient-line layouts remain outstanding, as does the remaining format pane layout migration. These controls are an incremental migration, not a claim of full native parity.

- Validation: both paint-selection browser scenarios pass, including disclosure state, multi-selection line values, saved PPTX and Undo. Editor unit tests (67), Svelte diagnostics, format, lint and types pass.

## Placeholder opening size

- Native Mac PowerPoint opens sample 01 with its title at the authored size on two lines, although the master supplies bare `normAutofit`. Inheriting that editing policy must not trigger a new shrink estimate during rendering. Explicit saved scales remain effective.
- Added a regression that reproduced an incorrect 0.75 scale before the fix. Both text-preview suites pass (17 tests); title and showcase fidelity return to their CI baselines. The animation copy rollback test now compares all unzipped package parts, excluding only ZIP entry timestamps.

## Arrow galleries

- Native Mac PowerPoint exposes six types at each endpoint and nine sizes. Saved XML confirms Size 2 is small width / medium length and Size 4 is medium width / small length; Size 5 is medium / medium. The size control remains available with No Arrow. All temporary native edits were undone and the disposable deck saved.
- Added visual start/end galleries with independent type and size edits, mixed-selection readback, keyboard navigation and Escape dismissal. Non-line selections and locked selections disable these fields. Endpoint edits preserve the other properties and the opposite endpoint.
- Validation: all six types and nine sizes round-trip through saved PPTX, mixed selections undo atomically, and gallery state survives reload. All 67 editor unit tests pass. The full pre-arrow browser run passed 184/186 scenarios; both failures were stale tests for Escape committing table text and the Text Box disclosure, and both pass after updating those expectations.

## Solid paint transparency

- Native Mac Line transparency at 25% saves `a:alpha val="75000"`. The temporary change was undone and the disposable deck saved; no native edits remain pending.
- Fill and Line now expose percentage sliders and numeric fields, including mixed selections, one-step undo, decimal percentages and saved reloads. Color changes preserve opacity; opacity changes preserve theme references and non-alpha color transforms. Non-solid paint disables these controls.
- Validation: all 2,865 core tests and 67 editor unit tests pass; the browser regression covers mixed values, Undo/Redo, color changes, invalid percentages, keyboard sliders, reload and No Fill. Inherited paint editing and native radio/gradient layouts remain outstanding.

## Format Shape tabs

- Confirmed the native Fill & Line, Effects and Size & Properties tabs through the reconnected Mac UI and screenshots. Shape, connector and group selections now use those three tabs; pictures, charts and tables retain their existing controls pending their own native comparison.
- Tab navigation supports arrow keys, Home/End and localized names. Existing controls stay mounted while hidden so relative scale baselines, disclosure state and geometry origins survive tab changes. More Rotation Options selects Size & Properties before focusing Rotation.
- Browser coverage verifies control visibility, keyboard focus, scale input retention, Undo and Japanese labels. Existing geometry, text-box, paragraph and selection tests now explicitly navigate to the appropriate tab.
- This is a navigation migration. Native effect disclosures and complete Size & Properties contents remain outstanding.
- Native effect audit: Shadow orders presets, color, transparency (0–100%), size (1–200%), blur (0–100 pt), angle (0–359°), distance (0–200 pt). Reflection has presets, transparency/size (0–100%), blur/distance (0–100 pt). Glow has presets, color, size (0–150 pt), transparency (0–100%). Soft Edges has presets and size (0–100 pt). Opening these disclosures made no document changes.

## Format pane visibility

- The Format pane is closed when the editor opens, as in PowerPoint, and opens from Format Shape… / Format Picture… / Format Chart Area… / Format Video… / Size and Position… (object menu, named after the selection as natively), Format Background… (slide and thumbnail menus, Design ▸ Background Styles), and the contextual tabs' Format Pane buttons. Once open it follows the selection: an object shows its format, the slide shows Format Background.
- Still different: the editor's slide options (Skip during presentation, Slide numbers, placeholder resets and the capability list) have no PowerPoint pane; they scroll below the Background settings, above the pinned Apply to All / Reset row, so they stay reachable. Table cells' menu offers Size and Position… and Format Shape… as PowerPoint's does.
- Native Format Shape opens Fill & Line, even when Size & Properties was previously selected. Size and Position opens Size & Properties. Closing the pane expands the slide editing area; these operations leave Undo disabled on the disposable audit document.
- Added a close button, matching context commands, and reopening from More Rotation Options. Closing preserves the scale baseline, returns keyboard focus to the selection, and gives the reclaimed width to the canvas. Opening and closing Selection Pane respects the prior format-pane visibility.
- Native gradient fill audit: preset gradients, Type, Direction, Angle (0–359.9°), selectable stops with add/remove, stop color, position (0–100%), transparency (0–100%), brightness (−100–100%), and Rotate with shape. The temporary gradient change was undone; No fill and disabled Undo confirm restoration. Native gradient controls and Shape Options/Text Options remain outstanding.
- Validation: pane close/reopen, Selection Pane and view-mode browser scenarios all pass (4 tests). The complete tab-migration browser run passed 185/189; all four stale tab-navigation cases pass after correction. All 67 editor tests, Svelte diagnostics and root format/lint/types/build checks pass. Core suite passed 2,867 tests with 109 skips. The CI comment rollback assertion now compares every unzipped part instead of ZIP timestamps; its 11 tests pass.

### Gradient stop rendering

- Shape gradient readers now preserve stop opacity, including composed alpha transforms. Effective shape gradients also expose resolved stop colors with theme/color-map lookup and brightness transforms while retaining the original color tokens.
- Linear and radial SVG previews use these values. Regression tests cover an imported theme stop with luminance and alpha transforms, rendered output, and save/reload preservation.
- Native gradient editing controls and slide-background gradient transforms remain outstanding.

### Gradient stop editing

- Existing single-shape gradients expose Type, Angle, stop selection/add/remove, Color, Position, Transparency, Brightness and Rotate with shape in Fill. Each edit uses document history and project persistence. Locked shapes disable these controls.
- The existing gradient setter accepts stop opacity/brightness and preserves explicit scaling/rotation options. Invalid stop settings fail before replacing the old fill. Native positive brightness writes luminance modulation plus offset; negative brightness uses modulation only.
- Native add-stop comparison places the new stop between the selection and its next neighbor, or before the final 100% stop, using an interpolated RGB color with zero brightness. Temporary edits in the disposable native deck were undone and saved; Undo is disabled.
- Browser coverage verifies edits, add/remove, Undo followed by editing, interpolated opacity, invalid input restoration, radial angle disabling and saved reloads. Core: 2,875 passed / 109 skipped with four workers and a 30-second timeout; Svelte: zero errors/warnings.
- Remaining: picture/texture and slide-background fill radio controls, preset and path-direction galleries, multi-selection editing, preserving arbitrary imported color transforms during edits, and complete native geometry. Preview honors Rotate with shape for linear gradients, including the aspect ratio of rotated shapes (verified against a 4:1 rectangle rotated 45° in Mac PowerPoint). Shape-following gradients remain approximations; path-gradient rotation and gradient scaling still need native parity. Slide-background gradient transforms also remain outstanding. Full parity is incomplete.

- Stop handles now drag directly with local position feedback, one history entry on release, Escape/pointer-cancel rollback and selection/version-change cancellation. Browser coverage verifies the actual SVG stop color/opacity as well as saved values, drag/Undo and cancellation.

### Fill type selection

- No fill, Solid fill and Gradient fill use radio controls in the native order. Selecting No fill hides paint controls; selecting Gradient fill exposes its inline controls. Type changes apply to the selected shapes in one history entry and respect locked selections.
- The editor remembers each shape's solid/gradient settings, including after closing/reopening the pane, and clears them when opening another document. New gradients use the Mac default stop positions 0/74/83/100%, brightness 95/55/55/70%, linear angle 90°, and scaled coordinates. Native switching from gradient to solid and back retained the old gradient; the audit changes were undone and saved (Undo disabled).
- Regression coverage includes restoring a custom gradient, new gradient defaults, Undo/Redo, Japanese/English labels and multiple selected shapes. Picture/texture and slide-background radio controls, preservation of arbitrary imported color transforms remain outstanding.

### Linear gradient direction gallery

- Shapes with matching gradient stops can now edit them together in one history entry. Unedited per-shape angle, scaling and rotation settings are retained; mixed angles are blank and mixed rotation is indeterminate. Native comparison applied 35% transparency to the first stop of two selected shapes while retaining their different `rotWithShape` flags. The temporary change was undone and saved with Undo disabled. For differing stop sets, the pane now displays an empty stop track and blank, disabled stop/angle/direction controls while keeping Type and Rotate with shape available, as observed in Mac PowerPoint. Changing rotation retains each shape’s distinct stops; browser coverage checks saving, undo and reload. Native comparison used three stops on one shape and two on the other, then reverted and saved both temporary edits. Preset and path-direction galleries remain outstanding.

- The Fill pane offers the eight native linear direction choices in the observed order (45°, 90°, 135°, 0°, 180°, 315°, 270°, 225°), with preview swatches and keyboard selection/Escape cancellation.
- Choosing a direction preserves stops and sets scaled coordinates. Native comparison of a 45° choice saved `a:lin ang="2700000" scaled="1"` even when the previous gradient used `scaled="0"`. The disposable change was undone and saved; native Undo is disabled.
- Direction edits share the angle field, project persistence and Undo/Redo. Preset galleries, path-gradient direction choices and exact popup geometry remain outstanding.

### Pattern fill gallery

- Pattern fill exposes the native 48 presets in the observed order, in six columns of rectangular swatches, followed by Foreground and Background controls. New patterns use `pct5`, `accent1` and `bg1`, confirmed from Mac PowerPoint saved XML. The temporary native change was undone and saved; Undo is disabled.
- Preset and color changes apply to multiple selected shapes in one history entry, respect locks, and persist through save/reload. Partial pattern updates retain untouched theme colors and imported transforms. Fill-type switches remember resolved pattern settings; opening or creating a presentation clears the remembered settings.
- Native color menus, exact gallery swatch rendering/spacing, inherited pattern fills and preserving arbitrary transforms across fill-type switches remain outstanding. Full visual and operational parity is incomplete.

### Picture/texture fill

- Selecting Picture or texture fill inserts the native default texture, switches the pane to Format Picture, and adds a Picture category. Controls include Insert, Clipboard, Texture, Transparency, Tile picture as texture and Rotate with shape.
- Tiled mode exposes offsets X/Y (−1,584 to 1,584 pt), scales X/Y (0–100%), Alignment and Mirror type. Saved default XML uses `a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"`.
- Stretch mode replaces tile controls with four offsets (left/right/top/bottom, −100,000% to 100,000%). Both temporary audit changes were undone and saved; Undo is disabled.
- Core opacity reading/writing and preview support image-filled ordinary shapes. Round-trip tests cover transparency, removal and invalid-input preservation.
- `getShapeImageFillLayout` / `setShapeImageFillLayout` read and edit tile/stretch placement without replacing media, crop or effects. Native left offset 25% was verified as `fillRect l="25000"`; the reference was restored and saved. Stretch offsets now affect preview and clip to the shape. Core tests cover both pictures and image-filled shapes, mode changes, invalid-input preservation, round trips and XML schema validity.
- PNG/JPEG tile preview now repeats at the image’s physical size with nine alignment choices, offsets, separate X/Y scaling and alternating X/Y/XY reflections. `getShapeImageIntrinsicSize` reads dimensions, explicit fill DPI and PNG pHYs/JPEG JFIF density (96 DPI fallback). Native default texture metadata is 128 × 128 pixels at approximately 144 DPI. Raster tests cover repetition, reflections and shape clipping.
- Source-cropped image fills now render and round-trip in automated tests, including all tile mirror modes; native visual comparison remains pending.
- Picture fill controls now expose image insertion, transparency, tile/stretch offsets, scaling, alignment, mirroring and rotation, with multi-selection edits in one undo transaction. Visual comparison of this pane remains outstanding. Image insertion also remembers the preceding solid, gradient or pattern fill so switching back restores its settings. Image fills are remembered across type switches, including media, tile/stretch placement, opacity and supported crop values. Mac PowerPoint restored 35% transparency after switching to gradient and back; all four temporary edits were undone and the reference saved. Imported negative/outset crops, explicit DPI overrides and arbitrary image effects are not yet preserved by this cache.
- Texture gallery: Format Shape and Format Background (Picture or texture fill ▸ Texture ▾) and the ribbon’s Shape Fill ▸ Texture show PowerPoint’s 24 textures in native order and names (English and Japanese), five per row with Mac PowerPoint’s 47 pt pitch and 36 pt tiles, keyboard navigable across the five columns. As in Mac PowerPoint, the Format pane shows a Texture row with the label on the left and a placeholder swatch ▾ button right-aligned with the other value controls (it does not show the current texture), the gallery hangs from the button’s right edge, the gallery has no More Textures... (Insert... sits beside it) and no swatch is marked as the current fill; the ribbon submenu keeps More Textures..., which opens the picture file chooser (the native ribbon submenu was not captured). Microsoft’s texture bitmaps are not copied: each tile is generated in the browser from seeded, periodic noise and patterns that evoke the named material, as a seamless 128 × 128 PNG with ~144 DPI pHYs (64 pt per tile, like the native media). Choosing one applies native default tiling (offsets 0, 100%, top left, no mirror, rotate with shape) to every selected shape or selected slide in one undo step, keeping picture transparency; Apply to All carries it to every slide. Native visual comparison of the textures is not meaningful (they are approximations).
- Default texture (superseded by the session rule in "Ruler and texture: native verification"): choosing Picture or texture fill for shapes or slide backgrounds with no picture to restore inserts the generated Papyrus texture with the native tiling, in one undo step, instead of opening a file chooser; Insert... still chooses a file. A shape's pane title becomes Format Picture. Mac PowerPoint captures (`test/fixtures/native/texture-capture.md`) showed Papyrus, the first gallery texture, for both targets, saved as `a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"` with no `rotWithShape` or `dpi`. Browser coverage checks the saved tiled fill, Undo, the Japanese UI and that no file chooser opens. Outstanding: native media density varied between sessions (≈144 DPI and ≈72 DPI) while the generated tiles are always ≈144 DPI; Format Background in PowerPoint also gains Effects and Picture categories, and a shape's Picture category is not offered.
- Rotation-independent image rendering, native comparison of source-cropped tiles and additional image formats/EXIF resolution remain outstanding.

- Picture fill placement modes now remember tile and stretch settings separately for each shape. In Mac PowerPoint, setting Scale X to 60%, disabling Tile picture as texture, and enabling it again restored 60%; the audit changes were undone and saved. Rotation remains a shared setting when switching modes.

- Inserting a replacement picture retains transparency and remembered stretch offsets, switches to stretch mode, and keeps the prior tile settings available. Native verification retained 35% transparency, restored a 25% left offset, and restored Scale X 60% when tiling was enabled again. The three temporary edits were undone and saved. Arbitrary image effects and source crop behavior during replacement still need comparison.

- Clipboard inserts image data using the same settings-preserving transaction as file insertion. Missing-image and denied-access errors leave the document unchanged; browsers without clipboard image reading disable the button. Browser tests inject clipboard responses for image insertion, empty contents and permission denial. Actual OS permission prompts and native clipboard-format conversion remain unverified.

- Browser coverage verifies insertion, replacement, remembered tile/stretch settings, switching back to the preceding fill, Undo/Redo and saved reloads. It reproduced and fixed a radio selection bug when restoring a remembered picture; clicking an already selected picture fill also retains its selection.

### Slide-background fill

- Selecting Slide background fill hides all fill controls and writes `p:sp useBgFill="1"` with no fill choice inside `p:spPr`. It preserves the geometry and line settings. The temporary comparison was undone and saved; Undo is disabled.

- The Fill pane now exposes Slide background fill for ordinary shapes, applies it to multiple selected shapes, hides paint controls, and supports Undo/Redo and saved reloads. Other fill setters clear the background flag.
- Preview paints only the slide background through the shape, covering intervening objects. Raster tests verify fixed slide coordinates through shape rotation/reflection and group scaling, translation, rotation and reflection. Exact native rendering across every background kind remains unverified.

### Path-gradient direction audit

- Mac PowerPoint's Radial direction gallery lists Bottom Right, Bottom Left, Center, Top Right and Top Left. The Bottom Right option saves `fillToRect l="100000" t="100000"` and `tileRect r="-100000" b="-100000"`; Center saves four 50000 focus insets and an empty tile rectangle. All three temporary operations were undone and saved with Undo disabled.
- Fixed the gradient reader's omitted focus insets (zero, not 0.5) and percentage decoding (integer units of 1/100000 or explicit percent strings). Regression cases fail before the fix and check save/reload after it.
- Preview focus coordinates now use edge insets, fixing corner directions that previously rendered at the center. Regression tests cover all four corners and the center, using edge insets for each focus. Subsequent native color verification corrected the stop order (see below).
- Added the five Radial direction choices in native order. Each updates both focus and tile bounds while retaining stops and rotation. Native Bottom Left, Top Right and Top Left saves confirmed the remaining corner insets; all temporary changes were undone and saved with Undo disabled. Browser coverage checks all five saved settings, selected gallery entries, Undo/Redo and reload.
- Rectangular exposes the same five directions in Mac PowerPoint; its Center save uses four 50000 focus insets and an empty tile rectangle. The editor now shares these direction choices and displays rectangular thumbnails. Path keeps a visible, disabled Direction control in the native pane; the editor matches that state. All three temporary native edits were undone and saved with Undo disabled.
- Native `tileRect` is read and written through the gradient API, with round-trip and schema tests; the rectangular preview now uses its expanded bounds for the five native directions. Shape-following slide fills remain radial approximations. Type changes initialize geometry rather than remembering prior directions; the additional native audit below confirms the defaults.

- Rectangular preview now uses four continuous edge gradients instead of a radial approximation. Raster regressions cover the center and four corner directions, expanded tile origins, first-stop focus color, intermediate stops and transparent shared edges. Mac PowerPoint screenshots and saved XML confirmed red at the focus for a red 0% / blue 100% gradient; the previous reversed stop order was wrong and is corrected for radial previews too. Three temporary native edits were undone and saved with Undo disabled. Arbitrary imported/inverted rectangles, shape-following contours and comprehensive path rotation/scaling parity remain outstanding.

- Native type-switch audit: Linear → Radial saves Bottom Right; selecting Top Left and changing to Rectangular resets to Bottom Right; returning to Radial also resets to Bottom Right. Returning to Linear writes `ang="2700000" scaled="1"` and an empty tile rectangle. Linear → Path writes a centered focus and empty tile rectangle. The editor now uses these defaults while retaining stops and rotation. Browser regression reproduces the old failure, then covers this sequence, Undo/Redo, saved reload and the disabled Path direction. All six native changes were undone and saved; Undo is disabled and the initial linear angle is restored to 0°.

### Gradient editing preview transparency

- The gradient stop track now includes stop opacity; previously changing Transparency updated the slide but left this editing preview opaque. Browser regression reproduces the unchanged preview before the fix and verifies partial/full transparency, Undo/Redo and saved reload afterward. Mac PowerPoint confirmed that 60% transparency also changes the track appearance; the temporary edit was undone and saved with Undo disabled.
- The native color-menu audit could open the Color Picker container, but subsequent screen capture failed with ScreenCaptureKit error -3811. No color was applied. Theme/standard/custom color-menu parity remains outstanding.

### Gradient color palette

- Gradient stops now offer ten base theme colors and ten standard colors. Theme selections retain their scheme reference when saved; choosing a base color resets the stop brightness while retaining the other stops. Keyboard selection, Escape cancellation, Undo/Redo and selected-state restoration after reload have browser regression coverage.
- More Colors opens the browser color picker. A native Chrome accessibility check verified the RGB fields and changing red from 192 to 128; the resulting color well read RGB 128/80/77 and the editor recorded the change. This verifies custom color entry, not equivalence with PowerPoint's custom color dialog.
- Tone rows, recent colors, a palette-level eyedropper, slide-specific theme/color-map resolution, integration with other color controls and exact native popup geometry remain outstanding. The current theme swatches use the presentation theme.
- Solid fill and outline now share the palette, retaining theme references and keeping their raw selected color separate from the resolved swatch. Locked selections disable both controls.

- Pattern foreground/background controls now use the same palette. Untransformed theme references remain selected after save/reload and survive switching to solid fill and back. `getShapePatternFill` keeps its resolved RGB default and offers `preserveTheme` for this editing path; transformed colors retain the existing RGB-resolution behavior. Theme/color-map resolution per slide and preservation of arbitrary color transforms while switching fill types remain outstanding.

### Text color palette

- The text formatting toolbar now offers theme and standard colors for selected characters, caret typing, selected objects and table cells. Theme choices are written as scheme references. Browser regressions cover caret formatting in shapes and cells, preservation of surrounding runs, selected-text color changes, Undo/Redo and reload.
- The shared palette resolves theme tokens for its preview swatch and custom-color starting value. Tone rows, recent colors, exact native popup geometry and slide-specific theme/color-map resolution remain outstanding. Highlight colors still use the existing picker.

### Background colors

- Slide and shared-layout background controls use the theme/standard palette and retain scheme references in saved files. Mixed slide backgrounds show no selected swatch. Browser coverage checks multi-slide application in one history step, Undo/Redo, shared-layout save/reload and selected swatches in English/Japanese.
- Mac PowerPoint inspection confirms Design → Background Styles → Format Background, with Solid, Gradient, Picture or texture and Pattern radio buttons; Hide Background Graphics; Color and Transparency; and Apply to All / Reset Background at the bottom. Opening this panel did not change the audit deck (Undo disabled). The editor still exposes slide/layout settings in its existing sections: this native panel structure and the background fill-type controls remain outstanding, as do exact palette geometry and slide-specific color maps.
- Native gradient-background audit: choosing Gradient fill creates two stops at 0% and 100%, Linear type, angle 0°, with the first stop selected. Controls include preset gradients, type, direction, angle, stop add/remove, color, position, transparency and brightness. Rotate with shape is disabled and indeterminate for the background. The audit change was undone and saved; Undo and Reset Background returned to disabled. Core currently reads background gradients/patterns but only exposes solid-color and image background setters, so the panel needs the missing editing capabilities as well as its layout.
- Background gradient readers now share the shape gradient parser, retaining stop opacity/brightness and focus, tile, scaling and rotation metadata for slides, layouts and masters. Regression tests cover imported XML and save/reload. This fixes the data prerequisite for background editing; it does not add the missing background setters or complete theme/brightness rendering.
- `setSlideBackgroundGradientFill` now writes background gradients using the shared DrawingML writer. The command palette has English/Japanese labels and a field-based form; registry dispatch and save/reload are tested. Validation occurs before replacing the attached background. The native Format Background panel, pattern/solid-transparency editing and full background preview fidelity remain outstanding.
- Background gradient stop readers now provide `resolvedColor` for the presentation theme and DrawingML transforms, so previews paint stop brightness instead of the untransformed base color. A regression covers slide, layout and master backgrounds through the SVG renderer. Slide-authored backgrounds also use the slide's effective color map. Layout/master readers still use the presentation's first-master theme and do not incorporate a consuming slide's color-map override; per-master theme resolution remains outstanding.
- Background gradients now expose the shared stop/type/direction/color/position/transparency/brightness controls in Slide options, with multi-slide transactions, undo and persistence. Rotate with shape is disabled and indeterminate for backgrounds, matching the native pane. Native observation on the restored audit deck showed a two-stop red/blue gradient at angle 0 and scaled false; these colors reflect the session's prior gradient, not verified fresh-install defaults. The editor starts new backgrounds with its theme-based stops at angle 0. Dedicated Format Background pane placement, gradient presets, fill-type memory, pattern backgrounds and solid-background transparency remain incomplete.

- Background controls now resolve explicit slide, layout, and master fills in order. Editing an inherited gradient writes a slide override; reset restores inheritance, and solid overrides block ancestor gradients. Unit coverage checks the cascade; browser coverage checks layout preservation, single-slide edits, reset, Undo, and reload. Theme-reference gradient fills and consuming-slide color-map overrides remain unverified.

- Solid background transparency is now editable with a percentage slider and number field. `setSlideBackground` accepts optional opacity; readers preserve alpha (including inherited fills), and the SVG preview paints transparent backgrounds over white. Native confirmation: entering 40% transparency and committing with Return writes `<a:schemeClr val="bg1"><a:alpha val="60000"/></a:schemeClr>`; the temporary change was undone and saved. Setting an AX stepper value alone changed its text without committing the document, so the earlier angle-switch experiment does not establish background fill-memory behavior. That behavior remains unverified.

- Native Apply to All audit: on a one-master, three-slide DSL deck with red and blue slide overrides and a green layout override, applying the first slide's background removes every slide/layout `<p:bg>` and stores red on the master. A separate pattern audit stores `pct5` with `accent1`/`bg1` on the master. Reset Background becomes disabled; Undo restores the overrides. Both temporary native edits were undone and saved with Undo disabled. `applySlideBackgroundToAll` and the Slide options button now perform this master-based operation, preserving DrawingML tokens and remapping background image relationships. Core tests cover solid/pattern/image save/reload, inherited reapplication, relationship reuse, and invalid input atomicity. Multiple-master native behavior remains to be compared.

### Dedicated background pane

The Mac Format Background pane was inspected on the installed desktop app: it
has its own close button and Fill disclosure, with Apply to All and Reset
Background outside the disclosure. The editor now separates background controls
from slide/layout options and the generic capability list. Browser checks cover
closing/reopening, collapsing Fill, selected-slide edits, undo and reload.

The Design entry is currently a direct Format Background button. Native uses
Background Styles → Format Background... alongside a 12-style gallery and Reset
Slide Background; that menu/gallery remains to be matched. Picture/texture fill
remains outstanding. Hide Background Graphics is covered below. This is partial parity,
not completion of the overall UI match.

### Hide Background Graphics

Audited the installed Mac PowerPoint using a copy of `layout-decoration.pptx` at
`/private/tmp/pptx-hide-background-audit.pptx`. Checking Hide Background Graphics
removed both the master's blue footer bar and the layout's TEMPLATE bar and
picture logo. Slide title text and placeholders remained visible. Saving wrote
`showMasterSp="0"` on the slide root; the layout root was unchanged. Undo and Save
restored the reference document and unchecked the control.

The Format Background pane now exposes this checkbox below the fill choices,
including mixed state for selected slides. Core APIs retain all template content,
read both XML false representations, and remove the override when showing graphics.
SVG previews suppress inherited decoration while preserving the background fill
and slide content. Tests cover round-trip, template-part preservation, and editor
selection, undo, and reload. Apply to All was also audited: it writes the same flag to every slide and layout.
After unchecking an individual slide, its layout decoration reappears but its
master decoration stays hidden because the layout still carries the flag. The
implementation and rendering tests cover this distinction. All temporary edits
were undone and saved.

### Background pane footer

The Mac Format Background pane keeps Apply to All and Reset Background against
the bottom edge, separate from the scrollable fill controls. The editor now does
the same. A short-window browser check verifies that scrolling the pattern
settings leaves both actions in place. The native gradient inspection used only
temporary edits, which were undone and saved.

### Background Styles: theme-reference audit (implementation pending)

On the same Mac audit document, choosing Style 6 changed the master background to
`p:bgRef idx="1002"` with `a:schemeClr val="bg2"`; the slide and all layouts still
had no explicit background. Style 4 used index 1001 with `bg1`, and Style 12 used
index 1003 with `bg1`. Both dark styles changed the master's color map to
`bg1=dk1, tx1=lt1, bg2=dk2, tx2=lt2`. Style 6 kept the light mapping. Reset
Background remained disabled for these inherited backgrounds.

The theme's second and third background fill styles were radial gradients using
`phClr` plus tint, shade, and saturation transforms. Consequently, a correct
gallery must resolve the theme background fill list and preserve its color map;
substituting twelve fixed colors or gradients would not match this behavior.
The background reader now resolves theme-referenced gradients through the owning
master theme, substituting `phClr` and applying the effective color map and color
transforms. Tests cover slide, layout, and master references, a separate owning
master theme, rendered radial stops, and unchanged background/theme XML on save.
Gradient readers and writers now retain ordered color transforms when changing
stop positions, direction, brightness, and transparency. The editor clears the
old color transforms when choosing a replacement color. Regression tests cover
saved color transforms and inherited backgrounds edited through the browser.
The gallery, remaining style columns, multi-selection scope, and multiple-master
application behavior still need implementation or auditing.
Each temporary style change was undone before the next comparison, and the final
document was saved with Undo disabled.

Moving Style 12's first gradient stop from 0% to 10% in the Mac pane creates a
slide-local `bgPr/gradFill`. The stop keeps `schemeClr=bg1` with `tint=80000`
and `satMod=300000`; the second keeps `shade=30000` and `satMod=200000`.
Only the first stop position changes. Reset Background becomes enabled, while
the master keeps its original `bgRef`. This confirms that editor stop-position
changes must retain the imported color transforms. Both changes were undone and
the document was saved with Undo disabled.

Style 2 and Style 3 both use master `bgRef idx="1001"` with `schemeClr=bg2`.
Style 2 keeps the light color map (`bg1=lt1, tx1=dk1, bg2=lt2, tx2=dk2`),
whereas Style 3 switches it to the dark map (`bg1=dk1, tx1=lt1, bg2=dk2, tx2=lt2`).
Neither creates a slide-local background, and Reset Background remains disabled.
Together with Style 1 and Style 4, this verifies all four first-row columns;
the remaining gradient-row combinations and application scope are still pending.
Both temporary changes were undone and the audit file saved with Undo disabled.

A two-slide native check selected only the second slide and chose Style 7.
PowerPoint changed their shared master's background to `bgRef idx="1002"`,
`schemeClr=bg2`, with the dark color map. Neither slide nor any layout gained a
background override. This confirms shared-master scope for a single selection;
multiple-master native behavior remains unaudited. The background change and
slide duplication were both undone, then saved with Undo disabled.

The core `setSlideMasterBackgroundStyle` operation now writes these theme
references and swaps the light/dark mappings while retaining accent mappings.
Regression tests cover the seven audited presets, shared-master inheritance,
round-trip preservation of slides/layouts/themes, separate master isolation,
and atomic rejection of invalid indices or missing theme fills.

The Design ribbon now opens a four-column Background Styles gallery with theme
solid/gradient previews, the selected master style, Format Background..., and
Reset Slide Background. Keyboard navigation, Escape, master-wide application,
undo, and save/reload are covered by browser tests. The read-only
`getSlideMasterBackgroundStyles` reader resolves preset colors without mutating
the document. Native connection was rechecked successfully; the menu exposes
12 presets and both footer actions, with Reset disabled on an inherited fill.
Exact native gallery dimensions still need a usable popover screenshot.
Gallery swatches now use the slide renderer, including rectangular gradient
contours and focus. Browser coverage also verifies resetting a slide override,
undoing the reset, and retaining inheritance after reload. Custom pattern/image
theme swatches remain incomplete; multiple-master native selection behavior remains unaudited.

Background picture placement groundwork: the core can now read and edit stretch
and tile placement on direct or inherited picture backgrounds. Editing an
inherited picture remaps image relationships onto the slide without duplicating
media or modifying the master. Placement parsing/writing is shared with shape
picture fills. Round-trip tests cover stretch offsets, tile scaling/alignment/
mirroring, inherited-image edits, and rejection of invalid settings. The native
picture pane was rechecked (Insert, Clipboard, texture, transparency, tile,
offsets, disabled Rotate with shape); the temporary change was undone and saved.
These background picture controls still need to be connected to the editor UI.

The background pane now displays Picture or texture fill for image backgrounds
and provides stretch offsets, tile offsets/scales, alignment, and mirror type.
Rotate with shape is disabled as observed in Mac PowerPoint. Switching modes
remembers their settings within the editing session. A browser test covers an
inherited image edited into a slide override, other-slide preservation, mode
memory, undo and reload. Choosing a new image uses the file picker; native
remembered fill-type switching, texture presets, Clipboard and transparency are
still outstanding for background pictures.

Background pictures now expose the Clipboard source button. File and clipboard
insertion share a guarded asynchronous transaction: unavailable/non-image
clipboard contents and denied access report an error without changing the deck;
selection or document changes during loading cancel the insertion. Browser
coverage exercises success, undo, empty and denied reads, and a deferred read
completed after changing slides. Native clipboard availability/format behavior,
background transparency, textures, and remembered fill-type switching remain.

Background picture transparency now reads inherited alphaModFix opacity and edits
only the selected slides, preserving placement and media. The pane provides a
slider and percentage field; the preview renders opacity. Core coverage checks
inheritance, round-trip preservation, reset, invalid input and SVG output; browser
coverage checks edits, undo and persistence. Background image placement still
needs to be applied by the preview renderer; textures and remembered fill-type
switching also remain incomplete.

Background image placement now reaches the preview: stretch uses the whole source
inside the specified fill rectangle, and tile uses natural image size, scale,
alignment, offsets and alternating reflections. Shape and background tiles share
the renderer. Raster tests verify quadrant placement, inherited tiles and each
mirror mode; save/reload coverage verifies scale and offsets. This resolves the
placement-preview gap noted above. Background source cropping/effects beyond
opacity, unsupported intrinsic image formats, texture presets and remembered
fill-type switching still require work.

Background solid, gradient and pattern settings are now remembered per slide while
switching fill types in an editing session. Native Mac confirmation: entering a
15% gradient stop position with Return, switching to Solid fill and back to
Gradient fill retained 15%; the audit changes were undone and saved. Browser
coverage verifies selected-slide restoration, solid transparency, pattern choice,
gradient stop/type values, inherited gradient transforms, undo/redo and the saved
active fill. Picture background restoration, texture presets and native memory
behavior across document reopening remain unverified or incomplete.

Background picture source rectangles now reach stretch and tile previews, including
inherited images. `getSlideBackgroundImageCrop` reads the effective crop without
changing the source. Raster coverage checks cropped quadrants, negative outsets
and empty source regions before and after inheritance/save/reload. This closes the
source-crop preview gap; fill-type picture restoration and effects beyond opacity
still need implementation.

Picture backgrounds are now remembered per slide when switching fill types during
an editing session. The editor stores an independent background copy and restores
its full XML and dependencies, including source crop, placement, opacity and
imported image effects. `copySlideBackground` supports direct/inherited backgrounds
and copies referenced parts across presentations without flattening them. Core
coverage checks independent snapshots, save/reload, media sharing within a deck
and missing-reference failure; browser coverage checks fill switching, undo/redo
and persistence. Native Mac confirmation: switching a picture background with
40% transparency to Solid fill and back restored 40%; the audit changes were
undone and saved with Undo and Reset Background disabled. In two native audits,
the offset field showed 37% after editing (Return or Tab) but returned to 25% after
switching fill types; the editor currently retains the latest placement, so this
native placement-memory difference remains unresolved. Texture presets and preview rendering of
remaining image effects still require work.

Follow-up placement audit: explicitly confirming Offset left and saving the native
document wrote `<a:fillRect l="37000"/>`. Switching to Solid fill and back still
displayed 25%, ruling out an uncommitted numeric input. Switching to tile and then
Solid/picture restored stretch with 25%, while transparency remained 40%. All
seven temporary operations were undone, Undo and Reset Background were disabled,
and the reference was saved. A separate manually patched copy showed a content
repair warning on save, so observations from that copy are excluded from parity
evidence. Initial-state versus session-memory behavior still needs an independent
valid-fixture comparison before changing the editor's restoration rule.

A valid API-generated image-background fixture also saved the edited 37% offset
without repair, then restored 25% on Solid/picture switching. More importantly,
a second valid document starting at 18% restored 25% without any placement edit.
This rules out restoring each document's initial placement: native session state
crosses documents. Tile mode also reused 60% from an earlier document. Opacity
still retained the latest 40%. All temporary content edits were undone and saved.
The editor's per-slide memory remains a known difference; do not replace it with
an initial-layout rule based on the earlier 25% fixture alone.

Native background Picture > Picture Corrections exposes sharpness, brightness
and contrast. Saving brightness 25% and contrast -20% used an
`a14:imgLayer` source relationship and `a14:brightnessContrast` extension alongside
the rendered image, rather than simply adding `a:lum`. Background correction UI
and editing this original/rendered-image pair remain outstanding.

## Adjacent zoom stops

Native status buttons moved 123% to 130% for Zoom In and to 120% for Zoom Out;
the temporary view was restored to its original 120%. Editor buttons previously
rounded before stepping, causing 123% to skip to 110% when zooming out. They now
advance to the adjacent 10% stop, with existing 10–400% limits. Controller coverage
checks both views, exact stops and limits; browser coverage exercises custom
percentages through the Zoom dialog and status buttons without changing document
history. Full operation and visual parity remains incomplete.

## Native Slide Sorter zoom range

- Mac PowerPoint's Slide Sorter Zoom dialog disables Fit and 400%. Entering 400 displays “The number must be between 20 and 200.” Normal view retains 10–400%.
- The native sorter slider reports position 750 at 80%, confirming a 20–100% lower half. The status-bar Fit button remains available and sets 100%, despite the disabled dialog preset.
- The editor now applies view-specific bounds in its controller, dialog and both slider mappings. Browser coverage checks rejected out-of-range input, disabled presets, slider position, endpoint clamping, Fit and independent Normal zoom without document revisions.
- The temporary reference was restored to sorter 80% and Normal 120%; no content edits remain. Color-palette internals remain unverified because native accessibility exposes an empty Color Picker container and screenshot capture returns a one-pixel image.

## View ribbon display controls

- Native View exposes Gridlines and Guides checkboxes plus a Notes toggle; all three are disabled in Slide Sorter. Notes shows off in the sorter and restores its prior on state in Normal.
- Added the missing Gridlines and Notes controls, sharing the existing display preference and notes pane state. Guides now also disables in Slide Sorter. Browser coverage checks Grid Options synchronization, notes visibility, view switching and unchanged document revision.
- Inspected native Outline View and Ruler without editing content, then restored Normal view and the original disabled ruler. These two editor surfaces remain outstanding; this change does not claim complete View ribbon parity.

- Follow-up native menu inspection confirmed that all five Grid and Guides submenu commands (including Smart Guides, Snap to Grid and Grid Options) disable in Slide Sorter. The editor menu and ribbon Grid Options now match. Keyboard submenu opening skips disabled items. The reference is restored to Normal 120% with Undo disabled.

### Ruler display (2026-09-25)

Native Mac View > Ruler toggles horizontal and vertical rulers; Slide Sorter disables the command. Native screenshots are working again. The reference Mac displays centimeter rulers with zero at the slide center and positive distance labels on both sides; vertical numbers are rotated. The reference Ruler setting was restored off after inspection.

The editor now exposes Ruler in its View ribbon and menu, persists display visibility outside document history, and tracks the slide bounds through zoom, scrolling, and viewport changes. Browser coverage checks actual ruler/slide coordinate alignment, mode switching, reload persistence, and unchanged document revision. This is display coverage only: native indent/tab handles, configurable measurement units, independent vertical-ruler preferences, and exact ruler styling still need implementation/comparison. Microsoft's [ruler documentation](https://support.microsoft.com/en-us/powerpoint/show-or-hide-the-ruler) confirms the view restrictions and separate vertical-ruler preference.

Selecting text in the native reference moves the ruler origin from slide center to the text body. Setting the first-line marker through accessibility to 1 cm serialized `indent="358775"`; setting the hanging marker from zero serialized `marL="358775" indent="-358775"`, preserving the first-line position. These are saved XML observations; pointer dragging did not change the reference, so snapping and drag behavior remain unverified. All temporary changes were undone and saved, and Ruler was restored off with Undo disabled.

The core now provides `setParagraphIndent` for left/right/first-line overrides, including table cells, selective inheritance reset, and atomic range validation. It is registered in the editor command catalog and powers the interactive ruler handles below.

The editor now displays first-line, hanging, and left-indent handles while editing unrotated horizontal text. Ruler origins follow the editable text body's padding, including custom text rectangles. Completed drags apply to the selected paragraphs, preserve selection, and create one undo step; Escape cancels without saving a change. Browser coverage checks the text origin, each handle's exported indentation, unaffected paragraphs, selection, cancellation, and undo. Native pointer snapping, live text reflow during dragging, mixed-indent marker appearance, selected-but-not-editing shapes, rotated/vertical text rulers, and tab stops remain outstanding; this does not establish complete ruler parity.

## Paragraph tabs and current handoff (2026-09-25)

Paragraph tab stops and default spacing now have public APIs, staged Tabs dialog editing, history, SVG/browser preview placement, and direct text editing support. Native Set/Clear/Clear All state and reset behavior were compared and temporary dialogs cancelled. Browser regressions cover left, center, right and decimal positions in preview and direct editing, including typing and Undo, plus default interval changes. Direct editing retains literal tab characters for selection and clipboard offsets.

The user requested stopping at this checkpoint and resuming in a later session. See [NEXT_SESSION.md](NEXT_SESSION.md) for branch/PR instructions, implementation files, verification commands and remaining work. In particular, complex wrapping, multi-run/multi-tab layouts, locale decimal separators, complete ruler tab interactions and all-operation Mac parity are not established by this increment.

## Resumed ruler tab interactions (2026-10-01)

The user resumed implementation toward full Mac parity. The editor now adds left,
center, right and decimal tab stops from the ruler, moves them by dragging and
removes them by dragging away. Escape cancels without a document change; completed
gestures apply only to selected paragraphs through the existing tab API and undo
transaction. Keyboard arrows and Delete also edit the focused tab marker. Browser
coverage checks all four types, unaffected paragraphs, preserved text selection,
move/cancel/delete, Undo and save/reload. Existing indent coverage still passes.

Native clicking in the ruler added a tab and Undo removed it. Native dragging
failed through the accessibility bridge, so its snapping and movement behavior
remain unverified. Ruler visibility was restored off with Undo disabled. Outline
View was inspected and restored to Normal 120%; ordinary text boxes did not appear
in its outline. Mixed tab marker display, live text reflow during dragging,
rotated/vertical text and complete view/menu parity remain outstanding.

## Ruler: mixed selections, live reflow, rotated/vertical text, locale decimals (2026-10-07)

PowerPoint was not driven for this increment. Several of these points were checked
natively afterwards; see "Ruler and texture: native verification" below.

- Mixed selections: the ruler shows the first selected paragraph's indent markers
  and tab stops, the behaviour Office rulers are generally described as having.
  It was not observed in Mac PowerPoint here, nor was whether PowerPoint dims
  mixed markers (the editor does not). Dragging an indent applies the same distance to each selected
  paragraph relative to its own values; moving a tab stop moves it only in
  paragraphs that have it, keeping each paragraph's alignment; adding a stop adds
  it to all. Both choices are assumptions.
- Live reflow: dragging an indent marker or tab stop (including dragging a stop
  off the ruler to remove it) reflows the editing view on a disposable copy of the
  shape. The document is not changed until release, which commits one undo step;
  Escape restores the layout. Native snapping increments while dragging remain
  unverified.
- Rotated and flipped text: markers stay on the axis-aligned horizontal ruler.
  Positions are measured along the text's own lines, with the shape unrotated
  about its centre, so indents and tabs keep their true lengths at any angle.
  This mapping is an assumption, not a native observation.
- Vertical text (`vert`, `eaVert`, `wordArtVert`, `mongolianVert`, `vert270`):
  markers move to the vertical ruler and measure down the lines, or up for
  `vert270`. Arrow keys follow the marker's screen direction. The editing view
  and the HTML preview now apply `marL`/`marR` along the lines (logical
  inline padding), matching the SVG layout and PowerPoint's rotated layout; they
  previously added a left margin. Right-to-left paragraphs keep their previous
  physical sides.
- Locale decimal tabs: decimal stops align on the decimal separator of each run's
  `lang` (Latin-digit form, e.g. `,` for `de-DE`/`fr-FR`, `.` for `en-US`/`ja-JP`;
  runs without `lang` or with an unknown tag use `.`), in both the preview layout
  and editing. Whether PowerPoint uses the run language or the system locale is
  unverified (Word is reported to use the system's decimal symbol).
- Browser coverage: `ruler-transformed-text` (30° rotation origin and live
  reflow, `vert`/`vert270` markers, glyph alignment against the painted preview,
  vertical tab clicks and keys, mixed selection with Japanese labels, live tab
  drag/remove/cancel) and `decimal-tab-locale` (German vs English separators in
  the preview and the editor). Pure geometry, indent rules and tab moves have unit
  tests.
- Still outstanding: complex wrapping, native comparison of all of the above, and
  ruler units other than centimeters.

## Ruler and texture: native verification (2026-10-07)

Mac PowerPoint (English UI, centimetre rulers) was driven through AppleScript
and the accessibility API; selections were made with AppleScript `select`, and
embedded texture media was identified by its MD5. The editor was then matched.

Ruler, verified:

- Mixed paragraphs: a selection across two paragraphs with different indents
  shows the **last** paragraph's markers, whichever paragraph has the larger
  indent (the levels were swapped to check), and also for the whole text range.
  The first paragraph's are not shown and the markers are not blanked. The
  editor showed the first paragraph's; it now shows the last paragraph's
  indents and tab stops. AppleScript cannot select backwards, so "the paragraph
  holding the selection end" fits the capture equally well; the editor uses the
  last paragraph, which does not depend on selection direction.
- Shape selected, not editing text: no markers and no text band; the ruler
  measures the slide with 0 at its centre. The editor already behaved this way;
  `rulers` now covers it.
- Rotated text (30°): the horizontal ruler stays unrotated, with its band and
  markers at the box's unrotated left edge and width. Matches the editor.
- `vert` text: markers are drawn on the vertical ruler from the box's top inset
  edge, along the lines; the horizontal ruler shows no markers. Matches the
  editor. AppleScript's "vertical" text orientation writes `mongolianVert`, not
  `vert`; its ruler behaves the same, as in the editor.
- Decimal tabs: alignment uses the separator of the run's language (`a:rPr/@lang`),
  not the UI or system locale. With the lines `12,50`, `3.25`, `1234` and
  `1.234,5`, en-US aligns the two `.` and right-aligns the others to the tab;
  de-DE (Set Proofing Language) aligns the two `,`. Matches the editor;
  `decimal-tab-locale` now uses the same lines.

Ruler, still assumed: an indent drag over mixed paragraphs moves each paragraph
by the same distance from its own values, and moving a shown tab stop moves it
only in paragraphs that have it (the capture says nothing about drags). The
native white text-area band on the ruler while editing is not drawn.

Texture default, verified: PowerPoint keeps one in-memory "last texture" per
application session. It starts as Papyrus; it changes only when a texture is
chosen in the Format Shape / Format Picture pane (choosing one in Format
Background does not change it); Picture or texture fill uses it for both shapes
and slide backgrounds; it is shared by documents in the same session; it resets
to Papyrus on relaunch and is never stored in the document. Every texture fill
is tiled with `tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"`.

The editor now keeps the same value in module memory (`defaultTexture()` /
`rememberShapeTexture()` in `core/textures.ts`), not in the deck, history or
localStorage, so a page load plays the part of a relaunch. Undo does not reset
it. Assumed: Shape Fill ▸ Texture on the ribbon sets it too, as a texture chosen
for a shape (the ribbon submenu was not captured). `texture-gallery` replays the
native sequence (Papyrus → Canvas for a shape, Denim for a background, Sand for
a shape, Granite for a background) and checks that a reload returns to Papyrus.

## Outline editing (2026-10-01)

Native Mac Outline View shows title and body placeholders, excluding ordinary
text boxes. Its status-bar Normal button remains selected. Enter at the end of a
title inserts a slide using the same layout; Enter in the middle moves the suffix
to the new slide's title. These operations were compared using
`/private/tmp/pptx-outline-audit/reference.pptx`, undone, and the view restored to
Normal 120% with Undo disabled.

The editor now offers Outline View in the View ribbon and menu, with a separately
resizable navigation pane. Placeholder text edits preserve existing run formats,
save through the existing document history, and commit on leaving the view. Enter
in a title splits it into a new slide. Browser tests cover both languages, empty
new titles, formatted title splitting, Undo, switching views with pending input,
and saved reloads. A round-trip test checks empty placeholders and exclusion of
ordinary text boxes.

This is not complete outline parity. Collapse/expand, promote/demote, cross-slide
text selection and drag reordering, rich formatting display, body outline levels,
and the full native outline context menu remain outstanding.

A follow-up comparison with a Title and Content slide confirmed that splitting
the title also moves its following body paragraphs onto the new slide. The editor
now copies the body placeholders with their formatting and relationships and
clears the original body. Bilingual regression tests reproduce the prior failure
and verify body movement, run formatting, Undo and saved reloads. The native
comparison used `/private/tmp/pptx-outline-audit/body.pptx`; the temporary text
edit was undone.

## Outline collapse (2026-10-01)

Double-clicking a native outline slide icon collapses its body; Undo restores it.
PowerPoint persists the state in `outlineViewPr/sldLst` with slide relationships
in `viewProps.xml.rels`. The temporary collapse was undone and saved.

The editor now matches this gesture, persists each slide's state in that same
OOXML representation, and supports Undo. Bilingual browser tests check collapse,
Undo, saved reload and expansion; schema-validated round-trip tests check slide
identity and preservation of grid settings. Removing a slide also removes its
outline entry, preventing dangling references. This does not implement the full
outline context menu, hierarchy editing or cross-slide text selection.

The native outline context menu exposes Collapse and Expand submenus, each with
selected-slide and all-slide commands. The editor now provides these four
commands on outline slide icons, with keyboard submenu navigation and one-step
batch Undo. A shared-library batch mutation validates slide ownership before
changing the package and updates the shared view part in one batch. Right-clicking an icon
after editing text now selects its slide before opening slide commands. The
reference was restored to Normal 120% with Undo disabled. Other outline menu
commands and right-click text editing still need comparison and implementation.

## Outline body Tab (2026-10-01)

Native Tab at the beginning or middle of a body paragraph changes its paragraph
level instead of inserting a tab character. Saving the reference after Tab at
“Second” produced `a:pPr lvl="1"` on the second paragraph only. All temporary
changes were undone; Normal view, 120% zoom and disabled Undo were restored.

Outline body Tab now updates the selected paragraphs through `setParagraphLevel`,
retaining text and run formatting, and supports saved persistence and Undo.
Selection uses actual paragraph elements, so soft line breaks do not become
separate paragraphs. Shift+Tab, title demotion, hierarchy display and the full
Promote/Demote menu remain outstanding; this is not complete outline parity.

## Outline view keyboard switching and native Promote audit

- Native Mac PowerPoint on the reference deck confirms Command-4 opens Outline View and Command-1 returns to Normal. The editor now handles Command-4 and displays it in View > Outline View.
- Browser coverage switches from outline text editing to Normal before the autosave delay, verifies saved text and formatting, and returns with Command-4 in English and Japanese.
- Native Promote on the top-level body paragraph `Second point` creates a second slide titled `Second point`; the following `Third point` body moves to that slide, while `First point` stays on the original slide. This behavior is not implemented yet. A paragraph-level clamp alone would be incorrect. The native audit was undone, saved and returned to Normal at 120% with Undo disabled.

## Outline title formatting and promotion comparison

- Title Enter now copies the existing paragraph XML and remaps its relationships. Hyperlinks, paragraph properties and untouched fields survive the split instead of being rebuilt from plain text and character formatting.
- The existing `setShapeParagraphs` writer accepts `{ source, range? }` for a UTF-16 range (exclusive end) or all text. Destination body settings remain unchanged. Unit tests cover cross-slide and cross-presentation links, paragraph levels/tabs, partial runs, fields, self-copy, invalid ranges and save/reload. English/Japanese browser regressions verify split-title hyperlinks after saving.
- Native comparison: promoting two selected level-zero body paragraphs creates two separate slide titles, retaining the preceding body on the original slide. This is observed behavior, not yet implemented. The disposable native document was restored to one slide, Normal 120%, Undo disabled, then saved.

## Outline slide movement

- Slide-icon context menus now expose Move Up / Move Down, sharing the existing multi-slide reordering command and history. Boundary commands disable at the first/last possible position.
- English/Japanese browser coverage checks single and multiple slide moves, context-menu selection retention, Undo/Redo, saved PPTX order and reload. Both cases pass against isolated source and ordinary builds.
- Native menu inspection confirmed Move Up is available for the second of two slides and Move Down is disabled there. Exact text-selection/movement behavior remains unverified; this implementation covers slide-icon selections. All temporary native operations were undone and the reference was saved at one slide, Normal 120%, Undo disabled.

## Outline range updates and clipboard review

- The canonical paragraph-level reader/writer now accepts UTF-16 ranges. Relative level changes clamp each selected paragraph independently and commit the text body once; the outline Tab action uses this batch path instead of repeatedly scanning and rewriting the whole body. Tests cover exclusive ends, carets, empty final paragraphs, surrogate boundaries, rich-text retention and table-cell persistence.
- Clipboard metadata rejects unsupported underline/strike tokens, out-of-range sizes/spacing and invalid colors before text edits begin. Valid scheme-prefixed colors survive. HTML paste ignores sub-point font sizes that the OOXML writer cannot represent. Both defects were reproduced before fixing.
- Verification so far: full core suite 3,061 passed / 109 skipped before the additional table-cell test; all 5 paragraph-range tests then passed. Editor unit suite 83 passed. The standalone HTML size browser regression passed. Subsequent ordinary-build integration results are recorded below.
- Native audit of Command-[ through the accessibility bridge did not change the nested paragraph; this is not evidence that the shortcut is absent in PowerPoint. Reference body.pptx restored to Normal 120%, one slide, Undo disabled, and saved.

## Verification checkpoint and outline selection lifecycle

- A frozen ordinary build of `10e68754` completed all 212 Chromium integration tests with zero failures (38 minutes). This run predates the paragraph-range, slide-movement and clipboard changes described above.
- CI on the same commit passed static checks, Node 22/26, preview fidelity and OOXML validation; its Node 24 browser run found an intermittent outline selection error (211 passed, 1 failed). A deterministic bilingual regression reproduces that error by delivering a queued selection event after switching views.
- Outline selection handlers now read the event's textarea instead of a binding that Svelte has cleared during unmount. Title splitting captures the owner document before replacing its input. Both lifecycle regressions pass against the isolated source build.
- The latest core/editor builds and DSL type check pass; Svelte reports zero errors and warnings. All 11 focused ordinary-build browser tests pass: bilingual outline editing, movement and selection lifecycle, formatted text clipboard, HTML clipboard and HTML font-size boundaries.
- Native Shift+Tab delivered through the bridge again behaved as Tab and raised the level. This shortcut remains unverified. Both temporary edits were undone and saved; the reference is back in Normal view, 120%, one slide, Undo disabled.

## Outline body promotion and batch text preservation

- Body Promote now raises nested paragraphs one level and splits selected root paragraphs into separate slide titles, retaining the preceding body on the original slide and following paragraphs on the generated slides. Shift+Tab and the body context menu use the same operation; Tab and Demote lower selected paragraph levels.
- The generated slides use the current layout. Layouts without both a title and a body slot report an error before insertion. Title Demote, cross-slide text selection, formatted/indented outline rendering and the remaining native body context-menu commands are still incomplete.
- Batch overloads on `addSlideAt` and `setShapeParagraphs` allocate shared IDs and copy paragraph ranges without rescanning the source for every destination. `getShapeParagraphElements(shape)` reads all paragraphs in one pass. Tests cover source-as-destination, same-slide and cross-slide targets, links, save/reload and invalid ranges/layouts before text or slide mutation.
- Clipboard import rejects negative/out-of-range kerning and baseline percentages outside the signed OOXML integer range.

### Mixed-level promotion verification

Mac PowerPoint promotes a selected root paragraph into a new title and moves a selected nested paragraph to level zero in the new body. Verified with `Second point` / `Third point` in the reference deck; saved XML agrees with the implementation. Restored the reference to one slide, original text and levels, Normal view at 120%, Undo disabled. English/Japanese browser tests also pass for menu-based multiple-root promotion, Undo/Redo and saved text.

## Outline clipboard and title demotion

- Outline copy/cut includes pending text and run formatting; paste accepts the editor clipboard and external HTML. Cut/paste participate in undo and persistence.
- Title Demote (including Tab) appends the title and following body to the preceding slide body and removes the source slide. The first title is unchanged. Extra non-outline objects and missing destination body placeholders are rejected before mutation, pending native comparison.
- Native comparison confirmed title demotion preserves following body paragraph levels; the reference `/private/tmp/pptx-outline-audit/body.pptx` was restored to one slide, Normal 120%, with Undo disabled and saved.
- The canonical `setShapeParagraphs` API accepts `{ sources }` to concatenate paragraphs, retaining fields, formatting and slide relationships.
- Validation: core 3,068 passed / 109 skipped; editor 88 passed; six English/Japanese browser cases passed (outline editing, clipboard and selection lifecycle). Root types, Svelte and core/editor builds passed. This is targeted validation after the earlier frozen 212-case suite, not a new full-browser pass.

## Outline context clipboard and slide dragging

- The outline text context menu now offers Cut, Copy and Paste. Rich HTML formatting survives menu clipboard operations; Cut changes the document only after a successful clipboard write. Delayed clipboard responses cannot edit a different input, selection or document revision.
- Slide icons now drag single, contiguous or disjoint selections, show an insertion marker and retain selected slide order. Drops from outside the pane and stale document revisions are ignored. Undo and saved/reloaded order are covered.
- Four English/Japanese browser cases passed against the ordinary editor build (clipboard and slide movement). Mac multi-selection uses Command-click; Control-click invokes the system context menu. Exact native drag scrolling and placement remain unverified.

## Outline body paragraph movement

- Native Mac PowerPoint Move Up/Down moves the selected body paragraphs without automatically moving subsequent nested paragraphs. Move Up is disabled for the first body paragraph. The disposable reference was restored to its original text/levels, Normal view at 120%, Undo disabled, and saved.
- Body text context menus now provide Move Up/Down with boundary disabling, whole-paragraph selection after movement and one-step Undo. The existing `setShapeParagraphs` writer can concatenate ordered ranges into one target, retaining paragraph XML, fields and relationships.
- Core tests cover self-reordering, formatting, links, paragraph levels, empty trailing paragraphs and rejected ranges without mutation. Cross-slide movement and title-text movement remain unverified.
- Validation: 3,069 core tests passed / 109 skipped; 89 editor unit tests and both English/Japanese outline browser cases passed. Format, lint, root/DSL types and Svelte checks passed; core/editor builds passed. The full-browser checkpoint remains the older frozen build.

## Slide ordering relationship fixes

- Slide order reading, sorting and individual moves now resolve absolute and normalized relative relationship targets. Regression tests first reproduced a load error, dropped slides on sort, and a failed individual move; save/reload retains both slides after the fixes.
- Editor multi-slide movement applies the final order in one canonical sort instead of repeatedly rewriting the deck. Bilingual browser tests pass for single, contiguous and disjoint selections, Undo and saved order.
- Validation: full root suite 3,070 passed / 109 skipped before the individual-move extension; the extended ordering/deck suite passes 19 tests. Editor unit suite 89 passed. Format, lint, root types and core/editor builds pass.
- Native title-text Move Up shifts the preceding slide's last body paragraph below the selected title. Move Down shifts the current slide's first body paragraph above the title onto the preceding slide. This is different from moving a slide icon. The reference was restored to one slide, Normal 120%, Undo disabled and saved. Title-text boundary movement is not implemented yet.

## Outline title boundary movement

- Title-text Move Up/Down now transfers the adjacent body paragraph across the title boundary, following the native comparison above. The first title cannot move; operations require body placeholders on both slides and a nonempty source. The title remains selected and one Undo restores both bodies.
- Uses the existing paragraph concatenation and distribution APIs, retaining paragraph XML and relationship targets. Tests cover both directions, paragraph levels, links and save/reload. Complex layouts and cross-slide text selection remain in the outstanding list.
- Validation: full root suite 3,071 passed / 109 skipped; editor unit suite 91 passed; English/Japanese outline browser tests both passed, including empty preceding body, both movement directions, title selection, formatting persistence and Undo. Format, lint, root/DSL types and Svelte pass; core/editor builds pass.

## Outline demotion confirmation and missing placeholders

- Mac comparison confirmed that demoting a title with additional objects shows a Yes/No warning before deleting the slide and its objects. The editor now presents that warning; No preserves the document and Yes merges outline text with undo support.
- A deleted body placeholder is restored from the existing layout. On Title Only, demotion preserves the layout and creates a body placeholder that inherits directly from the master, using an unmatched index as observed in native saved XML.
- Core and editor round-trip tests cover both missing-body cases. Browser coverage exercises confirmation, cancellation and Undo in English/Japanese for both existing-body and Title Only slides.
- The reference body.pptx was restored to one slide, Normal 120%, saved with Undo disabled after the comparison.

## Outline paragraph hierarchy

- Outline editing now renders paragraph levels, bullets and numbering through the shared text input. Native comparison confirms the unformatted view uses fixed-size text, 24px body rows and approximately 10px indentation steps. Title text is bold in this view.
- English/Japanese outline browser coverage passes 12 cases on an isolated build, including title splitting, movement, clipboard, demotion confirmation and paragraph geometry. Additional multiline paste coverage verifies empty paragraphs, literal markup, Japanese text, selection offsets, saving and Undo.
- Shared canvas text-input regressions pass 12 cases; the standalone HTML parser test also passes after correcting the isolated test environment's source path. Native Show Formatting visibly changes the outline text sizes; its switch, formatted text rendering and cross-slide selection remain outstanding.
- Native comparison used body.pptx; Show Formatting was toggled twice, Normal 120% restored and saved, with Undo disabled. No pending native document edits remain.
- Guide read/write ID lookups now use a Set/Map, preserving first-match metadata without quadratic scans. The six guide API tests pass.

- Outline text context menus now share the slide-icon Collapse/Expand submenus, including all-slide actions. Bilingual browser coverage checks text-origin actions, keyboard submenu navigation, Undo and persistence.
- Latest root unit suite: 3,072 passed / 109 skipped. Svelte validation: zero errors and warnings.

## Ordinary-build verification and native audit follow-up

- The frozen ordinary build at `4b61cb1e` completed all **223 Chromium browser tests**, with no failures or skips. GitHub CI also passed static checks, Node 22/24/26, preview fidelity and OOXML validation for that commit. These results predate the paragraph-rendering changes above.
- A follow-up attempt could not reliably activate Show Formatting through the accessibility bridge. The similarly named `showFormatting` OOXML attribute belongs to Slide Sorter, so it is not evidence that the outline preference is persisted. Outline formatting-display persistence remains unverified.
- Native Command-Shift-N in Outline View added a second empty Title and Content slide using the same layout as the selected Title and Content slide. Saved XML confirmed the two placeholders and shared layout. The addition was undone; Normal view, 120%, one slide, disabled Undo and saved state were restored.

- The new ordinary build passes all 13 targeted outline/clipboard browser tests (English/Japanese), root format/lint/TypeScript, DSL type checking and core/editor builds. Root unit results remain 3,072 passed / 109 skipped.

## Outline text slide actions

- Text context menus now expose New Slide, Duplicate Slide and Delete Slide through the existing undoable slide commands. Outline New Slide inherits the selected slide layout, matching the observed native Title and Content case; special handling after a Title Slide still needs native comparison.
- English/Japanese browser tests verify insertion with empty placeholders, whole-slide duplication including extra objects, deletion of the last slide, saved content and one-step Undo for all three operations (2/2 passed).

## Outline formatting display (2026-10-01)

- Add the English/Japanese Show Formatting checkbox to outline text and slide menus. It uses session view state, not a presentation transaction. Switching keeps literal text offsets and the original run formatting for editing, clipboard and saving.
- Native comparison: temporarily bolding “First” in the reference body's first paragraph appeared on the slide. Toggling Show Formatting changed the outline between compact formatted text and uniform plain text; turning it back on and Undo restored the original. Normal view, 120%, saved document and disabled Undo were restored.
- The formatted renderer resolves inherited run styles through the existing public reader and escaped HTML exporter. Foreground follows the UI, so dark appearance stays readable. Quarter-scale point sizes are an approximation of the observed compact view; exact font metrics, size limits, default state and persistence across native sessions remain to be compared. Do not call this pixel parity.
- Browser coverage in English/Japanese checks checked state, bold/italic display, unchanged save revision on toggles, editing/saving with original point sizes, and one-step Undo. Existing outline slide actions remain in the same test.
- Validation: final menu/formatting tests 2/2; existing clipboard, selection lifecycle and outline editing tests 6/6. The first expanded test used Control+End, which did not move the caret on this Mac; it now sets the insertion range explicitly. Site unit tests 93/93; format/lint, Svelte (0 errors/0 warnings) and editor build pass. No core behavior changed in this increment.

# アウトライン入力中の継承書式（2026-10-01 追記）

複数マスターを持つ資料で、Show Formatting 有効時の入力プレビューが別マスターの文字サイズを参照する問題を修正。元の shape を段落・run の継承元として渡す。ブラウザーで入力直後の 20px → 10.6667px の変化を修正前に再現し、修正後は日英 2 ケースで表示保持・保存・Undo を確認した。これは継承元の修正であり、ネイティブとの表示倍率やメトリクスの完全一致を示すものではない。

## Parallel implementation: insertion, outline navigation and batch ordering

- Generic New Slide now follows the current layout, with the observed Title Slide
  exception selecting a content layout from the same master. The context menus in
  Normal and Outline views and Command/Control-Shift-N use this policy. Explicit
  layout choice and the blank-slide API retain their existing meanings.
- The Selection Pane moves selected siblings together, retaining relative stacking
  order and selection. Group children stay within their group, and a whole batch
  participates in a single undo step. Native batch-drag gesture comparison remains
  outstanding.
- Plain Up/Down at the start/end of an outline textbox moves into the neighboring
  title or body after committing pending input. Modifier-key text navigation remains
  available. This does not implement cross-slide text-range selection.
- The native comparison connection failed again with a 0×0 capture error while
  opening Arrange. No reference-document content was changed in this attempt.
- Validation: 9/9 targeted browser tests pass across new-slide-layout, outline-view,
  outline-slide-menu and selection-pane, including save/Undo and English/Japanese
  outline cases. Core tests: 3,073 passed / 109 skipped; site tests: 96/96.
  Format, lint, core/Svelte/DSL type checks and core/editor builds pass. This is not
  a fresh full-browser run or a complete native parity audit.

## User preview feedback: text editing and workspace chrome

- The user reproduced a white rectangle behind text while editing the dark review
  slide. The editable overlay used a hard-coded white background. It now stays
  transparent and hides only the original text glyphs during editing, retaining
  slide artwork and table cell fills. The regression test first failed on the
  white background, then passed for dark-slide editing, text color, save/reload,
  and preserving the neighboring cell while editing a filled table cell.
- Font selection now offers a searchable common-family list while retaining direct
  input. A browser test verifies partial-run Arial selection, persisted mixed fonts,
  and Undo. This does not enumerate locally installed fonts.
- Embedded preview removes the duplicate command/help bar, combines save status
  with document commands, and reduces the outer header to 40px. English/Japanese
  layouts pass at 1500px and 900px widths, including Preview toolbar restoration.
- All three focused browser tests pass. These usability corrections do not
  establish complete Mac PowerPoint UI parity.

## Outline text ranges and formatting boundaries

- Shared outline ranges now span title/body fields and slides. Copy, cut, paste,
  ordinary typing, IME replacement, keyboard extension and Ribbon character/paragraph
  formatting consume the shared range. Formatting is one undo transaction; queued
  caret restoration does not steal focus from a subsequently selected control.
- Browser regressions cover the integrated paths in `outline-cross-selection`,
  `outline-ribbon-selection`, `outline-selection-lifecycle`, and
  `outline-reorder-add-edit`. The focused integration run passed 15 cases on
  c4335475. This supersedes earlier statements that cross-slide ranges are absent;
  exact native selection gestures and all input methods still require comparison.
- SVG line wrapping now finds Japanese punctuation boundaries across formatting
  runs and maps the segments back to their original styles. Regression coverage
  includes narrow boxes, explicit breaks, spaces, different character sizes and
  Latin wrapping. The root suite passes 3,093 tests (109 skipped) on 4ce7310d.
- The preset-shape Latin-word fidelity regression was subsequently fixed in
  76810d7a: Mac PowerPoint also splits words wider than the entire text area when
  `latinLnBrk` is omitted/false. Normal word-boundary behavior is retained. Tests
  cover both single and styled runs, and CI run 37032402762 passed Preview
  fidelity without lowering the baseline. This resolves that regression, not
  overall visual parity.

## Video poster frame editing (2026-10-02)

- Video Format exposes Poster Frame > Current Frame and Image from File..., using the existing picture-image API without replacing the media relationship. Poster images remain visible until inline playback or seeking begins.
- Native Mac inspection of `/private/tmp/pptx-poster-audit/deck.pptx` confirmed these menu labels and Current Frame disabled before playback, enabled after playback. Reset now decodes the first video frame and replaces only the poster image. Native before/after files in `/private/tmp/pptx-poster-audit/` confirm unchanged video bytes and playback timing, with the replacement PNG matching the first decoded frame (52.9 dB PSNR; decoding color rounding). Trimmed-video Reset semantics and the remainder of Video Format layout/effects still need comparison.
- Browser coverage checks saved poster bytes, preserved video bytes and playback settings, undo, reload, idle visibility and localized keyboard dismissal. This does not establish complete Video Format parity.

## Video Format ribbon (2026-10-02)

- The contextual tab now renders video controls instead of generic Shape Format groups: preview transport, poster frame, border/effects, alternative text, arrangement, inline dimensions, and the format pane.
- Native Mac inspection confirmed the group order and inline Height / Width / Lock Aspect Ratio controls. Color, the video style gallery, Crop, and individually exposed arrangement buttons remain incomplete. Corrections now exposes the native 5×5 brightness/contrast preset layout and a link to the dedicated Video tab with brightness/contrast sliders and numeric fields; recolor/crop pane controls and pixel-level correction rendering still require comparison. Border/effect/alternative-text dialogs still reuse the existing editor commands; this is not proof of native dialog parity.
- Aspect-ratio locking now reads and persists OOXML `noChangeAspect` in both the video ribbon and size pane. Native comparison confirms unchecking removes the attribute and checking writes `1`, without adding an Undo entry. Current-snapshot persistence matches this observed operation; behavior across older Undo/Redo snapshots and Shift-modified dragging still requires native comparison. Native video drag comparison confirmed that saved aspect locking constrains corner handles, while edge handles stretch only their corresponding axis. The canvas now uses that saved constraint for corner drags; unlocked corners remain freely resizable.
- A trimmed-video Reset experiment produced a generic play-icon poster, so it did not resolve whether Reset uses the source start or trimmed start. Both temporary changes were undone and saved; the fixture no longer contains a trim element.

### Video corrections gallery (2026-10-02)

- Native Mac audit: brightness columns and contrast rows each use -40, -20, 0, +20, +40 percent. The gallery shows generic landscape thumbnails. +20/+20 saves `a:lum bright="20000" contrast="20000"`; one Undo restores both. The reference deck was restored and saved.
- The editor applies both values in one transaction, and the inline video uses the same transfer function as the poster renderer. This does not establish pixel-level agreement with native PowerPoint.
- Movie Correction Options opens the dedicated Format Video > Video tab. Brightness/contrast sliders and numeric fields share the existing OOXML commands; the correction preset gallery is shared with the ribbon. Native Crop fields remain to be matched. Native Crop exposes picture width/height/offset X/Y plus crop width/height/left/top, rather than four crop percentages.

### Video pane color reset (2026-10-02)

- Native pane Reset removes grayscale, duotone (including Sepia), bi-level black/white, and brightness/contrast from the poster blip. It preserves the poster relationship, media, geometry, border and 3D formatting. Ribbon Reset additionally removes shape formatting and restores rectangular geometry.
- The pane now resets these color effects in one Undo action; opacity and unknown extensions remain intact. Core round-trip and browser tests cover reset, Undo/Redo and saved reload. The reference file was restored through named Undo operations and saved; its slide XML matches the baseline exactly. Native evidence is under `/tmp/pptx-video-reset-{before,after,ribbon,sepia-before,sepia-after,bilevel-before,bilevel-after}.xml`.

### Video recolor gallery (2026-10-02)

- Ribbon Color and pane Recolor presets share 21 choices with native grayscale, sepia, washout, threshold, and theme duotone values. More Variations exposes base theme colors, five rows of theme shades, and standard colors; native Accent 1 lighter 80% and darker 25% transforms are verified through saved XML. Selection is resolved with the same DrawingML color-transform code as rendering.
- Native saved XML confirms washout bright=70000/contrast=-70000, dark tint=45000/satMod=400000, and light shade=45000/satMod=135000. All temporary native document changes were undone and saved; the slide XML matches the baseline.
- Browser coverage verifies saved values, Undo/Redo, reload, live video effects, and preserved media/shape formatting. Core tests cover invalid-input atomicity, effect ordering before extLst, and transformed-color round trips.
- Remaining differences include custom-theme shade rules, exact gallery thumbnail rendering and native Crop fields. This is partial Video Format coverage, not complete PowerPoint parity.

- Washout rendering now uses the MS Office brightness/contrast order documented by LibreOffice Bitmap::Adjust: half brightness before contrast, half after. The native dark #262626 to approximately #D9D9D9 observation is covered by regression tests; arbitrary-image pixel parity still needs broader native comparison. Corrections and recolor thumbnails share the editor transfer calculation.

- Video brightness/contrast sliders now update the canvas during dragging and commit one history entry on release. Browser regression reproduces the old unchanged transfer before release, then verifies live rendering, saved OOXML and one-step Undo/Redo. Pointer cancellation rolls back the live edit.

### Video ribbon reset (2026-10-02)

- Native saved XML confirms ribbon Reset removes shape fill, line, effects and 3D formatting, restores ellipse to rectangle, and retains crop, dimensions and media/poster relationships. Evidence: `/tmp/pptx-video-ribbon-reset-{styled,result,oval,oval-result,cropped,crop-result}.xml`. Temporary changes were undone and saved; the slide XML matches the baseline.
- The Adjust group now exposes Reset as one Undo action. This is separate from Video pane color reset and Poster Frame reset.

- Browser regression covers reset, one-step Undo/Redo, save/reload and preservation of playback settings and media bytes. Core coverage checks unknown XML preservation and rejects non-video input without mutation; the command registry disables reset for unsupported selections.

## Video Crop picture and frame position

- The Video pane exposes picture width/height and center offsets separately from crop-frame width/height and absolute left/top, in centimeters. Picture edits hold the frame fixed; frame edits preserve the picture rectangle. Crop Reset expands the frame to the full picture.
- Mac PowerPoint saved XML verified both axes and Reset. Browser regression covers all eight fields, negative offsets, signed-percentage overflow rejection, undo/redo, save/reload and media/poster/playback preservation. Locked selections disable the controls.
- Rotated/grouped native coordinate behavior, crop drag handles, the ribbon crop menu and complete Video Format geometry remain unverified or incomplete. This is numeric Crop-pane coverage, not complete video UI parity.

## Table-cell text direction

- Preview now reads the existing cell `vert` property in both HTML and pure SVG paths, using the same direction mapping as shape text. Previously table cells always rendered horizontally despite preserving the property in OOXML.
- Inline editing follows all six stored vertical directions. Logical paragraph sizing preserves upright RTL glyph positions; bottom-to-top editing counter-rotates asymmetric insets so its text rectangle matches the preview.
- Browser comparisons cover per-character bounds on entry and editing/save retention in all six directions. Pure SVG tests verify saved clockwise/counterclockwise cells. Native direction calibration and existing pure-SVG upright/East-Asian approximations remain unverified; this is not evidence of full native parity.

## Editing tab kerning across formatting runs

- Reproduced right-tab misalignment at 586.72px for a 600px stop when `AV.12` spans multiple equally sized Arial runs. Browser shaping retains kerning across those run boundaries, including color changes; summing separate canvas widths did not.
- The editing layout now shapes consecutive runs with matching font metrics together, retaining the original DOM and UTF-16 offsets. Decimal positioning also retains kerning at the decimal boundary. Text fragments are collected before measurement to avoid repeatedly measuring growing prefixes.
- Focused browser regressions cover right/center/decimal alignment across color boundaries, plus existing capitalization, tracking, kerning and vertical-tab cases. This does not prove native ruler gestures, locale-specific decimal behavior or complex wrapping parity.

## Font color replacement on WordArt

- Setting a font color now replaces every DrawingML fill choice, including gradient and pattern fills. Clearing the direct color removes the local fill so inheritance can apply. Previously non-solid fills remained beside the new solid fill, producing invalid OOXML.
- Regression coverage includes all six fill choices, preservation of unrelated run properties, rejected-color preservation, and saved/reloaded PowerPoint-native gradient and pattern fixtures validated against the presentation schema.
- This fixes color mutation and serialization; gradient and pattern text rendering and the full WordArt gallery remain incomplete.

- `TextFormat.textFill` now represents gradient and pattern glyph fills, retaining theme references and color transforms. Effective formatting treats solid color and non-solid fills as one inherited choice.
- Pending typing formats in the canvas, outline, and notes now replace the old fill choice when a font color is selected. A browser regression verifies pattern-filled notes, a pending Bold command, a new font color, subsequent typing, and save/reload while preserving the original text's pattern. This does not establish gradient/pattern rendering or native WordArt gallery parity.

## Speaker-notes paragraph and soft breaks

- Enter inserts a paragraph separator; Shift+Enter inserts an OOXML `a:br` within the current paragraph. Notes editing tracks their UTF-16 positions independently so paragraph-end formatting is indexed by actual paragraphs.
- Replacing a selected paragraph separator with a soft break now commits even when the visible string is unchanged. The public notes API exposes separator kinds and accepts the inserted separator kind on ranged edits.
- Nine focused core tests pass, including emoji offsets, mixed-separator replacements and save/load/schema checks. Browser coverage verifies existing separator conversion, subsequent typing, save/reload and Undo. Native interaction comparison remains outstanding while screen capture returns size 0×0.
