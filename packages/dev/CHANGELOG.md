# @office-kit/pptx-dev

## 0.12.0

### Minor Changes

- 4410413: All 95 animation presets of the reference desktop app, written exactly as it writes them, and the Sequence (text build) option:

  - **Every gallery preset.** `setShapeAnimation`, `updateSlideAnimation` and `getSlideAnimations` cover the 35 entrance, 24 emphasis and 36 exit presets in the galleries of the reference desktop app (Mac, 16). Each writes the preset id, subtype, behaviours, default duration and build entry the reference desktop app writes for the same gallery item; a test compares every one against XML saved by the reference desktop app. New `AnimationEffect` tokens include `expandIn`, `swivelIn`, `basicZoomIn`, `centerRevolveIn`, `floatIn`, `growTurnIn`, `riseUpIn`, `spinnerIn`, `stretchIn`, `boomerangIn`, `bounceIn`, `creditsIn`, `curveUpIn`, `dropIn`, `flipIn`, `floatingIn`, `pinwheelIn`, `spiralIn`, `basicSwivelIn` and `whipIn`; the exits `contractOut`, `collapseOut`, `floatOut`, `shrinkTurnOut`, `sinkDownOut`, `stretchyOut`, `curveDownOut` and the counterparts of the entrances; and the emphasis effects `fillColor`, `fontColor`, `growShrink`, `lineColor`, `transparency`, `boldFlash`, `brushColor`, `complementaryColor`, `complementaryColor2`, `contrastingColor`, `darken`, `desaturate`, `lighten`, `objectColor`, `pulse`, `underline`, `colorPulse`, `growWithColor`, `shimmer`, `teeter`, `blink`, `boldReveal` and `wave`. An exhaustive `switch` over `AnimationEffect` needs the new cases.
  - **Breaking: `build` replaces `byParagraph`.** `setShapeAnimation` and `updateSlideAnimation` take `build: 'asOneObject' | 'allAtOnce' | 'byParagraph'` (the reference desktop app's Effect Options ▸ Sequence) instead of `byParagraph: boolean`, and `SlideAnimationStep.build` (`AnimationTextBuild | 'custom'`) replaces `buildByParagraph`. Replace `byParagraph: true` with `build: 'byParagraph'`. `'allAtOnce'` is new: it gives every paragraph its own effect, all starting together. Fill Color and Line Color take no build and refuse one.
  - **Breaking: `zoomIn` / `zoomOut` are now the reference desktop app's Zoom** (preset 53, a fade with a zoom). The plain scale they wrote before is the reference desktop app's Basic Zoom, now `basicZoomIn` / `basicZoomOut`.
  - **Changed output to match the reference desktop app.** `fadeIn` / `fadeOut` write the reference desktop app's fade filter instead of an opacity animation; `spin` defaults to 2 s; exit behaviours carry no `fill`; a shape with no fill and no line writes no `animBg`. `durationMs` now rescales every behaviour of a multi-behaviour preset together, as the reference desktop app's Duration box does; Transparency and Bold Reveal hold until the end of the slide and refuse a duration.
  - **Editor.** Every tile of the Entrance and Emphasis galleries and every Exit Effects item is enabled, in the reference desktop app's order and with its English and Japanese names. Each writes the preset's own default duration. Effect Options ▸ Sequence offers As One Object, All at Once and By Paragraph, and the animation pane lists every preset and has a Sequence choice. The slide show and Preview play an approximation of the new presets (a pose from scale, rotation, translation and opacity, or a colour filter for the emphasis colour effects).

- d8e455d: The reference desktop app's 74 built-in table styles are now fully supported.

  - `setTableStyleId` accepts a built-in style's English name as well as a GUID (`setTableStyleId(table, 'Light Style 1 - Accent 2')`). Applying a built-in style writes the reference desktop app's definition of it into `ppt/tableStyles.xml` (creating the part when a deck has none), as the reference desktop app does, so Keynote, Google Slides and LibreOffice draw the same table. A deck's own definition of a GUID is kept. `addSlideTable` likewise writes the definition of its default style.
  - New `BUILTIN_TABLE_STYLES` lists every built-in style (`id`, `name`, gallery `category`) in the order of the reference desktop app's Table Styles gallery.
  - `getTableCellAppearanceEffective` now resolves every built-in style, not only Medium Style 2 - Accent 1 and No Style, Table Grid, and reports `fillOpacity` for translucent fills. New `getTableBackgroundEffective` resolves the table background (`a:tblPr` fill, else the style's `a:tblBg`).
  - The preview draws all table style parts for built-in and custom styles, including translucent bands and the Themed Styles' background.
  - The editor's Table Design ▸ Table Styles gallery offers all built-in styles, grouped as in the reference desktop app (Best Match for Document, Light, Medium, Dark), with swatches drawn by the preview renderer that follow the Table Style Options check boxes, the reference desktop app's English and Japanese names, and Clear Table.

- 1be45ba: The editor's right-click menus for text, table cells, pictures and the slide background now match the reference desktop app (Mac):

  - **Text being edited** gets the reference desktop app's menu instead of the browser's: Cut, Copy, Paste, Exit Edit Text, Font..., Paragraph..., Bullets ▸ and Numbering ▸ galleries, Format Shape..., Lock, Hyperlink... and New Comment. The menu leaves the caret and selection in the text, and its commands act on the selected range.
  - **Table cells** offer Insert ▸ (columns left/right, rows above/below), Delete ▸ (columns, rows, table), Select ▸ (table, column, row), Merge Cells and Split Cells..., plus the text commands. Clear cell text, Select all cells and Select table left the menu (Delete, ⌘A and Select ▸ Select Table do the same).
  - **Pictures** offer Change Picture ▸ From a File..., Crop, Format Picture... and the reference desktop app's other picture items.
  - **The slide background** offers Paste Special..., New/Duplicate/Delete Slide, Hide Slide, Ruler, Grid and Guides ▸ toggles, Zoom..., Format Background..., Slide Show and New Comment; Layout ▸, Reset Slide and Select all left it (they remain on the Home tab and ⌘A).

  Menus show the reference desktop app's Mac shortcut hints, including in submenus; commands the editor cannot perform yet are disabled with the reason as a tooltip.

- 375fa71: The editor's contextual ribbon tabs now match the reference desktop app's (Mac) layout and sizes:

  - **Shape Format** shows the reference desktop app's in-ribbon shape strip, Quick Styles and Text Art strips, a large Shape Fill / Text Fill button with Outline and Effects menus beside it, and the expanded Arrange group (Bring Forward, Send Backward, Selection Pane, Reorder Objects, Align, Group, Rotate). Text Box ▾ adds a Vertical Text Box, and Size gains Lock Aspect Ratio. Below 1300 pt the Insert Shapes and Arrange groups collapse as they do in the reference desktop app.
  - **Picture Format** is new: a selected picture shows it instead of Shape Format, with Corrections, Color and Transparency presets, Change Picture, Reset Picture, Picture Border, Picture Effects, Crop (with Crop to Shape), Size and Format Pane. Remove Background, Artistic Effects, Compress Pictures, Picture Quality, picture styles, Picture Layout and Animate as Background are shown disabled with the reason.
  - **Table Design** gains the six Table Style Options check boxes, a table style strip (No Style, No Grid; No Style, Table Grid; Medium Style 2 - Accent 1), Shading, Borders (twelve edge choices drawn with the Pen Style, Pen Weight and Pen Color settings) and Text Art styles for the selected cells.
  - **Table Layout** (renamed from Layout) gains Select, Delete ▾ (columns, rows or the table), Insert Row Above/Below, Insert Column Left/Right, Split Cells, Row and Column size boxes, Distribute Rows/Columns, the six cell alignment toggles, Text Direction, Cell Margins presets, Table Size and the Arrange group. Row and column changes keep the table's frame the size of its grid.
  - Japanese labels use the reference desktop app's (Mac) wording; Format Pane is now 書式ウィンドウ and Reorder Objects オブジェクトの並べ替え.

- c8512fe: Animations tab:

  - Effect Options lists the emphasis effects' own options. Spin has Direction and Amount (Quarter Spin to Two Spins), Grow/Shrink has Direction and Amount (Tiny, Smaller, Larger, Huge), and Transparency has Amount (25–100%). The colour effects show the theme and standard colour palette. The animation pane offers the same choices.
  - Exit Effects opens a gallery of tiles under the reference desktop app's group headings (Basic, Subtle, Moderate, Exciting), like Entrance and Emphasis. A narrow ribbon's Emphasis Effects button opens the same kind of gallery.
  - Picking a different effect gives it that effect's own default duration instead of keeping the old one.
  - Preview plays the spin angle, grow/shrink size and transparency amount that are set. Multi-part presets such as Bounce, Boomerang, Center Revolve, Rise Up, Float, Drop, Flip, Whip, Curve Up, Teeter, Wave, Pulse and Blink now follow each part's own timing.

- c535941: feat: copy slides and objects to the system clipboard. Copying slides or objects in the editor now puts them on the clipboard for other presentation apps and documents: slides and drawn objects paste there as pictures at their size on the slide, and text boxes and tables copied on their own paste as text and tables. Another editor — in another tab or browser — pastes them back as editable slides and objects, with their pictures, layouts and themes. Pictures and text copied in other apps paste as a picture or a text box.

  feat: the slide thumbnail menu has **Download Selected Slides...**, which saves just the selected slides as a `.pptx` to import into another presentation app.

  feat: `renderSlideToSvg(pres, slide, { background: false })` draws only the slide's own shapes on a transparent surface, without the background or master and layout graphics.

- 2912bdc: The editor's Format Shape pane now follows the reference desktop app (Mac) more closely:

  - **Shape Options / Text Options** switch at the top. Text Options has the reference desktop app's Text Fill & Outline, Text Effects and Textbox categories and applies to the text of the selected shapes.
  - **Effects** shows the reference desktop app's Shadow, Reflection, Glow, Soft Edges, 3-D Format and 3-D Rotation sections with their preset galleries, color buttons, sliders and boxes, instead of the editor's list of effect commands. Text Effects offers the same sections for the text (Soft Edges and Keep text flat are shown disabled, with the reason, because the library cannot write them yet).
  - **Every section starts collapsed**, as in the reference desktop app, and stays as you left it for the session. Size and Position... still opens the Size and Position sections.
  - The Line section gains **Sketched style** (shown disabled: the library does not write sketched lines yet), and the Begin/End Arrow type and size buttons are the reference desktop app's 39 pt gallery buttons. Arrow sizes are named Arrow L Size 1–9 / Arrow R Size 1–9 as in the reference desktop app.
  - The flip checkboxes left the Position section (the reference desktop app has none there); Arrange ▸ Rotate keeps Flip Vertical and Flip Horizontal.

- 550b55c: Gradient lines, sketched lines and more text paints:

  - **Gradient lines.** `setShapeStroke(shape, { fill: { kind: 'gradient', ...gradient } })` writes `<a:ln><a:gradFill>` with the same gradient options as `setShapeGradientFill`, and a text outline takes the same `fill` (`setShapeTextFormat(shape, { outline: { fill } })`). `getShapeStrokeGradient(shape, pres?)` reads it back; `getShapeStroke` / `getShapeStrokeEffective` now report `{ kind: 'gradient' }` for such lines (previously `inherit`), so a `switch` over `ShapeStroke['kind']` needs a `gradient` case.
  - **Sketched lines.** `setShapeStrokeSketch(shape, 'curved' | 'freehand' | 'scribble' | null)` / `getShapeStrokeSketch(shape)` write and read the reference desktop app's Sketched style (`ask:lineSketchStyleProps`). Other `<a:ln>` extensions are kept, and a sketch the reference desktop app saved (which replaces the geometry with the hand-drawn path) gets its original geometry back when the sketch is changed or removed.
  - **Text fills.** `TextFormat.textFill` accepts `{ kind: 'none' }` (`<a:noFill/>`) and `{ kind: 'image', bytes }` (a PNG, JPEG, GIF, BMP, TIFF or WebP stretched over the text, as the reference desktop app writes it), on shapes, table cells and notes.
  - **Keep text flat.** `setShapeTextFlat(shape, flat)` / `getShapeTextFlat(shape)` write `<a:bodyPr><a:flatTx/>`.
  - **Preview** strokes gradient outlines on shapes, connectors and text, paints no-fill and picture-filled text, and draws a hand-drawn approximation of sketched outlines.
  - **Editor**: Gradient line in Line and Text Outline (with the gradient type, direction, angle and stop controls of Gradient fill), the Sketched style menu (None, Curved, Freehand, Scribble), No fill and Picture or texture fill for text, and Keep text flat are now enabled. Text Effects ▸ Soft Edges stays disabled, as in the reference desktop app. The Text Fill gradient and pattern options no longer fail to apply.

- 6359001: The editor's Insert, Draw and Design tabs now match the reference desktop app's (Mac) layout:

  - Insert has the reference desktop app's ▾ menus: Table offers a grid that inserts a table of the chosen size; Pictures, Screenshot, Video and Audio list the reference desktop app's sources, with those the browser cannot reach shown disabled; Chart opens the chart dialog on the chosen type; Text Box ▾ can also draw a vertical text box. In a narrower window 3D Models, SmartArt and Chart become small rows and Date & Time, Slide Number and Object small icons, as in the reference desktop app.
  - Draw has Eraser ▾ (Stroke Eraser) and Add ▾ with Add Pen, Add Pencil and Add Highlighter.
  - Design's Colors and Fonts menus list all of the reference desktop app's built-in color sets and font pairs, with the deck's current ones on top; Slide Size marks the current size; the Themes gallery fills the window's width. The Legacy (2007 - 2010) fonts are now Calibri / Calibri, as the reference desktop app lists them.
  - Button sizes, group spacing and two-line captions follow the reference desktop app in English and Japanese.

- afe80eb: Slide, notes and handout masters can now be edited:

  - Slide masters: `addSlideMaster`, `removeSlideMaster`, `getSlideMasterLayouts`, `getSlideMasterPlaceholders`, `getSlideMasterName` / `setSlideMasterName`, `isSlideMasterPreserved` / `setSlideMasterPreserved` and `setSlideMasterPlaceholderIncluded`. A new master gets the reference desktop app's default structure: eleven default layouts and its own copy of the theme.
  - Layouts: `addSlideLayout`, `removeSlideLayout`, `addSlideLayoutPlaceholder` (content, text, picture, chart, table, SmartArt, media and online image, plus vertical content and text), `setSlideLayoutTitleIncluded`, `setSlideLayoutFootersIncluded` and `setSlideLayoutBackgroundGraphicsHidden`.
  - Removing a master or layout that slides still use throws, as does removing the last one.
  - Notes and handout masters: `getNotesMasterPlaceholders` / `setNotesMasterPlaceholderIncluded`, `getHandoutMasterPlaceholders` / `setHandoutMasterPlaceholderIncluded`, `getNotesPageSize` / `setNotesPageOrientation` and `getHandoutSlidesPerPage` / `setHandoutSlidesPerPage`. If the deck has no such master, the first edit creates the reference desktop app's default one.
  - `@office-kit/pptx-preview` adds `renderSlideLayoutToSvg`, which draws a layout's or master's background and decorative shapes without its placeholders.
  - In the editor, the Slide Master, Handout Master and Notes Master tabs no longer have disabled commands:
    - Insert Slide Master, Insert Layout, Delete, Rename, Preserve, Master Layout and Insert Placeholder all work, as do the Title, Footers and Hide Background Graphics checkboxes.
    - Orientation, the placeholder checkboxes and Slides Per Page also work.
    - The master views draw the deck's decorative shapes and its real placeholder positions.

- af4e181: The editor's title bar now carries the reference desktop app's (Mac) menu bar — File, Edit, View, Insert, Format, Arrange, Tools, Slide Show, Window and Help — in place of the Edit and View buttons:

  - **Menus match the reference desktop app** in English and Japanese: order, separators, submenus and shortcut glyphs, with items enabled and checked as the reference desktop app's are with nothing, a shape or text selected (Undo names the last edit; Apply To Defaults becomes Apply Object Style when something is selected). Items run the editor's existing commands; what the editor cannot do yet is disabled with the reason as a tooltip. While text is being edited, the menus leave the caret and selection in the text.
  - **Keyboard shortcuts follow the reference desktop app's**, from the same table the menus show. Changed keys: ⌘G is Find Next, so Group is ⌥⌘G and Ungroup ⌥⇧⌘G; ⌘K inserts a hyperlink and the command search moved to ⌘? (Help ▸ Editor Help); Pick Up / Apply Object Style are ⇧⌘C / ⇧⌘V (⌥⌘C to pick up in Japanese) instead of ⌥⌘C / ⌥⌘V; Paste and Match Formatting is ⌥⇧⌘V (was ⇧⌘V); Replace is ⌃H (was ⌘H). While editing slide text, ⌘L/⌘E/⌘R align it, and ⌘T, ⌘K, ⌥⌘M and the object-style keys act on the selected text. Control still stands in for Command where the reference desktop app has no Control shortcut of its own.
  - The command search now closes with Escape.

- 13c84de: The editor's window now matches the reference desktop app's (Mac) measurements:

  - **Zoom percentages** mean what they mean in the reference desktop app (Mac): at 100% a slide point is one screen point (a widescreen slide is 960 px wide; it was 1280 px), so Fit to Window now reads about 120% in a full-size window instead of about 76%. Fit leaves the reference desktop app's 22 pt margin around the slide.
  - **Home ribbon**: a 72 pt command row with the reference desktop app's button sizes and group spacing; Add-ins and Designer are separate groups; the second Font row is in the reference desktop app's order (… Character Spacing, Change Case, then Text Highlight Color and Font Color). The extra Font dialog (A…) button is removed — use Cmd+T or Character Spacing ▸ More Spacing....
  - **Layout**: the thumbnail pane is 249 pt wide, the notes pane opens one line tall, and the status bar uses the reference desktop app's sizes.
  - **Format Shape pane**: a compact title, larger category tabs, chevron section headers, and one-line label / control rows with the reference desktop app's 26 pt controls and 112 pt pop-ups. The Line section now starts with No line / Solid line / Gradient line, replacing the No outline button, and hides its settings for No line.
  - **Right-click menus** on objects and slide thumbnails list the reference desktop app's items in its order, with its separators and 24 pt rows. Objects gain Lock/Unlock, Reorder Overlapping Objects and Action Settings...; Link... and Edit Alt Text... are now Hyperlink... and View Alt Text...; commands the editor cannot perform yet are shown disabled with the reason. Thumbnails gain Select All, Zoom... and Slide Show, and no longer list Layout or Reset Slide (they remain on the Home tab).

- d92bbd7: Outline View now edits across slide boundaries as the reference desktop app's (Mac) outline does:

  - **Select across slides** by dragging with the mouse or Shift-clicking, as well as with Shift+arrow keys.
  - **Delete, Cut, typing, pasting and Enter** over a selection that crosses a slide title remove the slides whose titles are selected: the remaining text after the selection joins the paragraph where it starts, and the last slide's remaining body moves up. This now works from any title or body paragraph, not only from a title. Slides with other objects ask for confirmation first.
  - **Backspace** at the start of a slide title (or **Delete** at the end of the text before it) merges that slide into the previous one.
  - **Drag a bullet** to move the paragraph and its sub-points to any position, including another slide; dragging sideways changes their level, and dragging a top-level bullet left turns it into a new slide.
  - **Drag a slide icon to the right** to demote the slide into the previous slide's body.
  - **Drag selected text** to move it, or hold Option (Control elsewhere) to copy it, including into another slide.

  Each of these is one Undo step.

- acb95fc: Picture Format: Compress Pictures, Reset Picture & Size and the Artistic Effects gallery.

  - New `getShapeImageArtisticEffect` reads the Artistic Effect the reference desktop app applied to a picture or image fill (`a14:imgProps`, e.g. `'pencilSketch'`), or `null`. The embedded picture is already the effect's result, so the preview keeps drawing it as is; the effect, its JPEG XR original and the relationship to it survive edits, duplication and saving.
  - Fix: setting a picture's transparency, brightness, contrast or recolor wrote the effect after the picture's `a:extLst`, which is schema-invalid; it now goes before it.
  - The editor's Compress Pictures (Picture Format and File ▸ Compress Pictures...) offers the reference desktop app's Picture Quality choices (High Fidelity, HD 330, Print 220, On-screen 150, Email 96 ppi, Use Original Quality), Delete cropped areas of pictures and Apply to. It downsamples PNG and JPEG pictures in the browser to the chosen resolution of their frame and removes cropped-away pixels, as one undo step.
  - Reset Picture ▸ Reset Picture & Size now works: it also removes the crop and restores the picture's natural size at its own resolution.
  - Artistic Effects shows the reference desktop app's gallery (English and Japanese names) with the picture's current effect checked. Applying an effect stays unavailable: the reference desktop app stores its own rendering plus a JPEG XR original, which the editor cannot produce.

- 21d59f7: Picture Styles, and Compress Pictures that labels pictures the way the reference desktop app does.

  - New `setShapePictureStyle(picture, 'Metal Oval')` applies one of the reference desktop app's 28 built-in picture styles. It writes exactly the `p:spPr` markup the reference desktop app (Mac, 16.113) saves for that style (geometry, fill, border, effects and 3-D, with literal colors) and keeps the picture, its crop and its position. `getShapePictureStyle` returns the style a picture carries exactly, or `null`. New `BUILTIN_PICTURE_STYLES` lists the style names in the order of the reference desktop app's gallery.
  - New `setShapeImageCompressionState` / `getShapeImageCompressionState` write and read the picture's `a:blip/@cstate` (`'print'`, `'screen'`, `'email'`, …) together with the `a14:useLocalDpi` extension the reference desktop app writes beside it.
  - Preview: pictures now draw their effects (outer and inner shadow, glow, soft edge, reflection), their own fill and an approximation of their 3-D: the camera rotation as a flat projection and a top bevel as edge lighting.
  - Fix (preview): a reflection faded the wrong way, strongest at its far edge, and ignored its end position; it now starts at the shape's edge and fades out by `endPos`. A soft edge blurred the whole shape; it now only feathers the outline.
  - The editor's Picture Format ▸ Picture Styles gallery works: the 28 styles in the reference desktop app's order, drawn by the preview renderer, with the reference desktop app's English and Japanese names as tooltips and the applied style checked.
  - The editor's Compress Pictures writes `cstate` for Print, On-screen and Email, as the reference desktop app does, and only replaces a picture's pixels when cropped areas are removed or the resampled picture is smaller. Picture Quality now opens the reference desktop app's menu: Compress Pictures... and Upscale Picture (unavailable: it uses the vendor's cloud AI service).

- 1d447f8: Editor ruler: indent and tab markers now also appear for rotated, flipped and vertical text. Rotated text is measured along its own lines as if the shape were unrotated, and vertical text (`vert`, `eaVert`, `vert270`, …) is measured on the vertical ruler. While you drag an indent marker or a tab stop, the text reflows immediately; the change is still saved as one undo step on release, and Escape restores the original layout. With several paragraphs selected, the ruler shows the first paragraph's markers; dragging moves each paragraph relative to its own indents, and moving a tab stop changes only the paragraphs that have it. While editing, decimal tabs align on the run language's decimal separator, such as `,` for German.
- 842cfd9: The editor's Transitions, Animations, Slide Show, Record, Review and View tabs now match the reference desktop app (Mac):

  - Every tab uses the reference desktop app's 72 pt command row, button widths, groups and order. Custom Show and Record (Slide Show tab) and Delete (Review tab) open menus, and the menu buttons show ▾.
  - The Transitions gallery lists all of the reference desktop app's transitions in its order, shows as many tiles as fit (10 at 1512 pt, 6 at 1200 pt) and pages with ‹ ›. Transitions the library cannot write (2010-and-later effects such as Morph) are shown disabled with a reason. Effect Options opens the reference desktop app's per-effect menu (for example Push: From Bottom / Left / Right / Top, Wipe and Cover: eight directions, Shape: Circle / Diamond / Plus, Fade: Smoothly / Through Black) instead of a dialog. Duration, Sound, On Mouse Click, After and Apply To All form one Timing group, and Duration can be set for a slide with no effect, as in the reference desktop app. Gallery tiles now write the reference desktop app's default options (Push From Bottom, Split Vertical Out, Random Bars and Blinds Vertical, Clock Clockwise).
  - The Animations tab has separate Entrance and Emphasis galleries with the reference desktop app's effects (unsupported ones disabled). At narrower widths the Emphasis gallery collapses into an Emphasis Effects ▾ button. Effect Options is now the reference desktop app's Direction / Sequence menu (From Bottom, Left, Top or Right, plus As One Object or By Paragraph). The galleries no longer have a None tile; remove an effect from the Animation Pane, as in the reference desktop app.
  - Japanese uses the reference desktop app's (Mac) wording for the new gallery and menu items.

- 24a6ae0: The 2010-and-later transitions and the filter animations of the reference desktop app:

  - **Transitions.** `setSlideTransition` writes the 2010-and-later transition effects — `vortex`, `switch`, `flip`, `ripple`, `honeycomb`, `prism`, `doors`, `window`, `ferris`, `gallery`, `conveyor`, `pan`, `glitter`, `warp`, `flythrough`, `flash`, `shred`, `reveal` and `wheelReverse` — the twelve `prstTrans` presets (`{ effect: 'prstTrans', preset: 'curtains' }`, with `invertX` / `invertY`) and Morph (`{ effect: 'morph', morphOption: 'byObject' | 'byWord' | 'byChar' }`). They are written the way the reference desktop app writes them, inside `mc:AlternateContent` with a `<p:fade/>` fallback for readers that do not know them. New options: `pattern` (glitter, shred), `isContent`, `isInverted`, `hasBounce` and the directions each effect takes. `getSlideTransition` reads all of them back; an extension effect the library does not know is reported as `prefix:local` (for example `p99:sparkle`) and kept on the slide. `TransitionEffect` gained these tokens, so an exhaustive `switch` over it needs the new cases.
  - **Animations.** `setShapeAnimation` writes the reference desktop app's filter entrance and exit effects with the preset ids, subtypes, filters and default durations the reference desktop app writes: `wipeIn` / `wipeOut`, `peekIn` / `peekOut`, `splitIn` / `splitOut`, `blindsIn` / `blindsOut`, `checkerboardIn` / `checkerboardOut`, `randomBarsIn` / `randomBarsOut`, `shapeIn` / `shapeOut`, `stripsIn` / `stripsOut`, `wheelIn` / `wheelOut`, `dissolveIn` / `dissolveOut` and `wedgeIn` / `wedgeOut`. New options `orientation`, `inOut`, `shape` and `spokes` (and `direction` for wipe, peek and strips) are accepted by `setShapeAnimation` and `updateSlideAnimation` and reported by `getSlideAnimations`. `flyIn` / `flyOut` take the four diagonal directions (`'topLeft'`, `'topRight'`, `'bottomLeft'`, `'bottomRight'`). `AnimationEffect` and `AnimationDirection` gained these tokens, so an exhaustive `switch` over them needs the new cases. An option given to an effect that does not take it throws, as `direction` always did. Without `durationMs`, shape, wedge and wheel run 2 s (the reference desktop app's default); everything else keeps 500 ms.
  - **Editor.** All 49 transitions and every Effect Options item are enabled, each tile writing the reference desktop app's default option and duration; the slide show plays an approximation of the 2010+ effects. The Entrance gallery enables Blinds, Checkerboard, Dissolve In, Peek In, Random Bars, Shape, Split, Strips, Wedge, Wheel and Wipe; Exit Effects lists the reference desktop app's exit gallery with their counterparts enabled; Effect Options offers each effect's directions, shapes and spokes, including Fly's diagonals. Preview plays the filter effects as an animated clip. All at Once and the remaining motion and emphasis presets stay disabled with a reason.

- f6395fd: The editor's views now match the reference desktop app (Mac):

  - The status bar view switcher has the reference desktop app's four buttons: Normal, Slide Sorter, Reading View and Slide Show. None is selected in Notes Page or the master views. The status text shows "Notes N of M" on the notes page and the master's name in master views. The Notes and Comments buttons appear only in Normal and Outline View.
  - Slide Sorter opens at 80% and lays out thumbnails like the reference desktop app: six to a row in a full-width window, centred, with slide numbers below.
  - Outline View uses the reference desktop app's indents: titles 36 pt in and 11.5 pt per level. Show Formatting is on by default and draws text at a third of its size. The outline menu adds Hyperlink… and the reference desktop app's item order.
  - Slide Master view shows each master and its indented layouts. Its Slide Master tab can rename layouts and change the theme, colors, fonts, background and slide size. A layout's placeholders can be moved by dragging or with the arrow keys. Commands the library cannot perform yet are shown disabled with the reason.
  - Handout Master and Notes Master views show the reference desktop app's default master pages and ribbons. They are view-only for now.
  - Notes Page view shows the whole portrait page, fitted to the window.
  - Reading View opens a full-window reading mode in the standalone editor.
  - ⌘1–⌘5 and ⌥⌘1–⌥⌘3 switch views as in the reference desktop app.
  - Japanese uses the reference desktop app's wording for all of the above.

### Patch Changes

- 88105a3: TSX rebuilds now reach the open editor as one undo step ("Source changed") merged with unsaved canvas edits, instead of reloading the deck and starting a new Undo history. A collision is offered as before (**Keep my edits** / **Use source**), and **Use source** can be undone in the editor.
- 1c7d8ee: Agents in the browser can now read and edit the embedded editor's presentation through the `mountEditor` handle:

  - `selection()` returns the shapes the user selected as `ShapeRef`s (`{ slideIndex, slide, shapeId, name }`), and `resolveShape(presentation, ref)` finds such a shape again, throwing once it has been deleted.
  - `apply(label, edit)` runs `edit(presentation)` with the `@office-kit/pptx` API as one undo step named "Agent: label" (「エージェント: label」 in Japanese). An edit that throws is rolled back completely and `apply` rejects with its error.
  - `on('change', listener)` reports every kept edit with its source (`'user'` or `'agent'`), including Undo and Redo; `on('selectionchange', listener)` reports the user's selection.

- 5e6bbc5: Built-in theme names are now neutral.

  - `createPresentation()` writes its theme as "Default Theme", with color, font and format schemes named "Default". Decks created earlier keep the names they were saved with.
  - The editor's Design tab names the first three themes Standard Theme, Classic Theme and Legacy Theme (標準テーマ, クラシック テーマ, レガシー テーマ), and the matching color sets and font pairs Standard, Classic and Legacy (標準, クラシック, レガシー). Picking one writes the new name into the deck. When an opened deck's color scheme uses the name another presentation app gives one of these sets, Theme Colors shows the editor's name for it.

- 74e4e7d: The editor and the dev tool no longer show third-party product names or our own branding in their UI.

  - **Editor.** The title bar starts with the File menu; the "◈ @office-kit/pptx Editor" mark is gone, so an embedded editor carries no brand. The Help menu's product-named help item is now "Editor Help" (エディター ヘルプ), the Tools menu's product-named add-ins item is "Add-ins..." (アドイン), and Share ▸ Send a Copy offers "PPTX Presentation". Tooltips for unavailable features describe what is missing (for example "Translation needs an online translation service.", "Macros (VBA) do not run in this editor.", "Soft edges are not available for text.") instead of naming a third-party product or service.
  - **Dev tool.** The preview, presenter and editor pages are titled "Presentation preview", "Presenter view" and "Presentation editor", and the scaffolded project instructions and agent prompts say "presentation" rather than naming a product.

- 3f4aefe: The READMEs and npm package descriptions describe compatibility in terms of presentation apps in general instead of naming a third-party product. The root README gains a Trademarks section. No code behaviour changes.
- ec775df: The editor's styled-text gallery is now called "Text Art" (テキスト アート) instead of using a third-party feature name: Insert ▸ Text Art, Shape Format and Table Design ▸ Text Art Styles (テキスト アートのスタイル), Text Art Quick Styles (テキスト アートのクイック スタイル) and Clear Text Art (テキスト アートのクリア). The presets and what they write are unchanged.
- 54b0607: The editor's ruler and texture defaults now follow what the reference desktop app (Mac) does:

  - **Ruler with several paragraphs selected** shows the last selected paragraph's indent markers and tab stops, as the reference desktop app does, instead of the first paragraph's.
  - **Picture or texture fill** inserts the last texture picked for a shape in this session (starting with Papyrus) for both shapes and slide backgrounds. Textures picked in Format Background do not change it. Like the reference desktop app's, which resets on relaunch, it is not saved and returns to Papyrus when the editor page is reloaded.

- 197b737: Transitions are saved the way the reference desktop app (Mac) saves them:

  - **`setSlideTransition` with `durationMs`.** `spd` is now the fastest speed at least as long as the duration (≤ 0.5 s fast, ≤ 0.75 s medium, otherwise slow). `fast` is left out because it is the schema default. A duration equal to its speed's own (500, 750 or 1000 ms) is written as that speed alone, with no `p14:dur` and no `mc:AlternateContent` unless the effect needs one. It reads back as `speed` without `durationMs`. Any other duration is written as before.
  - **Editor.** Each Transitions gallery tile now writes the element, attributes and duration the reference desktop app (Mac, 16) saves for it, and the ribbon Duration shows the reference desktop app's value. Twenty-three durations changed, for example Reveal 3.40, Curtains 6.00, Honeycomb 4.40, Shape 0.80 and Zoom 0.90. Default directions the reference desktop app does not write are no longer written, for example on Split, Reveal, Ripple, Shred, Cube and Fly Through, and Wind, Airplane and Origami no longer write `invX`. Ten Japanese gallery names now match the reference desktop app, for example 垂れ幕, 破砕, ハチの巣, 細分, 扉 and 窓.
  - **Durations of transitions without `spd`.** The ribbon Duration, the Slide transition dialog's speed and the slide-show preview now use the schema default (fast, 0.5 s) for a transition without `spd`. They used medium (0.75 s) before.
  - **Effect Options.** Choosing an option keeps a duration that is stored only as a speed. Before, the transition fell back to fast.

- Updated dependencies [c8512fe]
- Updated dependencies [4410413]
- Updated dependencies [d8e455d]
- Updated dependencies [0d26527]
- Updated dependencies [1c7d8ee]
- Updated dependencies [52ff93e]
- Updated dependencies [c8512fe]
- Updated dependencies [167bdc1]
- Updated dependencies [88105a3]
- Updated dependencies [c535941]
- Updated dependencies [550b55c]
- Updated dependencies [afe80eb]
- Updated dependencies [5e6bbc5]
- Updated dependencies [74e4e7d]
- Updated dependencies [3f4aefe]
- Updated dependencies [946dac4]
- Updated dependencies [9d0754c]
- Updated dependencies [af5b2b4]
- Updated dependencies [ec775df]
- Updated dependencies [d92bbd7]
- Updated dependencies [1d447f8]
- Updated dependencies [acb95fc]
- Updated dependencies [21d59f7]
- Updated dependencies [fc1acfa]
- Updated dependencies [1d447f8]
- Updated dependencies [2912bdc]
- Updated dependencies [ec775df]
- Updated dependencies [6ac3ead]
- Updated dependencies [197b737]
- Updated dependencies [24a6ae0]
- Updated dependencies [9cad210]
  - @office-kit/pptx@0.24.0
  - @office-kit/pptx-preview@1.0.0
  - @office-kit/pptx-editor@0.1.0
  - @office-kit/pptx-dsl@1.0.0

## 0.11.0

### Minor Changes

- b25a586: The editor now has the reference desktop app's Texture gallery. Under **Picture or texture fill** in Format Shape and Format Background, and in the ribbon's **Shape Fill ▸ Texture**, pick one of the 24 textures (Papyrus, Canvas, Denim, … Medium wood) to tile it across the selected shapes or slides with the reference desktop app's default tiling, in one undo step. **More Textures...** chooses a picture file. The textures are generated by the editor to resemble the reference desktop app's; they are not copies of its images.

### Patch Changes

- 4f9445b: Editor: choosing Picture or texture fill for a shape or slide background that has no picture now fills it with the default Papyrus texture, tiled as in the reference desktop app, in one undo step, instead of opening a file chooser. Use Insert... to choose a picture file. A shape filled with a picture or texture now shows the Format Picture pane title. The Texture gallery now matches the reference desktop app (Mac): five textures per row at the reference desktop app's size and spacing, no texture marked as the current fill, and no More Textures... in the Format pane (use Insert...; the ribbon's Shape Fill ▸ Texture keeps it). In Format Shape and Format Background, Texture is now a row with a right-aligned swatch ▾ button, and the gallery opens aligned to its right edge.
- 089140c: The editor canvas shades text that has a Soft Bevel or Sharp Bevel Text Art style (or any other text bevel) instead of drawing it flat, including while the text is being edited.
- da5eee5: Shapes drawn from the editor's Shapes gallery are visible again and match the reference desktop app: they get the theme's "Colored Fill - Accent 1" style (accent fill, darker outline, light text), brackets and braces get the outline-only line style with dark text, text typed into them is centered both ways, and they are named like "Oval 3". Previously every inserted shape had neither fill nor outline, so only its selection box showed.
- e860ae8: Preset shapes are now drawn from ECMA-376's own preset definitions instead of hand-written approximations. The preview, the editor canvas and its Shapes gallery icons showed many presets wrongly — hearts, lightning bolts, suns, moons, clouds, brackets and braces, bent and curved arrows, the equation shapes, stars with 7 to 32 points — and now match the reference desktop app, including the lit and shaded faces of cubes, cans and curved arrows, adjust handles on every preset, and elliptical arcs in custom geometry (whose `arcTo` angles were read as parametric instead of visual angles). Text in a preset now wraps inside the rectangle the definition gives it, such as an ellipse's inscribed rectangle, as the reference desktop app does.

  - `@office-kit/pptx`: new `getPresetGeometry(preset, size, adjustValues?)` evaluates any preset's paths and text rectangle in the same form as `getShapeCustomGeometry`.
  - `@office-kit/pptx-preview` (breaking): `shapeCustomTextRect(custom, extent)` is replaced by `shapeTextRect(shape)`, which reads custom and preset geometry alike, and `resolveTextBodyRect` now takes `(bounds, margins, region)` with the region from `shapeTextRect`.

- efdf5f3: Ribbon dropdowns (Design ▸ Colors, Fonts, Layout and Slide Size; Home ▸ Section, Columns, Text Direction and Align Text; Shape Format ▸ Edit Shape, Shape Effects and Text Effects; Review ▸ Language and Delete; Animations ▸ Exit Effects) open over the slide below their button instead of being clipped inside the ribbon, which also no longer grows a vertical scrollbar (#406).
- Updated dependencies [e860ae8]
- Updated dependencies [089140c]
- Updated dependencies [8db2360]
  - @office-kit/pptx@0.23.0
  - @office-kit/pptx-preview@0.14.0
  - @office-kit/pptx-dsl@0.10.0

## 0.10.0

### Minor Changes

- e18d19d: Apply the active slide's background to the whole presentation with `applySlideBackgroundToAll` or the editor's Apply to All button. Backgrounds move to slide masters and individual slide/layout overrides are cleared, preserving theme colors and image references. The editor supports undoing the whole operation together.
- e18d19d: Edit slide background gradients in the slide options panel, including stops, direction, position, transparency and brightness. Changes apply to the selected slides and support undo and saved reloads.
- e18d19d: Edit slide background patterns with a pattern gallery and foreground/background theme color controls. Selected slides support undo, reset and saving. Editing an inherited pattern creates a slide override while preserving the layout or master and the original color transforms.

  Background pattern readers accept `preserveTheme` to retain unmodified theme color references.

- e18d19d: Add `setSlideBackgroundPatternFill` to edit background patterns and colors. Partial edits preserve unspecified pattern colors and their original DrawingML transforms.
- e18d19d: Add a Japanese and English link editor that shows existing addresses and descriptions, applies links to selected text shapes, and removes links with undo/redo and saved-project persistence.
- e18d19d: Select table-cell ranges on the canvas using Shift-click or dragging while a cell is selected. Highlight the entire selected range, including merged cells, without moving the table.
- e18d19d: Choose text formatting at the caret before typing in shapes and table cells. Bold, italic, underline, font, size, color, highlight, and clear formatting apply to newly entered text, preserve surrounding text, and participate in undo and redo when editing is committed.
- e18d19d: Text can now carry its own outline, shadow and glow — the decorative half of a
  character format.

  `TextFormat` gained `outline`, `shadow` and `glow`, so every writer that takes
  one (`setShapeTextFormat`, `setShapeRunFormat`, `setTableCellTextFormat`, …)
  writes `<a:ln>` and `<a:effectLst>` into the run's `<a:rPr>`, and
  `getShapeRunFormat` / `getShapeRunFormatEffective` read them back. Passing
  `null` removes one effect and leaves the others alone. `GlowOptions` also
  gained the `opacity` its reader already reported, so a glow with `<a:alpha>`
  round-trips instead of losing it.

  The preview paints all three, and the editor exposes them in the text-format
  dialog and the properties panel in English and Japanese, including through the
  format painter. In the rasterised SVG text path only the outline is painted so
  far.

  `<Text>` in `@office-kit/pptx-dsl` keeps `shadow` and `glow` as the text box's
  own effects; per-run effects go through `paragraphs`.

- e18d19d: Selecting a chart shows the reference desktop app's Chart Design and Format tabs, and selecting a table shows Table Design and Layout.
- e18d19d: Add Clear text formatting to the preview's English and Japanese text and table-cell toolbars, with Ctrl/Command+Backslash for selected text. Clearing preserves hyperlinks and paragraph settings and supports Undo/Redo.

  The existing `setShapeTextFormat` and `setTableCellTextFormat` APIs now accept `{ reset: true }` to restore inherited run appearance before applying a new format, optionally limited to a character range.

- e18d19d: The editor follows the system's dark appearance, like the reference desktop app (Mac).
- e18d19d: Edit presentations directly in the development preview with English and Japanese controls. Canvas edits save with the project, survive server restarts and are included in CLI exports. Resolve conflicting source changes explicitly from the editor.
- e18d19d: The editor's Comments pane is now docked beside the slide instead of a modal dialog, so the ribbon, thumbnails and canvas stay usable while it is open, and it follows the selected slide.
- e9016f0: The preview no longer has a header bar: the editor fills the window. Agents opens from a new **Agents** button beside Comments and Share and docks as a task pane on the right, closed by default, in the editor's light or dark colors (including the terminal). **View ▸ Reading View** replaces the Preview button and opens the rendered viewer, whose status bar now holds Present, Presenter view, Download PPTX, Agents and **Normal** (Esc also returns to the editor). An agent's close button is now labelled "Close agent".
- e18d19d: Copy formatting from one object or text selection and paste it onto another —
  the paintbrush gesture, at last. Fill (solid, gradient, pattern or none),
  outline with its dash, cap, join, compound and arrowheads, shadow and glow, and
  the character and paragraph formatting of the text all travel; the text itself
  never does.

  Reachable from the right-click menu on an object, from the inline text toolbar
  while editing, and with Ctrl/Cmd+Alt+C and Ctrl/Cmd+Alt+V in both places, in
  English and Japanese. A picture fill, an inner shadow and effects the library
  cannot yet write are named in the toast instead of being silently dropped.

- 68a5940: `office-pptx dev` now merges TSX changes with canvas edits instead of asking you to choose between them. When the source is rewritten (for example by Claude Code) while you have saved or unsaved edits, changes to different shapes, slides, media and relationships are combined, and `build` / `exportDeck` export the merged deck. The editor only asks to **Keep my edits** or **Use source** when the same shape (or slide setting) was changed both in the editor and in the source, and it names what collided, e.g. "Slide 2: Title 1 was changed both here and in the source." Editor files saved by earlier versions are still read, but merge only after the next save.
- e18d19d: Add an optional proportional content fit when changing slide size, with English and Japanese page setup controls. Scale object coordinates, explicit text formatting, table dimensions and outlines, and center the original page on the new canvas. Invalid scaled dimensions leave the presentation unchanged.
- e18d19d: Move slide background editing into a dedicated Format Background pane with a close button and collapsible Fill controls. Open it from the Design tab or slide options; editing selected slides, applying backgrounds to all slides, and undo remain available.
- e18d19d: The Format pane starts closed, as in the reference desktop app, and opens from the object menu's Format Shape… / Format Picture… / Format Chart Area… / Format Video… / Size and Position…, from Format Background…, or from a Format Pane button. Once open it follows the selection: objects show their format, the slide shows Format Background (with the slide options below it).
- e18d19d: Preserve character formatting when copying, cutting and pasting selected text between editor shapes and table cells, including pending edits. Keep plain-text clipboard compatibility and support paste without formatting with Ctrl/Cmd+Shift+V.
- e18d19d: Add `isSlideBackgroundGraphicsHidden`, `setSlideBackgroundGraphicsHidden`, and `isSlideLayoutBackgroundGraphicsHidden` to inspect and control inherited decoration without deleting template content. Apply to All also copies the graphics visibility setting to slides and layouts. The preview honors both levels, and the editor's Format Background pane supports changing selected slides with undo and save/reload.
- e18d19d: Import common inline character styles when pasting HTML text into shapes or table cells, and include HTML formatting when copying text to other applications. Clipboard markup stays inert; unrecognized text layouts fall back to plain text.

  Preserve the reference desktop app's all-caps and small-caps formatting through clipboard round trips and inline text editing, including characters typed into an all-caps run.

- e18d19d: Add image crop aspect presets with locked resizing, a matching centered image frame, bilingual controls, and atomic undo.
- e18d19d: Restart kiosk presentations after the configured OOXML restart interval, returning to the first slide in the active show order and clearing the timer when presentation mode exits.

  Allow Escape to exit presenter view while a navigation button has keyboard focus.

- e18d19d: Edit a slide layout, not just apply one: `setSlideLayoutName`,
  `setSlideLayoutBackground`, `clearSlideLayoutBackground` and
  `setSlideLayoutPlaceholderBounds`.

  A layout could be read and applied but never changed, so a deck's shared design
  was fixed at whatever the template shipped. A layout handle now carries its own
  document and writes back into the layout part, and every slide on the layout
  follows the change — except where a slide set its own position or background,
  which still wins.

  Adding or removing a layout, adding a placeholder slot, and editing the slide
  master are still not supported.

  The editor exposes the settings through its properties panel and the
  Design ▸ Layout ribbon group, in English and Japanese.

- e18d19d: The New Slide and Layout menus show the layouts as a gallery of thumbnails, like the reference desktop app.
- e18d19d: The editor's Animations tab adds Preview, an Exit Effects menu and Animation Painter, and shows Path Animation and Trigger (disabled until the library writes them).
- e18d19d: The editor's Design tab now offers a Themes gallery, Variants, Colors and Fonts menus with the built-in theme color and font sets, Background Styles, Layout, Slide Size (Standard, Widescreen, Page Setup) and Design Suggestions.
- e18d19d: The editor gains Draw and Record tabs. Draw has pen, eraser and Lasso Select tools, a pen gallery (black and red pens, pencil, highlighter) with Add Pen, Ink to Shape (lines, triangles, rectangles and ellipses) and Draw with Trackpad; strokes are saved as editable freeforms. Record plays the show from the beginning or the current slide.
- e18d19d: The editor's Home tab now offers Paste, Cut, Copy and Format Painter; New Slide with a layout menu, Layout and Reset; Font; Paragraph with Bullets, Numbering and list-level buttons; Picture, Shapes and Text Box; Arrange, Quick Styles, Shape Fill and Shape Outline. As the window narrows, groups collapse into single buttons one at a time instead of all at once. Find and Replace moved to a new Edit menu, the Slide Show tab now precedes View, and the slide pane no longer shows its own add/duplicate/delete/move buttons (use the ribbon, context menu or keyboard).
- e18d19d: The editor's Home tab now has Section (add, rename and remove slide sections, shown as headings in the thumbnail pane), Columns, Text Direction and Align Text menus, and Add-ins, Designer and Convert to SmartArt in the reference desktop app's positions.
- e18d19d: The editor's Insert tab gains more commands. New: Screenshot, Action, Header & Footer (date, slide number and footer on the selected or all slides, with "Don't show on title slide"), Text Art, Symbol (inserts at the text cursor) and Video/Audio from a file. Commands the library cannot write yet (Cameo, Icons, 3D Models, SmartArt, Zoom, Object, Equation) are shown disabled with the reason.
- e18d19d: The editor's Insert, Design, Transitions, Animations, Slide Show and View tabs are reorganized without group captions, and a Review tab holds New Comment. Transitions and Animations show effect galleries with On Mouse Click / After / Apply To All and Start / Duration controls; Insert adds Date & Time and Slide Number; Slide Show adds Play from Start, Play from Current Slide, Presenter View and Hide Slide. The Thumbnails switch and Grid Options moved to the View menu.

  Fixed: after using a ribbon command while editing slide text, clicking a thumbnail or anything else that changes the selection now ends the text edit.
  Fixed: the Home tab no longer scrolls sideways in Japanese (or with wider system fonts) near a collapse width; it collapses the next group instead.

- e18d19d: The editor's Shape Format tab now offers Insert Shapes (with Edit Shape ▸ Change Shape), Shape Styles, Text Art Styles (Quick Styles, Text Fill, Text Outline and Text Effects), Alt Text, Arrange, Size and Format Pane.
- e18d19d: `SlideShowProperties` gains an optional `showMediaControls` (the reference desktop app's Show Media Controls, stored as `p14:showMediaCtrls`); `getSlideShowProperties` always reports it. The editor's Slide Show tab adds Rehearse Timings (time each slide while presenting, then keep the times as slide timings), Record, and the Use Timings, Play Narrations and Show Media Controls options; the commands that need a cloud service are shown disabled.
- e18d19d: The editor's status bar now shows the proofing language (the system language and region, e.g. "English (Japan)"), an Accessibility status that lists pictures, charts, tables and groups without alternative text and slides without a title (choosing one selects it), and Reading View and Slide Show buttons next to Normal and Slide Sorter. Reading View presents the current slide in the preview window without full screen.
- e18d19d: The editor's title bar has the reference desktop app's AutoSave switch (turn it off to stop saving the project until ⌘S), and the ribbon's tab row ends with Comments and Share buttons; Share ▸ Send a Copy downloads the `.pptx`.
- e18d19d: The editor's View tab now has Notes Page (the slide above its editable notes), Reading View and Slide Master, which opens a Slide Master tab for the current slide's layout; the layout commands left the end of the Design tab. Handout Master, Notes Master and Macros are shown disabled.
- e18d19d: Refreshed editor window chrome: the title bar and status bar use the neutral window color instead of the accent, the command palette appears as a Search box, ribbon tabs mark the active tab with bold text and an underline, selected slide thumbnails get a rounded accent ring, and the status bar shows "Slide 1 of 3" with Notes, Comments, view and zoom controls. Normal view now shows the notes pane by default with a "Click to add notes" prompt. Fixed: Undo while editing slide text no longer ends the edit when the notes pane is open, and the embedded editor's title bar no longer changes height (and shifts the canvas) when the save status changes.
- e18d19d: Preview selected audio and video directly in the editor, seek in quarter-second steps, and add or remove playback bookmarks. Bookmark times and media volume changes persist when saving the presentation.
- e18d19d: Say how a clip plays: `getShapeMediaPlayback` and `setShapeMediaPlayback`.

  A deck could embed a video or a sound but not state anything about playing it,
  so every clip waited for a click at the reference desktop app's default volume. The new pair
  reads and writes autoplay, loop, volume, mute, hide-when-stopped and (video
  only) full screen, from the clip's media time node. Omitted properties keep
  their current value.

  Trimming a clip is still not supported: the reference desktop app stores it in a 2010
  extension rather than in the core schema.

  The editor exposes the settings through its properties panel and command
  palette, in English and Japanese.

- e18d19d: Add a desktop-app-style Font dialog with font, character spacing, kerning, casing, baseline, and character-height controls for selected text, table cells, and outline text.
- e18d19d: Support selecting multiple slides with Shift, Ctrl/Command and Select All. Copy, cut, paste, duplicate, delete and reorder the selected slides as one undoable operation, with selection restored through history and Japanese/English selection counts.
- e18d19d: Format selected speaker-note text from the Home ribbon, preserving mixed formatting when editing and saving notes. Add APIs to read and update notes character formatting and change letter case.
- e18d19d: The object right-click menu now offers Edit Text, Group, Bring to Front and Send to Back submenus, Link, Edit Alt Text and New Comment. Duplicate, Delete and Copy/Paste formatting are no longer in it; their shortcuts (⌘D, Delete, ⌘⌥C / ⌘⌥V) still work.
- e18d19d: Edit links on images and other selected objects, and choose destinations within the presentation using the bilingual link dialog. Switching between web and slide destinations clears competing links and supports undo/redo and saving.
- e18d19d: Toggle text formatting in Outline View from its context menu without changing the presentation or adding an undo step.
- e18d19d: Allow `setShapeZIndex` to insert an ordered batch of sibling shapes in one update, preserving non-shape XML and rejecting mixed containers before mutation.

  Add a layer preview for Arrange > Reorder Overlapping Objects. Drag or use arrow keys to stage an order, cancel without changing the document, or confirm with one undoable change.

- e18d19d: Home ▸ Paste has the reference desktop app's options menu, including Keep Text Only.
- e18d19d: Render picture outlines along the image shape, including cropped masks and rotation. Add bilingual image-border color, width, and line-style controls with undo, redo, and saved persistence.
- e18d19d: Add setShapePreset and read picture presets through getShapePreset. Clip picture previews to their preset shape while preserving source-image crop and effects. Add bilingual image-shape controls for ellipse, rounded rectangle, triangle, diamond, pentagon, hexagon, star, and heart masks, with undo and saved persistence.
- e18d19d: Read and edit presentation drawing guides, guide visibility, grid spacing and grid snapping while preserving unrelated view settings.

  The editor adds grid settings, draggable and colored guides, guide/grid snapping, Slide Sorter, a View ribbon, resizable slide thumbnails and zoom controls. Edit speaker notes directly below the slide, with autosave and undo/redo. Display preferences remain separate from presentation edits.

  Keep View submenus inside the window, including when space at the right edge is limited.

- e18d19d: Play a slide's object animations in presentation mode. Clicks, the space bar, the arrow keys, the on-screen buttons and the presenter view all drive one click order: a slide's build is played through before the deck moves on, stepping back walks the build first, and a slide arrived at backwards is shown played out. An effect that runs with or after the previous one starts as the slide appears when nothing precedes it, a paragraph build reveals a paragraph per click, reduced motion keeps the order without the motion, and a slide that advances itself waits for the effects it has started.

  The presenter view runs the same player over its own copy of the slide, resumed at the same point in the same stop, so it shows what the audience can see — including an effect still fading in and one still waiting on its delay.

  Only what the deck states is played. An effect this library can read but not reproduce — an unknown preset, a target it does not model, a start condition or delay the tree does not state — keeps its place in the click order but is never approximated: nothing is chained onto an effect of unknown length, no click stop is invented, and everything such an effect touches is left exactly as it was drawn. A shape id that is drawn more than once is treated the same way, since nothing says which of them a timing tree means. The presentation controls say how many of a slide's animations are being left out, in English and Japanese.

- e18d19d: Play fade, push, wipe, cover, uncover, and zoom transitions in presentation mode with saved speed and direction settings. Respect reduced motion, cancel interrupted transitions, and start automatic advance after the effect finishes.
- e18d19d: The editor's title bar has the reference desktop app's icon Quick Access Toolbar (Save, Undo, Redo); New, Open and Download moved under its ⋯ menu.
- e18d19d: Add a bilingual find-and-replace dialog to the development preview, with match navigation, individual and bulk replacement, case sensitivity, current-slide scope, table cell support, keyboard shortcuts and undoable changes. Replacement values are literal text.

  Add an optional UTF-16 `range` to `setShapeText` and `setTableCellText` to replace an exact selection while retaining unaffected formatting, including when adjacent characters are identical. Invalid boundaries and split surrogate pairs are rejected.

- e18d19d: Keep browser recovery copies of committed editor changes and offer to restore or discard them after reloading. Isolate copies by project and editing session, preserve source conflict checks, and provide English and Japanese recovery controls.
- e18d19d: Display inline character formatting while editing text in shapes and table cells. Preserve selections across formatting controls, normalize native newlines, and support undo/redo for pending text and Japanese composition without losing rich clipboard styles.
- e18d19d: Apply background colors and images, layouts, skipped presentation state, and transitions to selected slides together in the visual editor. Mixed settings are indicated in English and Japanese, and each batch can be undone in one step.
- e18d19d: Reorder objects by dragging their names in the Selection Pane, with insertion feedback, edge scrolling, saved stacking order and undo support. Group children stay within their original group.

  Match the Mac Selection Pane with a single show/hide-all toggle and footer buttons for moving selected objects forward or backward.

  Use the Home Arrange menu for object stacking, grouping, alignment, rotation, flipping and opening the Selection Pane.

  Keep the alignment reference synchronized between the Arrange menu and the properties pane.

  Open and focus the numeric rotation control through Arrange > Rotate > More Rotation Options.

  Keep Enter-to-edit working from the table cell grid while preventing ordinary buttons from starting canvas text editing.

  Recognize Mac object arrangement shortcuts using physical keys: Option-Command-G for Group, Option-Shift-Command-G for Ungroup, Shift-Command-F/B for front/back, and Option-Shift-Command-F/B for one step forward/backward. Arrangement shortcuts also dismiss the open Arrange menu and no longer trigger Find.

  Restore a dissolved group with Arrange > Regroup or Option-Command-J, including edits made after ungrouping. Undo and redo preserve the former membership; opening a document starts a new regroup history.

  Distribute objects relative to the slide with equal outer margins, including one or two selected objects, from both Arrange surfaces.

- e18d19d: Open the editor's Selection Pane to select nested objects, rename them, and show or hide individual objects or the whole slide's objects with undo and autosave.

  Fix renaming and visibility changes on group shapes. Omit hidden objects and hidden group descendants from previews and canvas hit targets while preserving their editable content.

- e18d19d: Shapes opens the reference desktop app's shape gallery, and the chosen shape is drawn by dragging on the slide (or dropped at one inch on a click).
- e18d19d: Add a desktop-app-style Shape Quick Styles gallery to the editor Home and Shape Format ribbons, including theme and preset previews with Japanese localization.
- e18d19d: Add `isShapeLocked` and `setShapeLocked` for desktop-app-compatible object geometry locks, preserving text editing and unrelated drawing constraints.

  Add individual and all-object locks in the editor Selection Pane. Locked objects cannot be moved, resized, rotated, aligned or grouped; their lock state survives saving and undo/redo.

- e18d19d: Add a bilingual Set Up Show dialog to the development editor. Configure the show mode, slide range, looping, narration, animation and timings, with Cancel, validation, save/reload and one-step Undo support.

  Create, edit, copy, reorder and delete custom shows from the Slide Show ribbon. Arrange their slides in any order, including repeated slides, and preserve the sequence in the saved PPTX.

  Preview playback follows the saved slide range or custom sequence, looping, animation and timing settings. Window/kiosk behavior and narration playback are not yet applied in the preview.

- e18d19d: Add slide copy, cut and paste to the preview navigator's context menu and keyboard shortcuts, with independent snapshots, undo/redo and saved output. Allow importSlide to omit the target layout to preserve the source layout, master and theme, registering imported masters in the destination presentation.
- e18d19d: Right-clicking a slide offers Layout, Reset Slide, Grid and Guides, Format Background and New Comment.
- e18d19d: Slide numbers, as live fields rather than typed text.

  `setShapeTextField(shape, type, { text })` writes an `<a:fld>` — the slide's
  number, a date, a footer — replacing the shape's text body the way the reference desktop app
  writes one, and carrying the replaced text's formatting onto the field.
  `addSlidePlaceholder(slide, type)` restores a single slot the layout reserves
  (`sldNum`, `dt`, `ftr`, …), where `addMissingSlidePlaceholders` restores them
  all. `getPresentationFirstSlideNumber(pres)` reads `<p:presentation
firstSlideNum>`.

  The preview now substitutes `slidenum` fields with the slide's own position
  instead of whatever number the file last cached, counting from the deck's
  `firstSlideNum`. The editor gains a deck-wide slide-number switch in the slide
  panel and an "Insert field" command under Insert ▸ Text, both in English and
  Japanese.

- e18d19d: Set solid slide background transparency with the optional opacity argument to `setSlideBackground`. Background readers and previews retain the alpha channel, and the editor provides transparency controls for selected slides with Undo and save/reload support.
- e18d19d: Edit table cell margins on each side in points, apply them to selected ranges, and reset authored margins to the defaults. The bilingual preview controls show mixed values and preserve other margins when editing one side.
- e18d19d: Extend the existing paragraph formatting API to table cells, render their bullets and paragraph spacing, and add bilingual editor controls for individual cell paragraphs and selected cell ranges. Preserve rich text and support undo, redo, and save/reload.
- e18d19d: Navigate table cells with Tab and Shift+Tab and paste spreadsheet cells in the preview editor. Pasting expands unmerged tables, preserves cell formatting, supports quoted multiline values, and saves as one undoable edit. Merged-cell conflicts display an English or Japanese explanation before changing the table.
- e18d19d: Copy and cut selected table cell text as spreadsheet-compatible clipboard data, and paste into cells without entering text editing. Preserve quoted multiline values, apply overlapping pastes from a snapshot, and undo a cut or paste in one step.
- e18d19d: Add `setShapeTextLanguage(shape, lang)` and `getShapeTextLanguage(shape)` for the proofing language of a shape's text (`lang` on every run and paragraph end). The editor's Review tab gains Check Accessibility, Language, comment Delete/Previous/Next, Show Comments and Hide Ink; commands the library cannot support yet are shown disabled with the reason.
- e18d19d: The thumbnail right-click menu now offers New Slide, Duplicate Slide, Delete Slide, Add Section, Layout, Reset Slide, Format Background, New Comment and Hide Slide.
- e18d19d: Edit individual paragraphs or all paragraphs in a selected text shape through the preview’s Japanese/English formatting panel. Set alignment, bulleted and numbered lists, outline levels, line spacing, and paragraph spacing while preserving rich text and supporting undo/redo.
- e18d19d: Transitions can now carry a duration and a sound. `setSlideTransition` accepts `durationMs`, written as the reference desktop app (2010) writes it (`p14:dur` inside `mc:AlternateContent`, with the nearest `speed` as the fallback), and `getSlideTransition` reads it back, including from decks saved by the reference desktop app. New `setSlideTransitionSound` / `getSlideTransitionSound` embed a WAV sound or stop earlier sounds; changing the effect keeps the sound. The editor's Transitions tab gains Preview, Duration and Sound, and the preview plays fast/medium/slow transitions at the reference desktop app's 0.5/0.75/1 s.
- e18d19d: Crop images visually in the development preview using edge and corner handles, selection movement, and keyboard adjustments. Double-click a picture or use its image controls to open the bilingual crop dialog, preview the result, reset or cancel, and apply the crop in one undoable edit without changing the original image or frame.
- e18d19d: The editor's Text Art gallery now offers the full preset set: Shape Format ▸ Text Art Quick Styles and Insert ▸ Text Art show all twenty presets as styled "A" swatches, with the preset names as tooltips. Picking one writes the same fill, outline, shadow, glow, reflection and pattern settings the reference desktop app writes and replaces the previous Text Art style instead of layering on top of it. Theme-tinted colors stay linked to the deck's theme, and the Soft Bevel and Sharp Bevel presets write the reference desktop app's text 3-D settings. Clear Text Art also removes a bevel.

### Patch Changes

- e18d19d: Fix editor zoom buttons skipping the nearest 10% stop after entering a custom percentage. For example, zooming out from 123% now selects 120%.
- e18d19d: Choose Play Across Slides in the audio Playback ribbon and preserve the setting when saving or reopening the presentation.
- e18d19d: Keep each slide's solid, gradient, and pattern background settings when switching fill types during editing, including transparency and gradient stops.
- e18d19d: Preview background images with their stretch offsets, tile scale, alignment, offsets, and alternating reflections, including inherited backgrounds. Expose natural background image dimensions using embedded resolution and fill DPI, so tiling matches the image size instead of stretching it across the slide.
- e18d19d: Render background picture cropping for stretched and tiled images, including inherited backgrounds. Add `getSlideBackgroundImageCrop` to read the effective source rectangle without changing the deck.
- e18d19d: Keep Apply to All and Reset Background visible at the bottom of the Format Background pane while its fill settings scroll, including on shorter windows.
- e18d19d: Paste a clipboard image into a picture background from Format Background. Clipboard errors are shown in the pane, and changing the slide or document while an image is loading cancels the insertion.
- e18d19d: Edit picture background offsets, tile scale, alignment, and mirroring from Format Background. Layout changes preserve the image and support undo, save/reload, and remembered settings when switching between stretch and tile modes.
- e18d19d: Restore picture backgrounds when returning from another fill type during editing, preserving placement, transparency, cropping and imported effects. Add `copySlideBackground` to copy an effective background between slides or presentations while retaining its XML and referenced media.
- e18d19d: Read and edit slide background picture opacity, including inherited backgrounds. Format Background now provides a transparency slider and percentage field, and the preview displays the selected opacity.
- e18d19d: Add a Background Styles gallery to the editor's Design ribbon, with theme color previews, master-wide application, undo, and links to background formatting and reset. Expose `getSlideMasterBackgroundStyles` to read the available presets and master selection without modifying a presentation.
- e18d19d: Choose slide and shared-layout background colors from theme and standard palettes. Theme references survive saving, and mixed slide backgrounds no longer display one slide's color as though it applied to every selected slide.
- e18d19d: Localize preview navigation, presentation controls, status messages and accessible labels in Japanese and English. Preserve the editor language when switching views and when reloading directly into preview.
- e18d19d: Allow setting text formatting on blank autoshapes before typing. Preserve paragraph-end formatting through save/reload and subsequent text insertion, report zero paragraphs when a text-capable shape has no body, and show the preview formatting bar for empty shapes. Restore the Japanese text-format command label.
- e18d19d: Preserve character outlines, drop shadows, and glows while editing text inline, including theme colors and zoom-scaled effect geometry. Keep these visible effects in copied HTML as well.
- e18d19d: Edit slide background colors and images, reset backgrounds and select document
  layouts from a bilingual properties pane. Keep the controls synchronized with
  changes and history, and prevent delayed uploads from changing another slide.
- e18d19d: Allow setShapeTextFormat to target a UTF-16 text range while retaining surrounding
  run and paragraph properties. Reject invalid ranges without changing the document.

  Add a bilingual selected-text formatting bar to the development preview with
  bold, italic, underline, fonts, size and color, backed by persisted edits and undo.

  Keep property controls synchronized with canvas edits and preserve mixed text
  formatting when editing in the properties pane. Build the preview editor into
  the development package regardless of the working directory.

- e18d19d: Support `TextFormat.underlineColor` for independently colored underlines. Set it to `null` to follow the text color. Preserve the color when reading, editing, copying, and saving text, and expose underline color and Automatic in the editor's Font dialog.
- e18d19d: Expose the full DrawingML underline style set in the editor Font dialog, including localized Japanese labels.
- e18d19d: Support the reference desktop app's browse/window slideshow mode without entering fullscreen, including its configured slide navigation scrollbar.
- e18d19d: Preserve tab characters and paragraph default tab spacing in browser previews and direct text editing, including scaled text.
- e18d19d: Add bilingual alignment and equal-spacing controls to the development preview. Group and ungroup now act on the selected objects, retain stacking order, and restore selection through undo/redo. Grouped objects use a single selection frame so moving or selecting all does not also manipulate their children.
- e18d19d: Add bilingual speaker-note and slide-transition dialogs to the development preview, including automatic advance and applying a transition to all slides in one undo step. Correct speaker-note reads and updates to use the body placeholder without overwriting footers, read soft breaks and field results, and retain automatic-advance settings when no transition effect is selected.
- e18d19d: Restore in-progress canvas moves, resizes, and rotations when Escape is pressed or pointer/focus interaction is interrupted, preserving selection and redo history. Cancel range selection with Escape and avoid applying geometry edits for click jitter.
- e18d19d: Target canvas context menus at the clicked shape or table cell, preserve selected ranges, clear shape selection on empty space, and prevent right-button movement, resizing, and rotation. Keep the native text context menu while editing.
- e18d19d: Support the reference desktop app's Top, Middle and Bottom Centered text anchors through the existing `setShapeTextAnchor` API's `centered` option. Resolve inherited centering through `getShapeBodyPrEffective`, preserve paragraph alignment, and expose all six anchor choices in the editor with preview and text-editing support.

  Preview integrations can use `shapeTextAnchorOffset` to position editable text consistently with the rendered text block.

- e18d19d: Match the reference desktop app's detailed character spacing flow by opening the Font dialog with spacing and kerning controls.
- e18d19d: Keep chart titles, axes, data labels, trendline names, and legends readable when a chart or parent group is flipped. Preserve label rotation and its side of the anchor.
- e18d19d: Add a bilingual skip-during-presentation control and skipped-slide indicators in the editor and preview. Keep skipped slides editable while omitting them from keyboard, button, click, automatic, and presenter navigation. Disable presentation when every slide is skipped.
- e18d19d: Make slide insertion, duplication, deletion and reordering operate directly on the selected slide in the preview. New and duplicated slides appear beside the active slide. Add bilingual navigator controls, keyboard navigation and reordering, and slide context menus. Undo and redo retain the corresponding slide selection.
- e18d19d: Fall back to plain text when pasted formatting contains invalid kerning or baseline values, preventing invalid presentation output.
- e18d19d: Collapse the editor ribbon to leave more room for slides, and reopen commands by selecting a tab. Keep ribbon tabs horizontally scrollable in narrow windows.
- e18d19d: Add comment replies through `addSlideComment({ replyTo })` and `getCommentParent`, preserving the reference desktop app's p15 threading extensions through save and reload. Removing a comment also removes its descendant replies.

  The preview comments dialog supports creating and editing replies in English and Japanese, shows each reply's parent, and deletes threads with undo/redo support.

- e18d19d: Keep replies together with their parent comments in the preview, including after save and reload. Focus the new author field when adding a comment or reply, and keep keyboard focus in the dialog after deleting a thread. Imported orphaned or cyclic threads remain visible.
- e18d19d: Review and edit comments across the presentation in one dialog. The bilingual slide selector shows titles and comment counts, preserves drafts while switching slides, and applies all edits in one undoable transaction. Validate unfinished comments across slides and ignore untouched new rows.
- e18d19d: Open the Arrange commands directly from the compact Home ribbon, removing the extra Arrange submenu click.
- e18d19d: The embedded development editor now uses a single compact, non-wrapping command header and denser save status, keeping the canvas visible without removing editor commands.
- e18d19d: Concatenate text from multiple shapes with `setShapeParagraphs(target, { sources })`, retaining paragraph formatting and hyperlinks. In Outline View, demote a slide title into the previous slide's body with Tab or the context menu, with Undo support. Slides with additional objects or no destination body placeholder report an error before changing their content.
- e18d19d: Copy existing formatted text into a shape with `setShapeParagraphs(shape, { source, range })`. The optional UTF-16 range retains paragraph properties, fields, run formatting and link relationships while keeping the destination text body's settings.

  Splitting an outline title now retains hyperlinks and paragraph formatting in the new title.

- e18d19d: `getShapeCustomGeometry` now reports the `<a:rect>` a custom-geometry shape
  states for its text, as `textRect`, with its guide formulas already evaluated.

  The preview lays a custGeom shape's text in that rectangle instead of the whole
  bounding box, through the new `shapeCustomTextRect` in `@office-kit/pptx-preview`,
  and inline editing in the dev editor puts the caret in the same place. A custom
  shape that states no rectangle still gets its whole box — a preset's
  approximated region is never substituted for a shape that describes itself.

- e18d19d: Reset slideshow playback to all slides when its selected custom show is deleted, while preserving the other slideshow settings.
- e18d19d: Follow custom slide show links during playback and in presenter view, with the option to return to the calling slide when the linked show ends.
- e18d19d: Keep custom left, center, right, and decimal tab stops positioned when editing slide text directly, including after typing and undoing text changes.
- e18d19d: Edit individual objects inside groups in the development preview, with English and Japanese group navigation, transformed pointer/keyboard movement, and undoable deletion and duplication.

  Allow `copyShape` to preserve ancestor group transforms with `preserveGroupTransform`. Deletion and z-order operations now handle nested shapes in their owning group and preserve trailing extension metadata.

- e18d19d: Add Outline View for editing slide title and body placeholders. Text edits preserve existing formatting and save automatically; Enter in a title creates a new slide with the remaining title text. Resize the outline pane and undo edits through the existing document history.

  Splitting an outline title also moves its body text to the new slide while preserving formatting.

  Tab in an outline body demotes the selected paragraphs while preserving text and run formatting, with Undo and saved paragraph levels.

  Use Command-4 to open Outline View and Command-1 to return to Normal, including while editing text. Pending outline edits are saved when switching views.

- e18d19d: Preserve custom tab alignment while editing small-cap text or text with expanded and condensed character spacing.
- e18d19d: Keep custom tab alignment accurate while editing text with all-caps formatting or disabled kerning.
- e18d19d: Show the displayed font size in text formatting controls when text has no explicit or inherited size, instead of leaving the size field blank.
- e18d19d: Preserve double strikethrough while editing text and copying or pasting HTML, including text with a separately styled underline.
- e18d19d: A canvas edit that fails partway is now undone completely instead of leaving
  half-applied changes (such as an unused image) in the saved deck. Project Undo
  and Redo no longer leave a source file partially written when a write fails,
  and a refused text edit keeps its original reason when the source also changed.

  Undo and Redo now replace a project file by writing a temporary file beside it
  and renaming it, so a file inside a read-only directory can no longer be
  restored in place.

- b74d3f7: The editor now uses the office-kit palette: a rose accent and cool neutral grays in place of the previous orange and warm grays. Selection highlights, focus rings, the lasso and draw previews, and icon accents follow the accent color in both light and dark themes.
- e18d19d: Expose inherited bullet color, size, and font details through `getParagraphPropertiesEffective`.

  Preserve bullet formatting inherited from layouts and masters in previews and while editing text, including explicit follow-text overrides.

- e18d19d: Read stored end-of-paragraph formatting when the caret is in an empty paragraph, so inline formatting indicators and toggles use the intended typing style in text boxes and table cells.
- e18d19d: Allow `setShapeTextFormat` and `setTableCellTextFormat` to target a paragraph end mark with `paragraphEnd`, preserving existing runs and paragraph properties.

  Fix Increase/Decrease Font Size leaving empty paragraphs unchanged when formatting an entire text box or table cell.

- e18d19d: Allow paragraph formatting before typing in an empty text box. The preview keeps the chosen alignment and list style when entering text, adding paragraphs, and undoing or redoing the first input.
- e18d19d: Keep the last text input when Escape ends editing, matching the reference desktop app (Mac).
- e18d19d: Preserve exact trim and fade positions when opening media whose duration is not a multiple of 50 milliseconds. Keep keyboard and drag adjustments in 50 millisecond increments without rounding saved values or losing the clip endpoint.
- e18d19d: Restore inline text focus and selection after cancelling the Font dialog opened with Cmd/Ctrl+T.
- e18d19d: Match the reference desktop app's 0.1-point increments and rounding in the Character Spacing dialog, while preserving imported spacing precision when changing other font settings.
- e18d19d: Fix upward keyboard navigation from the font search field and keep the notes pane resize limit consistent with its visible height.
- e18d19d: Match the reference desktop app's font offset editor by requiring a percent value and validating the supported range for superscript and subscript offsets.

  Preserve mixed superscript and subscript offsets when an invalid entry is discarded and another font setting is applied.

- e18d19d: Choose a font size from the Home ribbon's size menu, including the sizes available in the reference desktop app (Mac). Custom sizes remain editable in the adjacent field.
- e18d19d: Keep format painter shortcuts aligned with the native text selection when copying or applying formatting.
- e18d19d: Preserve the shadow anchor and rotation setting when copying a shape's formatting.
- e18d19d: Close the Format Shape pane to expand the editing area, and reopen it from the object's Format Shape or Size and Position context commands. More Rotation Options also reopens the pane and focuses Rotation.
- e18d19d: Organize shape formatting into Fill & Line, Effects, and Size & Properties tabs. Retain field state when switching tabs and open the size tab from More Rotation Options.
- e18d19d: Add bilingual image upload and replacement dialogs, crop and appearance controls,
  and alternative text editing to the development preview. Preserve picture geometry
  and crop during replacement and support undo and persisted reloads.

  Correct signed image contrast rendering so zero is neutral and negative contrast
  reduces color separation without inverting the picture.

- e18d19d: Honor the saved full-screen video setting during slide shows. Keep playback at the same position when pausing, and restore the slide and keyboard focus when the video finishes.
- e18d19d: Preserve untouched coordinate and size precision when editing one geometry field in the preview. Reject empty, negative-size and out-of-range values before mutation, restore invalid inputs, and support fractional positions and rotations with an accessible bilingual rotation label.
- e18d19d: Choose from eight linear gradient directions in the Format Shape pane. Direction changes preserve gradient stops, support Undo/Redo, and use the same scaled angle settings as the reference desktop app (Mac).
- e18d19d: Keep linear gradients fixed to the slide when Rotate with shape is disabled, including on wide or tall rotated shapes.
- e18d19d: Preserve gradient stop transparency when reading shapes and expose theme-resolved stop colors. Shape previews now render the brightness and transparency saved by the reference desktop app for linear and radial gradient stops.
- e18d19d: Allow gradient stops to specify brightness and opacity, and gradients to preserve scaling and rotation settings. Reject invalid stop edits without discarding the shape's existing fill.

  Edit gradient stop color, position, brightness and transparency in the Format Shape pane, with add/remove controls and Undo support. Drag stops directly, commit each drag as one undoable edit, and cancel an in-progress drag with Escape.

- e18d19d: Choose gradient stop colors from the presentation theme or standard color palette. Theme choices remain linked to the theme after saving; custom colors remain available.
- e18d19d: Show gradient stop transparency in the editor's gradient preview track, including after undo and reopening a project.
- e18d19d: Match the reference desktop app when changing gradient types: Radial and Rectangular start from the bottom-right corner, Path starts from the center, and Linear starts at 45 degrees with scaling enabled. Preserve existing color stops and rotation settings.
- e18d19d: Add image-fill placement readers and setters for tile alignment, scale, offsets, mirroring, stretch offsets and rotation with the shape. Placement changes preserve the embedded image, crop and effects. Preview now renders stretch offsets and clips the image to the shape.
- e18d19d: Keep shape text input centered and match the preview's text orientation and scale inside rotated, reflected and resized groups. Honor the text rotation of vertically flipped shapes while editing.
- e18d19d: Keep shape text readable inside flipped groups, including nested rotated groups, while preserving the transformed text position and baseline direction.
- e18d19d: Snap, align and evenly distribute objects by their visible edges, including rotated shapes inside transformed groups. Keep smart guides in slide coordinates and preserve screen-space snap distances through group scaling.
- e18d19d: Show the shape's default text alignment in the Home ribbon, including centered text in shapes with no explicit paragraph alignment.
- e18d19d: Adjust character spacing from the Home ribbon using the reference desktop app's five presets or a custom expanded or condensed value. Preserve unmodified formatting when editing mixed selections, with undo and redo support.
- e18d19d: Use the Home ribbon to change fonts, size, emphasis, color, and highlighting while editing text or selecting text boxes. Switching ribbon tabs preserves the active text edit and its selected range.
- e18d19d: Add Increase and Decrease Font Size controls to Home, preserving different sizes within selected text and using the reference desktop app's font-size steps.
- e18d19d: Set line spacing from the Home ribbon using the reference desktop app's spacing presets. Paragraph options let you change alignment, indentation, paragraph spacing, and exact or multiple line spacing together, preserving the selected text range and undoing the changes in one step.
- e18d19d: Align paragraphs directly from the Home ribbon. Alignment buttons show the current selection's alignment and apply to selected paragraphs during text editing or all text in selected text boxes.
- e18d19d: Keep playback controls horizontal and clickable when editing rotated audio or video.
- e18d19d: Preserve font-relative character spacing when pasting formatted HTML into slide text or table cells.
- e18d19d: Image brightness and contrast now work on image-filled shapes. The editor preserves these corrections when switching fill types and restoring the image.

  Image corrections, opacity, and DrawingML color transforms now read both fixed-point and percent-suffixed values correctly. Hue offsets use angle units, preventing incorrect colors in imported presentations.

- e18d19d: Recognize `true` and `false` as well as numeric DrawingML flip flags in imported presentations. Flipped shapes and groups now retain their orientation in the preview and when ungrouped, including after saving and reloading.
- e18d19d: Show resize and rotation handles on each selected object. Resizing multiple objects keeps each object's opposite handle fixed, and rotating them keeps their individual centers fixed, matching the reference desktop app.
- e18d19d: Show inherited slide background colors and gradients in the editor. Editing an inherited gradient changes only the selected slides, preserving their layout and master backgrounds; resetting restores the inherited fill.
- e18d19d: Show and edit inherited placeholder positions and sizes in the preview property panel. Preserve untouched resolved coordinates and support aspect-ratio locking and undo back to layout inheritance.
- e18d19d: Allow effective run formatting to resolve layout and master inheritance from an original shape while reading a detached editing preview. Show inherited character styles during inline shape editing, including pending text and paragraph changes, without baking those styles into saved text or clipboard data.
- e18d19d: Resolve inherited text-box character styles in the inline formatting toolbar, including pending edits, so size indicators and formatting toggles agree with rendered text.
- e18d19d: Resolve inherited text autofit and columns through `getShapeBodyPrEffective`, and use them consistently in previews and editing controls. Allow `setShapeTextColumns` to author one column explicitly, overriding inherited columns; invalid column settings leave the previous values intact.
- 7680d70: `init` now scaffolds `theme.ts` with the office-kit rose accent (`#D6336C`) instead of `#E5481F`.
- e18d19d: Keep text at the same layout scale in the preview and editing view so entering text editing no longer shifts wrapped lines at fractional zoom levels. Preserve the size, spacing, and font of large or explicitly styled bullets while editing.
- e18d19d: Preserve explicit bullet color, size, and font while editing text in the preview.
- e18d19d: Keep undo and redo working across text input and formatting changes while editing shapes and table cells. Preserve pending text redo when stepping back through formatting history.
- e18d19d: Change list nesting while editing with Tab and Shift+Tab, or use Ctrl/Cmd+[ and Ctrl/Cmd+] for paragraph levels, including in table cells. Mixed paragraph selections retain their relative nesting within the supported levels. Preserve table-cell Tab navigation and ignore these shortcuts during Japanese composition. Explain the shortcuts in the English and Japanese list-level controls.
- e18d19d: Show bullet and numbered-list markers during inline editing, including nested numbering, without changing copied text or selection offsets. Share the preview renderer's numbering through `paragraphNumberLabels` so editing surfaces use the same counter and restart rules. Build inline paragraph text in one pass instead of copying the whole shape for every paragraph.
- e18d19d: Allow effective paragraph properties to retain original placeholder and slide inheritance when reading a detached shape preview. Display paragraph alignment, line spacing, spacing before and after, indentation and text direction during inline editing, with canvas zoom applied to absolute dimensions.
- e18d19d: Keep superscript and subscript text at its preview size and position when entering inline editing, including custom offsets and mixed font sizes.
- e18d19d: Keep shape and table-cell text margins and vertical alignment while editing text in the preview.
- e18d19d: Scale inline text and editing padding with canvas zoom while preserving text selection and original clipboard font sizes.
- e18d19d: Inline editing now reads vertical text the way the preview paints it, and
  shrinks autofit text by the same factor — a `<a:normAutofit/>` title no longer
  jumps back to its authored size the moment the caret appears.

  Two new exports carry the shared rules: `verticalTextStyle` / `textColumnsStyle`
  turn `<a:bodyPr vert=… numCol=… spcCol=…>` into the CSS both surfaces use, and
  `shapeAutoFitScale` reports the shrink factor the renderer applies to a shape,
  for the box it is actually laid out in.

- e18d19d: Show the crop panel heading in Japanese when Japanese is selected.
- e18d19d: Translate the color label in the Japanese paint controls and drawing-guide menu.
- e18d19d: Apply OOXML kerning thresholds consistently in SVG and HTML preview layout, fontkit and browser measurement, and inline text editing.
- e18d19d: Support opening context menus with Shift+F10 or the context-menu key. Move among enabled actions with arrows, Home and End, activate with Enter, and return focus on Escape. Keep menu key presses from changing the underlying slide or cell selection.
- e18d19d: Keep kiosk slideshows looping and prevent ordinary clicks or navigation keys from advancing slides. Explicit slide links and Escape remain available.
- e18d19d: Preserve the font size of leading empty lines in shapes and table cells, preventing text from jumping when editing begins.
- e18d19d: Prevent text from moving down when entering editing if a line break has a larger font size than the surrounding text, while retaining its formatting for selection and copying.
- e18d19d: Resolve line-break formatting with `{ breakIndex }` in getShapeRunFormatEffective and getTableCellRunFormatEffective. Preserve inherited font, size, and emphasis when selecting, editing, or copying line breaks.
- e18d19d: Keep character reflection and inner-shadow effects visible while shape or table-cell text is being edited, including tables in scaled groups, and align rotated editing text around the text body's inner rectangle.

  Add `renderTextEffectsSvg` to render character reflection and inner-shadow overlays in local text-box coordinates, including text-body rotation and vertical layout.

- e18d19d: Add an English/Japanese aspect-ratio lock for numeric shape sizing in the preview. Update paired dimensions in one undo step while preserving position and rotation, and reject proportional sizes outside the document format's limits.
- e18d19d: Prevent locked objects from being moved or resized through the editor's numeric property fields, and disable their rotation, aspect-ratio and flip controls.
- e18d19d: Show a contextual Playback tab when selecting audio or video. Change automatic or click playback, automatic start delay, volume, mute, looping, visibility, and video full-screen playback directly, with undo and saved settings.
- e18d19d: Play embedded audio and video in the local slide show using the saved autoplay, volume, mute, and loop settings. Stop playback when leaving a slide and show a retry control when the browser cannot play the media.
- e18d19d: Add Rewind After Playing to the media playback ribbon and return audio and video to the beginning after playback when enabled. Preserve the setting when saving and reopening presentations.
- e18d19d: Preserve automatic media start delays through `MediaPlayback.delayMs`. Presentation preview now waits for the saved delay before playing audio or video and cancels pending playback when leaving a slide or starting media manually.
- e18d19d: Edit audio and video trim ranges and fade durations, and apply them during slide-show playback.
- e18d19d: Edit media trim boundaries, fades, and playback position on a shared timeline with a waveform decoded from the embedded audio.
- e18d19d: Match the reference desktop app's media volume menu with Low, Medium, High, and Mute choices.
- e18d19d: Align the Home ribbon's compact groups with the reference desktop app by putting slide layout actions in Slides and separating Insert actions from Drawing and Arrange controls.
- e18d19d: Show the gradient format controls for selections with different stops. Keep rotation editable without replacing individual stop sets, and show unavailable stop and direction controls as blank and disabled.
- e18d19d: Show mixed rotation values in the preview property panel and apply entered angles to every selected object in one undo step. Preserve object positions, sizes and unselected shapes, and avoid repeated shape-list scans when resolving selections.
- e18d19d: Apply numeric position and size edits to every selected object in the preview. Show mixed values, preserve each object's own aspect ratio when locked, and reject oversized proportional changes before modifying any object. Support one-step undo and redo with English and Japanese controls.
- e18d19d: Edit matching gradient stops across selected shapes in one undo step while retaining each shape's angle, scaling and rotation settings. Show mixed angle and rotation values in the format pane.
- e18d19d: Move all selected objects together with Bring to Front, Send to Back, Bring Forward and Send Backward, preserving their relative stacking order and supporting a single undo step in the editor.

  The four existing stacking APIs also accept an array of sibling shapes. Selections from different parent containers are rejected before any shape is moved.

- e18d19d: Collapse the Home ribbon's font controls into a menu in narrow editor windows, keeping font selection available without taking space from the other groups.
- e18d19d: Choose line start and end arrows from visual galleries with six arrow types and nine sizes. Editing multiple lines preserves each endpoint's other settings and supports undo and saved reloads.
- e18d19d: Switch between no fill, solid fill, and gradient fill directly in the Format Shape pane. Show the selected fill's controls, restore its previous settings when switching back, and undo a fill-type change in one step. New gradients use the reference desktop app's four-stop defaults.
- e18d19d: Match the reference desktop app's (Mac) inline media toolbar with a filled seek track, centered playback controls, and a mute button. Keep volume adjustment in the Playback ribbon instead of showing an extra slider beside the media.
- e18d19d: Separate Fill and Line into collapsible sections, and edit compound lines, line caps and joins directly in the Line section. Mixed selections preserve distinct values until edited, with saved changes participating in undo and redo.
- e18d19d: Convert the reference desktop app's native Play in Background timing tree when changing an audio clip between automatic and click playback, preserving the existing timing IDs and media settings.
- e18d19d: Choose pattern fills from the Format Shape pane, with 48 visual presets and foreground/background color controls. Pattern changes support multiple selected shapes, Undo/Redo, saved reloads, and restoring the previous pattern after switching fill types.

  Starting a new presentation now clears remembered fill settings from the previous document.

- e18d19d: Show size and position in collapsible sections with centimeter units, and add percentage scaling from the selected objects' initial dimensions. Keep mixed selections, object locks, inherited placeholder geometry and undo working across these controls.
- e18d19d: Add a collapsible Text Box section with text direction, autofit, centimeter margins, wrapping and a staged Columns dialog. Apply edits to selected text shapes together with saved undo support.
- e18d19d: Group sibling shapes and ungroup groups inside an existing group without removing the outer group or changing its transform. Grouping preserves the members' stacking order regardless of selection order. Preview group commands support nested selections with Undo/Redo and reject selections spanning different parents.
- e18d19d: Show and edit playback settings for audio and video nested inside animation timing groups. Preserve enclosing start conditions when changing volume, looping, or video display settings, and account for parent delays when reading and editing simple automatic playback. Reject unsupported start-condition edits before changing the document. Also reject non-finite volume values instead of writing invalid XML.
- e18d19d: Add a bilingual layout picker for new slides in the development preview. Insert the chosen layout and its editable placeholders immediately after the current slide, with undo/redo and saved edits.
- e18d19d: Match the reference desktop app's generic New Slide command by preserving the current content layout and advancing title slides to the same master's content layout.
- e18d19d: Save speaker notes after IME text is confirmed, even when no further keystroke or focus change follows.
- e18d19d: Add direct font, size, emphasis, color, highlight, and clear-format controls to the preview property panel when selecting objects containing text. Apply formatting to all selected objects together, with English and Japanese scope labels and undo support.
- e18d19d: Read and update paragraph levels across a UTF-16 text selection with `getParagraphLevel` and `setParagraphLevel`. Relative updates preserve differences between levels and keep text, formatting and links intact. Indenting a long outline selection now updates its text body once instead of once per paragraph.
- e18d19d: Persist outline collapse with `getCollapsedOutlineSlides` and `setSlideOutlineCollapsed`, preserving the reference desktop app's per-slide view state. Deleting a slide removes its outline reference.

  In the editor, double-click a slide icon in Outline View to collapse or expand its body. The right-click menu also collapses or expands the selected slides or the whole outline. Each change supports Undo and survives saving and reopening.

- e18d19d: Fix outline slide merging doing nothing after confirmation when recently typed text was still being saved.
- e18d19d: Copy, cut and paste formatted outline text from the context menu. Delayed clipboard permission responses no longer risk changing a different selection or a closed editor.
- e18d19d: Make Outline View text selections continue across slide titles and body placeholders, preserving rich text on copy and grouping cross-field edits into one undoable transaction.
- e18d19d: Merge slides when deleting or replacing an outline range between slide titles, retaining the final slide's body text and links. Preserve ribbon focus when restoring the editing caret after the replacement.
- e18d19d: Keep the remaining body text attached to the title when cutting or typing over an outline selection spanning both fields.
- e18d19d: Allow `addSlidePlaceholder` to inherit a placeholder directly from the slide master without changing the slide layout.

  The editor now confirms deletion of additional objects when demoting an outline title, and restores a missing body placeholder even on Title Only slides. Cancel keeps the document unchanged; Undo restores the deleted slide and its objects.

- e18d19d: Keep inherited outline text formatting stable while typing in presentations that use multiple slide masters.
- e18d19d: Restore outline text focus and selection after cancelling the Font dialog opened with Cmd/Ctrl+T.
- e18d19d: Preserve the remaining body text when replacing an outline title-to-body selection with direct text input or an IME composition. Committing a composition is one undoable edit; cancelling it leaves the text unchanged.
- e18d19d: Outline View now moves the caret between adjacent title and body fields with Arrow Up and Arrow Down at text boundaries, preserving the existing pending-edit and undo behavior.
- e18d19d: Move selected slides up or down directly from the outline context menu, with undo and redo support.
- e18d19d: Keep surviving title paragraphs in the title when deleting or replacing an outline selection that extends into the slide body.
- e18d19d: Show paragraph indentation, bullets and numbering while editing outlines. Preserve selection, empty paragraphs and pasted text through editing, saving and undo. Collapse and expand slides directly from the outline text context menu.
- e18d19d: Move selected outline body paragraphs up or down from the context menu, retaining paragraph formatting and hyperlinks through save and Undo. The existing `setShapeParagraphs` API now accepts ordered source ranges for a single destination.
- e18d19d: Preserve the remaining body text when pasting over an outline selection spanning the title and body, including pasted formatting and paragraph breaks.
- e18d19d: Promote outline body paragraphs with Shift+Tab or the outline text menu. Nested paragraphs move up one level; selected root paragraphs become separate slide titles, retaining following body text, formatting and hyperlinks. Tab and Demote move paragraphs down one level. These operations support Undo and Redo.
- e18d19d: Fix extending outline text selections upward across slides with Shift+Up. Keep the original selection anchor when selecting backwards, including consecutive key presses, so copying and editing apply to the whole selected range.
- e18d19d: Apply ribbon formatting to selected outline text across multiple fields with a single undo, and keep ribbon focus when an outline edit finishes rendering.
- e18d19d: Avoid an outline selection error when Undo or a view change removes the focused text field.
- e18d19d: Match the reference desktop app's (Mac) outline title editing: Shift+Enter splits the title into a new slide, preserving the following text, formatting and body content.
- e18d19d: Add, duplicate and delete slides from outline text context menus. New slides inserted from the outline retain the current slide layout, and all three operations support undo.
- e18d19d: Reorder slides by dragging their icons in Outline View, including multiple selected slides. An insertion line marks the destination; the move preserves slide order and supports Undo.
- e18d19d: Preserve character formatting when copying, cutting and pasting text in Outline View, including pending edits and formatted HTML from other applications.
- e18d19d: Join the remaining body text onto the title when deleting an outline selection across their boundary, preserving subsequent body paragraphs and undo history.
- e18d19d: Fix Enter and Shift+Enter across an outline title and its body to split the slide and move the remaining body text into the new title. Preserve text links and restore the original slide with Undo.

  Fix Enter and Shift+Enter across adjacent slide titles in the outline to retain both slides and their remaining text instead of inserting a line break into the first title.

- e18d19d: Move selected outline title text up or down across adjacent body paragraphs, preserving paragraph levels, formatting and hyperlinks. Undo restores the previous slide bodies.
- e18d19d: Show the preview's default centered alignment for autoshapes in the inline toolbar and paragraph panel, while retaining left-aligned text box and table defaults and honoring explicit alignment.
- e18d19d: Fix existing text in shapes and table cells incorrectly adopting the font, size, or emphasis saved for newly inserted text at the end of a paragraph.
- e18d19d: Add `setParagraphIndent` to edit left, right, and first-line indentation in text boxes and table cells, or remove individual overrides to restore inherited values.

  Expose paragraph indentation in the editor command palette.

- e18d19d: Add the Tabs dialog inside Paragraph settings, with custom positions, alignment, default spacing, and individual or complete clearing. Changes remain pending until Paragraph is confirmed, support undo, and preserve unrelated tab stops when editing multiple paragraphs.
- e18d19d: Read and update paragraph Asian line breaking, Latin word wrapping, hanging punctuation, and font alignment with `setParagraphTypography`, including inherited settings and removal of local overrides.

  Add the Paragraph dialog's Line Breaks and Alignment tab. Apply changes to selected paragraphs or table cells together, preserve mixed values, and support undo and redo. These settings are saved in exported presentation files; preview rendering does not yet reproduce all typography rules.

- e18d19d: Choose pattern foreground and background colors from theme and standard palettes. Base theme colors stay linked when switching fill types and after saving.

  `getShapePatternFill` accepts `preserveTheme: true` to return untransformed theme references; its default continues to return resolved RGB colors.

- e18d19d: Read image crop offsets correctly when imported presentations store them as percentages. Include selected text and theme fonts in the editor's font picker, including custom font names.
- e18d19d: Read and update shape aspect-ratio constraints with `isShapeAspectRatioLocked` and `setShapeAspectRatioLocked`. The editor now preserves the saved Lock aspect ratio setting when resizing videos and using the size pane.
- e18d19d: Edit picture crop dimensions and position in centimeters, with separate controls for the source picture and crop frame. Reset restores the full picture without stretching it, and changes support undo.
- e18d19d: Insert a picture fill from the clipboard, retaining transparency and placement settings with undo and redo support.
- e18d19d: Keep the picture fill radio button selected when restoring a remembered image fill in the format pane.
- e18d19d: Preserve image detail when applying Washout and match the reference desktop app's brightness and contrast calculation in slide previews, live video, and correction thumbnails.
- e18d19d: Add a Play in Background command that enables automatic, looping audio across slides while preserving volume and rewind settings. Keep audio hidden during playback when Hide During Show is enabled.
- e18d19d: Add the reference desktop app's Top Left Corner and Center position references. Changing the reference retains the entered distance and moves selected objects; Undo restores geometry while retaining the chosen reference.
- e18d19d: Preserve the actual insertion and deletion positions when editing repeated text in shapes and table cells, keeping neighboring run formats intact. Keep emoji replacement boundaries valid during incremental editing.
- e18d19d: Presenter view no longer ends on its own when a slide show was started and
  closed with Escape just before opening it. Bottom-to-top (`vert270`) table cell
  text no longer jumps sideways when editing a cell with uneven left and right
  margins.
- e18d19d: Preview embedded videos in presenter view and control audience playback with play, pause, and seeking. Keep presenter media muted to avoid duplicate audio, and stop it when leaving the slide or ending the show.
- e18d19d: Follow internal slide links from the presenter window, synchronizing the audience slide and speaker notes. Support mouse and Enter activation and reject invalid destination indices.
- e18d19d: Preserve advanced underline styles when formatted text is copied from the editor and pasted back through HTML-only clipboard paths.
- e18d19d: Allow `getShapeImageDuotone` to preserve theme references and color transforms with `resolveColors: false`. Keep two-color image recoloring when switching a shape's fill away from an image and back in the editor.
- e18d19d: Preserve mixed text formatting during inline preview edits. `setShapeText` now accepts `preserveFormatting` for incremental changes, retaining unchanged runs, paragraph properties, fields and hyperlinks.
- e18d19d: Keep existing outline color, theme references, and opacity when changing only line width. Image border width and dash controls preserve the existing color instead of replacing it with an opaque resolved color.
- e18d19d: Preserve existing speaker-note run formatting and fields when editing notes in the editor. `setSlideNotes` now accepts optional `range` and `preserveFormatting` options for targeted text edits; its default replacement behavior is unchanged.
- e18d19d: Share the preview's preset text rectangle calculation through `resolveTextBodyRect`. Inline editing now retains the preview's constrained text region in triangles, diamonds, pentagons, stars and double arrows, and respects default centered autoshape paragraphs.
- e18d19d: Play horizontal and vertical blinds and comb slide transitions in presentation mode. Respect saved direction and speed, reduced motion and interrupted-transition cleanup.
- e18d19d: Play horizontal and vertical checkerboard slide transitions in presentation mode, respecting saved speed, reduced motion and interrupted-transition cleanup.
- e18d19d: Honor imported XML boolean spellings for through-black transitions and click advancement. Play cut-through-black transitions in presentation mode, respecting transition speed, reduced motion, and cleanup when interrupted or exited.
- e18d19d: Play dissolve slide transitions in presentation mode at the selected speed.
- e18d19d: Follow internal slide links in preview and presentation playback, including keyboard activation. Clicking internal or external links no longer also triggers ordinary slide advancement.
- e18d19d: Keep playing audio across slides when the presentation specifies a multi-slide playback range, and stop it when the range or presentation ends.
- e18d19d: Play wedge and newsflash slide transitions in presentation mode.
- e18d19d: Play horizontal and vertical random-bar slide transitions in presentation mode.
- e18d19d: Play random slide transitions in presentation mode, choosing a visual effect on each visit without changing the saved transition setting or speed.
- e18d19d: Play split, circle, diamond and plus slide transitions in presentation mode. Split honors horizontal/vertical orientation and inward/outward direction. Effects respect reduced motion, release interrupted layers, and wait to finish before automatic advancement.
- e18d19d: Play staggered strips slide transitions toward all four corners in presentation mode. Honor the saved direction and speed, reduced motion and interrupted-transition cleanup.
- e18d19d: Add bilingual corner handles for proportionally resizing multiple selected objects. Scale positions and dimensions together while preserving rotations and the opposite corner. Keep each resize in one Undo/Redo step and support Escape cancellation. This changes object geometry; font sizes and other appearance attributes retain their existing values.
- e18d19d: Add bilingual horizontal and vertical flip controls with mixed-selection state. Apply flip flags to every selected object in one undo step, preserving the other axis and unselected objects.
- e18d19d: Add English/Japanese No fill and No outline buttons to the quick appearance panel. Apply removal to all selected shapes in one undo step and preserve explicit no-paint settings when saved.
- e18d19d: Add bilingual quick controls for outline width in points and all eleven preset dash styles. Apply edits across the selection, preserve colors during width changes, reject invalid widths, and show inherited or mixed values.
- e18d19d: Honor saved automatic slide timing and click advance settings in presentation view. Cancel timers when leaving presentation, preserve their deadlines across unchanged live refreshes, and handle delays above the browser timer limit without advancing immediately.
- e18d19d: Add desktop-app-compatible Increase/Decrease Font Size keyboard shortcuts for canvas object selections, text editing, table cells, and outline text.
- e18d19d: Keep double, dotted, dashed, and wavy underlines visible while editing text and when pasting HTML. Preserve a solid strikethrough alongside patterned underlines.

  Preserve the equalized character height setting and other character formatting when copying and pasting text within the editor.

- e18d19d: Add optional speaker-notes color resolution through the notes master theme and color map. The notes editor uses it for display while retaining literal color references for editing and round trips.

  Expose notes line-break metadata and paragraph/break-aware editing so soft breaks survive save/load and undo.

  Keep the notes caret after inserted paragraph breaks and preserve theme-based typing colors across consecutive empty paragraphs.

- e18d19d: Add bilingual page setup to the development preview with 16:9, 4:3, 16:10 and custom slide sizes in inches or centimeters. Size changes apply to every slide, preserve object geometry, and support autosave and undo/redo.
- e18d19d: Add a bilingual table properties pane with cell selection, text, fill, alignment,
  row and column dimensions, insertion and deletion. Save edits and restore them
  through undo and reload.

  Allow setTableCellText to preserve unaffected formatting. Read soft breaks and
  field text in getTableCellText so editing does not silently drop visible content.

  Merge selected table cells in the preview while retaining their formatted text,
  and split merged cells. Expose splitTableCell and an append text policy for
  mergeTableCells; preserve text and reject malformed merges before mutation.

  Apply fill, alignment, bold and italic to selected cell ranges in a single undo
  step, including merged cells, from the English and Japanese preview.

  Use the shared text toolbar for table ranges, including font families, size,
  color and underline. Display saved color and mixed sizes when selection changes.

  Edit table border colors, widths and styles, with all-cell and outside-border
  placement. Reset borders and undo the changes from the bilingual preview.

  Double-click table cells directly on the canvas to edit their text and format
  selected text, with merged-cell hit areas, visible cell selection, save and undo.
  Allow setTableCellTextFormat to accept an optional UTF-16 range.

  Insert tables with a bilingual dialog for row and column counts, header rows
  and alternating row colors. Center new tables and select the first cell for
  immediate editing; preserve insertion and cell edits through history and reload.

- e18d19d: Preserve fonts, text colors, bold and italic formatting inherited from embedded table styles when displaying, editing and copying table text. Apply header, footer, banded row and column, and corner formatting in the reference desktop app's precedence order while keeping explicitly formatted cell text unchanged.
- e18d19d: Keep text with wrapping disabled on the same lines when editing, without scrolling its contents inside the text box as you type.
- e18d19d: Choose any of the reference desktop app's five radial gradient directions from the format pane, preserving the gradient stops and saving the matching focus and tile bounds.
- e18d19d: Group placeholder restoration and reset actions under Layout in the Home ribbon. Use distinct compact English and Japanese labels while retaining full accessible names and tooltips, and preserve button widths when the ribbon overflows. Allow sufficient height for group titles below two-line labels.
- e18d19d: Choose the five rectangular gradient directions from the format pane. The Path gradient direction control now remains visible and disabled, matching the reference desktop app.
- e18d19d: Keep multi-selection resize handles and width/height scaling aligned with the reference desktop app when selected shapes have different rotations.
- e18d19d: Restore the preceding solid, gradient or pattern fill settings when switching back after inserting a picture fill.
- e18d19d: Remember image fills when switching fill types in the editor. Returning to a previous picture fill restores its image, placement, transparency and supported cropping without reopening the file picker.
- e18d19d: Keep picture fill scale, alignment, mirroring and offsets when switching between tiled and stretched placement in the format pane.
- e18d19d: Keep grayscale and black-and-white image corrections when switching a shape's fill to another type and back to its picture fill.
- e18d19d: Render Background Styles thumbnails with the slide renderer so rectangular gradients, focus positions, and theme color adjustments match the slide preview.
- e18d19d: Add resetSlidePlaceholderGeometry to restore top-level placeholder position, size, rotation and flips from the current layout or master while preserving content and formatting. The bilingual development editor exposes it for individual or selected slides with undo/redo.
- e18d19d: Add `resetSlidePlaceholderTextFormatting` to restore inherited text, paragraph and text-body formatting on layout-bound placeholders while retaining content, hyperlinks, language and outline levels. Expose the operation in English and Japanese in the preview, with selected-slide support and undo/redo.
- e18d19d: Add `resetSlideLayout` to restore missing placeholders, layout geometry and inherited shape/text formatting together while preserving content and image relationships. Expose Reset layout in the English and Japanese preview, with multi-slide selection and a single undo step.
- e18d19d: Preserve a shape's saved aspect ratio when dragging its corner resize handles. Edge handles continue to resize only their corresponding axis, matching the reference desktop app.
- e18d19d: Add `addMissingSlidePlaceholders` to restore deleted layout slots without changing existing content or formatting. Expose the action in the preview editor in English and Japanese, with selected-slide support and undo/redo.
- e18d19d: Preserve imported gradient stop color adjustments when changing gradient position, direction, brightness, or transparency in the editor. Gradient stops now expose ordered `colorTransforms` for round-trip editing through the API.
- e18d19d: Retain picture fill transparency and remembered placement settings when inserting a replacement image from the format pane.
- e18d19d: Pressing Escape to close a collapsed ribbon group or a ribbon menu no longer also deselects the selected shape.
- e18d19d: Preserve group rotation and reflections when ungrouping objects. Children retain their transformed positions and compose the group angle and flips with their own, including nested groups. The development preview supports this through the existing English and Japanese ungroup controls and Undo/Redo.
- e18d19d: Keep rotated objects' resize handles aligned with the pointer and preserve the opposite anchor, including at minimum size. Hold Shift to maintain aspect ratio. Resize and rotation handles now have Japanese accessible labels.
- e18d19d: Edit paragraph indentation by dragging first-line, hanging, and left-indent markers on the ruler while editing horizontal text. Rulers follow the text body origin, preserve the text selection, and record each completed drag as one undo step. Press Escape to cancel a drag.
- e18d19d: Add, move and remove paragraph tab stops directly on the editor ruler. Choose left, center, right or decimal alignment, cancel a drag with Escape, and undo each completed gesture.
- e18d19d: Add a desktop-app-style Save Media As command for selected embedded audio and video, preserving the original bytes, MIME type, and filename extension.
- e18d19d: Apply shape appearance commands to every selected object in one undo step, including fills, outlines, effects and text formatting. Preserve text content and reject mixed non-text selections before changing text formatting.
- e18d19d: Keep quick fill and outline color controls synchronized with selected shapes, undo/redo and reload. Show mixed colors and non-solid or inherited paint states in English and Japanese instead of retaining the last edited color.
- e18d19d: Drag multiple selected objects together in the Selection Pane while preserving their relative stacking order. The operation supports grouped siblings and a single undo step.
- e18d19d: Allow text alignment and vertical anchoring on blank autoshapes. Add bilingual quick alignment controls that reflect mixed selections and apply to all selected text shapes with a single undo step.
- e18d19d: Keep double underline and double strikethrough when pasting HTML that applies decoration styles to underline or strikethrough tags.
- e18d19d: Support null in getShapeRunFormatEffective to resolve paragraph end formatting. Preserve empty shape paragraph sizes when displaying and editing text so the following text does not shift when editing begins.
- e18d19d: Allow getShapeRunFormatEffective to resolve fields with a fieldIndex selector. Preserve inherited fonts, sizes, colors, and emphasis for shape fields when displaying, editing, inspecting, and copying their text.
- e18d19d: Rotate multiple selected objects around their shared centre using a canvas handle. Preserve relative positions and angles, snap to 15-degree steps with Shift, and support cancellation and a single Undo/Redo step. The shared handle and rotation hints are available in English and Japanese. Single-object rotation now tracks the pointer angle from where the handle was grabbed.
- e18d19d: Preserve table cell ranges across panel focus changes and undo history. Global Delete clears the full range, and Shift+Arrow extends it from the active end without moving the table.
- e18d19d: Preserve and play links that return to the last viewed slide or end the slide show. Offer both destinations in the link editor for shapes, selected text, and table cells.
- e18d19d: Resolve next, previous, first, and last slide links against the active slide show's order, including repeated slides in custom shows. SVG navigation links now retain their action as a `#pptx-*` fragment so playback can choose the correct destination.

  Keep presenter playback active when it is opened immediately after leaving a fullscreen presentation.

- e18d19d: Match the reference desktop app's “Show without narration” and “Show without animation” checkboxes. Japanese show settings now disable the selected feature instead of enabling it, and saved presentations retain the same meaning in both languages.
- e18d19d: Preserve and render negative image crop offsets, including when switching away from a picture fill and restoring it in the editor.
- e18d19d: Add an English/Japanese alignment reference selector so multiple selected objects can align to slide edges and centers, with one-step undo and preserved dimensions.
- e18d19d: Preserve slides with absolute or normalized relative relationship targets when moving or sorting a presentation. Speed up moving multiple editor slides by applying the final order once.
- e18d19d: Show horizontal and vertical centimeter rulers from the editor View ribbon or menu. Rulers follow slide zoom and scrolling, and the display setting persists without changing the presentation.
- e18d19d: Add picture fill controls for image insertion, transparency, tile and stretch placement, alignment, mirroring, and rotation in the editor format pane.
- e18d19d: Edit solid fill and outline opacity through the existing shape setters while preserving imported theme colors and color transforms. The editor adds percentage transparency controls with mixed-selection support, undo and saved reloads.
- e18d19d: Choose solid fill and outline colors from theme and standard palettes in the format pane. Theme colors remain linked after saving, and locked shapes disable the color controls.
- e18d19d: Disable Grid and Guides menu commands and Grid Options in Slide Sorter, matching their availability in the reference desktop app (Mac).
- e18d19d: Match the reference desktop app's (Mac) 20–200% Slide Sorter zoom range in the slider, zoom buttons and dialog. Disable unavailable Fit and 400% dialog presets while preserving Normal view's independent 10–400% zoom.
- e18d19d: Keep the slide and text metrics, including small fonts, mixed-size paragraphs, and character spacing, in place when entering text editing, start editing with a single click inside text, and position the caret at the click. Adapt the Home ribbon to narrower windows without horizontal scrolling while keeping commands available in group menus.
- e18d19d: Tabs dialog buttons now allow clearing a typed tab position and setting an existing position, and reset the position after Clear or Clear All.
- e18d19d: Keep Delete and arrow shortcuts within selected table cells instead of deleting or moving the table. Support Shift+Arrow range selection, undoable range text clearing, and merged-cell navigation in the table pane, with English and Japanese guidance.
- e18d19d: Keep cell deletion and Select All scoped to the table selection. Add bilingual cell context menus, preserve a range when right-clicking inside it, and provide explicit whole-table selection. Keep context menus inside the viewport when opened near an edge.
- e18d19d: Preserve empty table paragraphs' authored font size in the preview and text editor, so blank lines no longer collapse to a default size. The font controls now resolve inherited formatting at an empty cell paragraph's caret. Pass a null run index to getTableCellRunFormatEffective to read the effective paragraph end format.
- e18d19d: Table fields such as dates and slide numbers now inherit cell text formatting in previews, text editing, and copied text. The existing getTableCellRunFormatEffective API accepts a fieldIndex selector to resolve field formatting.
- e18d19d: Keep table cell text readable when a table or its containing group is flipped, in both browser preview and SVG output.
- e18d19d: Edit the visible cell when clicking rotated or flipped tables. Keep cell selection highlights and text input aligned with the preview, including reflected and nonuniformly scaled parent groups, and preserve the displayed text orientation while editing.
- e18d19d: Honor table-cell text direction in previews and inline editing. Keep vertical text in place when editing cells with asymmetric margins, including bottom-to-top and upright right-to-left text.
- e18d19d: Convert shape and table-cell text with sentence, lowercase, uppercase, title, or toggle case through the existing text setters while preserving formatting, hyperlinks, and original run XML.

  Add a Change Case menu to the editor's Home ribbon for selected text, the word at the caret, and selected text shapes, including outline editing and Undo/Redo.

  Keep Home ribbon controls accessible without horizontal scrolling at intermediate window widths.

- e18d19d: Preserve supported text outline, shadow, glow and reflection effects when copying and pasting formatted text, while rejecting malformed effect metadata.
- e18d19d: Fall back to plain text when pasted character metadata contains unsupported underline or strike styles, font sizes, spacing, or colors. Preserve valid scheme-prefixed theme colors when copying formatted text.

  Ignore HTML font sizes smaller than one point so pasting web content cannot fail while applying its text formatting.

- e18d19d: The Text color button now shows the theme color chosen for subsequent typing
  (for example Accent 2) instead of keeping the previous swatch, in the canvas
  and the Home ribbon.
- e18d19d: Show the quick text content editor only for a single text-capable shape. Explain multi-selection editing in English and Japanese while retaining bulk text formatting, and hide the text editor for non-text objects.
- e18d19d: Keep slide and shape backgrounds visible while editing text, and hide the original glyphs so edited text is not drawn twice.
- e18d19d: Keep text edits when clicking another shape, and allow Shift-click to add that shape to the selection without an extra click.

  Clicking an empty slide area now ends text editing and clears the shape selection in one click.

- e18d19d: Paint text runs with a gradient fill (`<a:gradFill>`) or pattern fill (`<a:pattFill>`), such as the reference desktop app's gradient and pattern text-art presets. These runs were previously drawn in the default text color. A gradient spans the whole text block, including every line, as in the reference desktop app. Theme colors and their tints are resolved. Pattern fills use the same tiles as shape pattern fills. This works in both the SVG and the browser (`foreignObject`) text layouts. The editor canvas now shows these fills, also while the text is being edited.
- e18d19d: Support paragraph-internal line breaks when replacing shape or table-cell text with `newlines: 'break'`. The editor preserves these breaks for Shift+Enter and for Enter in title placeholders, matching the reference desktop app (Mac).
- e18d19d: Choose text colors from theme and standard color palettes when formatting selected text, objects or table cells, or preparing to type. Theme choices retain their theme references in saved presentations.
- e18d19d: Display theme-referenced background gradients instead of a solid color, including radial backgrounds selected with the reference desktop app's Background Styles gallery. Resolve gradient colors through the owning slide master's theme and color map while preserving the original theme and background XML on save.
- e18d19d: Close the collapsed Home drawing menu after an Arrange command so it no longer covers the slide or selection pane. Escape now dismisses the expanded group without clearing the selected shape.
- e18d19d: Add slide-background shape fills through `setShapeSlideBackgroundFill` and the editor's Fill pane, including multiple selection, undo and saved reloads. Fill readers expose the new `background` kind. Preview paints the slide background through these shapes while keeping it aligned through shape and group transforms.
- e18d19d: Create and edit common charts from English and Japanese preview dialogs.
  Change chart types, titles, series colors, category labels and numeric data,
  add or remove rows and series, and preserve edits through undo and reload.
  Keep embedded spreadsheet data synchronized with the saved chart.

  Set legend positions and show values, categories, series names and percentages
  from the chart dialog. Bind chart commands to exactly one selected chart.

- e18d19d: Allow `copyShape` to copy between presentations while preserving related images, charts, embedded workbooks and unknown dependencies. Allocate unique IDs for every shape inside copied groups and preserve their internal connector references.

  Fix cut and paste in the development preview. Copied content now retains its state even after editing or deleting the source or opening another document, and undo/redo restores the pasted selection.

- e18d19d: Keep heavy, long-dash, dash-dot, double-wave, and words-only underlines visible when entering text editing, without moving the text or changing solid strikethroughs. Expose `textUnderlineStyle` so custom HTML editors can share the preview's DrawingML underline rendering.
- e18d19d: Fix superscript and subscript in speaker notes jumping far from the text line, and allow existing underline and strikethrough styles to toggle off consistently.
- e18d19d: Add a separate bilingual presenter window with current and next slides, saved speaker notes, navigation controls, and a resettable elapsed timer. Keep speaker notes off the audience surface, synchronize live slide changes, and indicate when the presentation window closes.
- e18d19d: Read table-cell paragraph and outline-level character defaults with `getTableCellRunFormatEffective`. Preserve inherited fonts, sizes, emphasis, and colors when displaying, editing, inspecting, and copying table text.

  Resolve inherited shape and table text colors through the slide color map, including detached editing previews.

- e18d19d: Keep right, center, and decimal tabs aligned while editing text whose kerning crosses formatting-run boundaries.
- e18d19d: Render double, dotted, dashed, and heavy underline styles distinctly in previews, including dash-dot and double-wave patterns. Keep strikethrough solid when combined with a patterned underline, and omit spaces from words-only underlines.

  Enable Home font controls for selected table cells, including the Font dialog, while preserving mixed formatting and single-step undo.

- e18d19d: Stop the trim dialog's media preview at the selected endpoint and update audio fades smoothly during playback.
- e18d19d: Replace inherited Text Art gradient or pattern fills when choosing a font color
  for subsequent typing in the canvas, outline, and speaker notes.
- e18d19d: Reject invalid slide dimensions before modifying the presentation. Page setup now enforces PPTX’s 1–56 inch range in both inches and centimeters, with matching Japanese and English validation messages.
- e18d19d: Align custom tab stops along the text direction while editing vertical text, including upright glyphs and center, right, and decimal alignment.
- e18d19d: Keep bottom-to-top text in place when entering text-box editing with asymmetric margins.
- e18d19d: Add brightness and contrast presets to the Video Format ribbon, with a single undo action per selection. Keep these adjustments visible while playing a video in the editor.
- e18d19d: Update video brightness and contrast while dragging their sliders, with a single undo step for each completed drag.
- e18d19d: Adjust a video's picture size and offset independently of its crop frame in the format pane. Reset restores the entire picture, and crop changes support undo and redo.
- e18d19d: Add a dedicated Video tab to the format pane with desktop-app-style brightness and contrast controls.
- e18d19d: Add video formatting reset to the editor ribbon and the `resetShapeVideoFormatting` API. Reset removes color corrections, borders and effects and restores a rectangular shape while preserving the video, poster image, crop and dimensions.
- e18d19d: Add a dedicated video format ribbon with playback, poster frame, border, effects, arrangement, size, and format pane controls.
- e18d19d: Reset video recoloring, brightness, and contrast from the Video format pane without changing the poster, playback settings, or shape formatting.
- e18d19d: Set a video poster image from the current playback frame or an image file. Keep the poster visible until playback or seeking starts, and preserve the embedded video when changing its poster.
- e18d19d: Reset a video's poster image to its first frame from Video Format > Poster Frame, with undo support and without replacing the video or its playback settings.
- e18d19d: Add desktop-app-style video recolor presets to the Video format ribbon and pane.
- e18d19d: Add theme color shades to the video recolor palette and preserve their theme references when saving. Distinguish black-and-white threshold thumbnails and apply theme colors to duotone thumbnails.
- e18d19d: Add Gridlines and Notes toggles to the View ribbon. Disable Gridlines, Guides and Notes in Slide Sorter while retaining their settings when returning to Normal view.
- e18d19d: Preserve and edit wheel transition spoke counts, with Japanese and English controls. Play clockwise wheel transitions in presentation mode, including default four spokes and a fade for zero spokes.
- e18d19d: Use zero spacing between text columns when the presentation omits a column gap, matching the reference desktop app (Mac) in previews and inline editing.
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [4b27073]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
- Updated dependencies [e18d19d]
  - @office-kit/pptx-preview@0.13.0
  - @office-kit/pptx@0.22.0
  - @office-kit/pptx-dsl@0.9.0

## 0.9.1

### Patch Changes

- 3c326b7: fix: a double-click text edit whose new text contains `$&`, `$'` or `` $` `` no longer fails as "affects other text or slides". The edit rules now come from `@office-kit/pptx-dsl/source-edit`.
- Updated dependencies [3c326b7]
- Updated dependencies [3c326b7]
  - @office-kit/pptx-dsl@0.8.0
  - @office-kit/pptx-preview@0.12.0

## 0.9.0

### Minor Changes

- 92eef76: `init` now writes the project guide as `AGENTS.md` as well as `CLAUDE.md`, so coding agents other than Claude Code pick up the TSX workflow.

  Claude Code loads `CLAUDE.md` automatically; Codex and most other agents load `AGENTS.md`. A generated project only had the first, so those agents started with no instructions and reached for their own slide-building tools — producing a deck outside the project instead of editing `deck.tsx`. Both files carry the same text, and existing projects can get the same result by copying `CLAUDE.md` to `AGENTS.md`.

## 0.8.3

### Patch Changes

- 98273a7: Submit selected-area instructions and save in-place text with Shift+Enter. Remove the separate text editing dialog, reveal saved previews before visual review finishes, and use compact Undo/Redo icon buttons.
- Updated dependencies [28731c2]
  - @office-kit/pptx@0.21.0
  - @office-kit/pptx-dsl@0.7.0
  - @office-kit/pptx-preview@0.11.0

## 0.8.2

### Patch Changes

- Updated dependencies [4235418]
- Updated dependencies [4235418]
  - @office-kit/pptx-preview@0.10.0
  - @office-kit/pptx-dsl@0.6.0

## 0.8.1

### Patch Changes

- Updated dependencies [04eda82]
  - @office-kit/pptx@0.20.0

## 0.8.0

### Minor Changes

- cc7740b: Edit slides directly without switching modes: click objects for AI instructions, drag a region, or double-click text to edit in place. Undo and Redo now restore source and asset changes across AI turns, text edits, and external saves, with keyboard shortcuts and grouped concurrent agent edits.

### Patch Changes

- Updated dependencies [3eecc92]
- Updated dependencies [cb702c1]
  - @office-kit/pptx@0.19.0

## 0.7.0

### Minor Changes

- 758df05: Edit selected slide regions through a chosen coding agent, directly save uniquely identifiable text literals with guarded undo, and automatically provide screenshots of changed slides for bounded AI design review.

## 0.6.0

### Minor Changes

- 9b8f8a0: Reduce preview latency on large TSX decks by avoiding repeated parsing of existing slides, themes and relationships. Keep existing slide handles live when appending or duplicating slides.

  Verify builds when an embedded agent finishes and return failures to Claude Code or Codex for up to three automatic repair attempts. Support Shift+Enter in the Claude terminal and provide native TSX agenda examples and foreground/background contrast guidance.

### Patch Changes

- Updated dependencies [9b8f8a0]
  - @office-kit/pptx@0.18.3

## 0.5.0

### Minor Changes

- 596ec79: Run multiple coding agents in resizable horizontal or vertical panes, each with its own session and the focused slide as context. Refresh the preview with an Office Kit logo, dark workspace chrome, violet accents and a softly lit slide canvas. Fix terminal sizing so Claude Code's rightmost characters remain visible when resizing or splitting panes.

## 0.4.2

### Patch Changes

- 716a089: Reduce TSX save-to-preview latency by reusing the compiler and unchanged slide renders, while preserving fresh deck evaluation and invalidating previews when shared resources change.

## 0.4.1

### Patch Changes

- 1ca6ad7: Resize the preview chat panel by dragging its left edge or using the keyboard, with the chosen width remembered across reloads.

## 0.4.0

### Minor Changes

- 7982fc9: Embed interactive Claude Code in the preview, including its model selection, skills, settings and permission prompts. Attach focused-slide context with a prompt hook, preserve the session across browser reloads, simplify the chat layout, and render Markdown in Codex responses.

### Patch Changes

- Updated dependencies [bb5a337]
- Updated dependencies [3c9850f]
- Updated dependencies [55fc76b]
  - @office-kit/pptx-preview@0.9.9

## 0.3.0

### Minor Changes

- 9be0737: Add a preview chat panel backed by local Claude Code or Codex, with focused-slide context, automatic TSX edit previews, conversation recovery and cancellation.

## 0.2.4

### Patch Changes

- c47afdd: feat: slide text in the preview can be selected and copied

  The viewer rendered each slide inside a sandboxed `<iframe>` with pointer events disabled, so nothing on a slide could be selected. The slide now lives in a shadow root on the stage: the deck's SVG stays isolated from the viewer's DOM and CSS, but its text (XHTML inside `<foreignObject>`) is ordinary selectable content, and arrow-key navigation keeps working after clicking into a slide. In presentation mode a click still advances the deck unless it made a selection.

## 0.2.3

### Patch Changes

- Updated dependencies [7e89e62]
- Updated dependencies [36c4987]
- Updated dependencies [36c4987]
- Updated dependencies [36c4987]
- Updated dependencies [64fa3c7]
- Updated dependencies [7e89e62]
  - @office-kit/pptx@0.18.0
  - @office-kit/pptx-dsl@0.5.0

## 0.2.2

### Patch Changes

- 8278aa4: Initialize slide projects with separate slide files, a shared theme and a small deck entry. Guide AI revisions toward focused source patches and the running preview instead of regenerating the deck.
- Updated dependencies [983eb10]
  - @office-kit/pptx-dsl@0.4.0

## 0.2.1

### Patch Changes

- ad3cbd2: Keep slide previews steady while editing TSX: update only changed thumbnails and slides, preserve zoom and scroll position, and retain the previous frame until its replacement is ready. Transfer SVG changes incrementally, preload build workers between saves, and cancel obsolete evaluations so rapid edits and accidental infinite loops do not delay the next revision.
- Updated dependencies [71deec5]
- Updated dependencies [71deec5]
  - @office-kit/pptx@0.17.0
  - @office-kit/pptx-preview@0.9.6
  - @office-kit/pptx-dsl@0.3.0

## 0.2.0

### Minor Changes

- 4fe8fde: Add typed TSX authoring without a React or Vue runtime, including native text,
  shapes, images, tables and charts, template slide reuse and explicit editing,
  and a Raw callback escape hatch. Add local project initialization, PPTX export,
  template inspection and watch preview with error recovery, plus VSCode tasks
  and Claude Code authoring guidance.
- a9a2598: Add a view-only slide workspace with vertical thumbnails, keyboard navigation,
  fit/zoom and full-screen presentation. Live TSX updates retain the selected slide;
  all content changes continue to happen in source rather than on the canvas.

### Patch Changes

- Updated dependencies [4fe8fde]
- Updated dependencies [b8adc13]
- Updated dependencies [a13afc2]
  - @office-kit/pptx-dsl@0.2.0
  - @office-kit/pptx@0.16.1
  - @office-kit/pptx-preview@0.9.5
