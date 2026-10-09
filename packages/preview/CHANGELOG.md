# pptx-kit-preview

## 0.15.0

### Minor Changes

- d8e455d: The reference desktop app's 74 built-in table styles are now fully supported.

  - `setTableStyleId` accepts a built-in style's English name as well as a GUID (`setTableStyleId(table, 'Light Style 1 - Accent 2')`). Applying a built-in style writes the reference desktop app's definition of it into `ppt/tableStyles.xml` (creating the part when a deck has none), as the reference desktop app does, so Keynote, Google Slides and LibreOffice draw the same table. A deck's own definition of a GUID is kept. `addSlideTable` likewise writes the definition of its default style.
  - New `BUILTIN_TABLE_STYLES` lists every built-in style (`id`, `name`, gallery `category`) in the order of the reference desktop app's Table Styles gallery.
  - `getTableCellAppearanceEffective` now resolves every built-in style, not only Medium Style 2 - Accent 1 and No Style, Table Grid, and reports `fillOpacity` for translucent fills. New `getTableBackgroundEffective` resolves the table background (`a:tblPr` fill, else the style's `a:tblBg`).
  - The preview draws all table style parts for built-in and custom styles, including translucent bands and the Themed Styles' background.
  - The editor's Table Design ▸ Table Styles gallery offers all built-in styles, grouped as in the reference desktop app (Best Match for Document, Light, Medium, Dark), with swatches drawn by the preview renderer that follow the Table Style Options check boxes, the reference desktop app's English and Japanese names, and Clear Table.

- c535941: feat: copy slides and objects to the system clipboard. Copying slides or objects in the editor now puts them on the clipboard for other presentation apps and documents: slides and drawn objects paste there as pictures at their size on the slide, and text boxes and tables copied on their own paste as text and tables. Another editor — in another tab or browser — pastes them back as editable slides and objects, with their pictures, layouts and themes. Pictures and text copied in other apps paste as a picture or a text box.

  feat: the slide thumbnail menu has **Download Selected Slides...**, which saves just the selected slides as a `.pptx` to import into another presentation app.

  feat: `renderSlideToSvg(pres, slide, { background: false })` draws only the slide's own shapes on a transparent surface, without the background or master and layout graphics.

- 550b55c: Gradient lines, sketched lines and more text paints:

  - **Gradient lines.** `setShapeStroke(shape, { fill: { kind: 'gradient', ...gradient } })` writes `<a:ln><a:gradFill>` with the same gradient options as `setShapeGradientFill`, and a text outline takes the same `fill` (`setShapeTextFormat(shape, { outline: { fill } })`). `getShapeStrokeGradient(shape, pres?)` reads it back; `getShapeStroke` / `getShapeStrokeEffective` now report `{ kind: 'gradient' }` for such lines (previously `inherit`), so a `switch` over `ShapeStroke['kind']` needs a `gradient` case.
  - **Sketched lines.** `setShapeStrokeSketch(shape, 'curved' | 'freehand' | 'scribble' | null)` / `getShapeStrokeSketch(shape)` write and read the reference desktop app's Sketched style (`ask:lineSketchStyleProps`). Other `<a:ln>` extensions are kept, and a sketch the reference desktop app saved (which replaces the geometry with the hand-drawn path) gets its original geometry back when the sketch is changed or removed.
  - **Text fills.** `TextFormat.textFill` accepts `{ kind: 'none' }` (`<a:noFill/>`) and `{ kind: 'image', bytes }` (a PNG, JPEG, GIF, BMP, TIFF or WebP stretched over the text, as the reference desktop app writes it), on shapes, table cells and notes.
  - **Keep text flat.** `setShapeTextFlat(shape, flat)` / `getShapeTextFlat(shape)` write `<a:bodyPr><a:flatTx/>`.
  - **Preview** strokes gradient outlines on shapes, connectors and text, paints no-fill and picture-filled text, and draws a hand-drawn approximation of sketched outlines.
  - **Editor**: Gradient line in Line and Text Outline (with the gradient type, direction, angle and stop controls of Gradient fill), the Sketched style menu (None, Curved, Freehand, Scribble), No fill and Picture or texture fill for text, and Keep text flat are now enabled. Text Effects ▸ Soft Edges stays disabled, as in the reference desktop app. The Text Fill gradient and pattern options no longer fail to apply.

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

- 21d59f7: Picture Styles, and Compress Pictures that labels pictures the way the reference desktop app does.

  - New `setShapePictureStyle(picture, 'Metal Oval')` applies one of the reference desktop app's 28 built-in picture styles. It writes exactly the `p:spPr` markup the reference desktop app (Mac, 16.113) saves for that style (geometry, fill, border, effects and 3-D, with literal colors) and keeps the picture, its crop and its position. `getShapePictureStyle` returns the style a picture carries exactly, or `null`. New `BUILTIN_PICTURE_STYLES` lists the style names in the order of the reference desktop app's gallery.
  - New `setShapeImageCompressionState` / `getShapeImageCompressionState` write and read the picture's `a:blip/@cstate` (`'print'`, `'screen'`, `'email'`, …) together with the `a14:useLocalDpi` extension the reference desktop app writes beside it.
  - Preview: pictures now draw their effects (outer and inner shadow, glow, soft edge, reflection), their own fill and an approximation of their 3-D: the camera rotation as a flat projection and a top bevel as edge lighting.
  - Fix (preview): a reflection faded the wrong way, strongest at its far edge, and ignored its end position; it now starts at the shape's edge and fades out by `endPos`. A soft edge blurred the whole shape; it now only feathers the outline.
  - The editor's Picture Format ▸ Picture Styles gallery works: the 28 styles in the reference desktop app's order, drawn by the preview renderer, with the reference desktop app's English and Japanese names as tooltips and the applied style checked.
  - The editor's Compress Pictures writes `cstate` for Print, On-screen and Email, as the reference desktop app does, and only replaces a picture's pixels when cropped areas are removed or the resampled picture is smaller. Picture Quality now opens the reference desktop app's menu: Compress Pictures... and Upscale Picture (unavailable: it uses the vendor's cloud AI service).

- 1d447f8: Decimal tab stops align on the decimal separator of each run's language (`,` for `de-DE`, `fr-FR`, …) instead of always `.`. Runs without a language still use `.`. `decimalSeparatorOf(lang)` is exported. Paragraph indents in vertical text rendered as HTML now run along the text lines, as in the reference desktop app, instead of being applied as a left margin.

### Patch Changes

- 0d26527: Colors with a large `satMod` (or `satOff`) now resolve to the color the reference desktop app paints. `resolveDrawingColor` capped saturation at 100%, as ECMA-376 §20.1.2.3.27 describes, but the reference desktop app does not cap it. Because of that cap, the theme gradients that use `satMod` 300–350% came out too light and too gray in the preview. The Themed Style table backgrounds and the gradient fills of the built-in 2007-era themes are the visible cases: an orange accent was off by up to 70 levels per channel. They now match the reference desktop app's own exports within 2 levels.

  `resolveDrawingColor` now also applies the `hue`, `sat`, `lum`, `red`/`green`/`blue` (with their `Mod`/`Off` forms), `gamma` and `invGamma` transforms, which it ignored before.

- 946dac4: The READMEs describe compatibility in terms of presentation apps in general instead of naming a third-party product.
- 9d0754c: API documentation (the TSDoc shipped in the type declarations and in the editor's tool descriptions) no longer uses third-party product names. Behaviour that was checked against a specific desktop presentation app is now described as "the reference desktop app". The `@office-kit/pptx-dsl` package description and the `@office-kit/pptx-preview` npm keywords were reworded the same way. No API or output changes.
- 6ac3ead: Themed Style 2 tables now show the same background gradient as the reference desktop app. The reference desktop app does not blend this two-stop table background at a constant rate: the top color holds through the header row, then the blend gets steeper toward the bottom. The preview used a constant rate, which made the middle rows up to 21 levels per channel too dark. It now uses the curve fitted to the reference desktop app's exports and matches them within 2 levels for all six accents.
- Updated dependencies [c8512fe]
- Updated dependencies [4410413]
- Updated dependencies [d8e455d]
- Updated dependencies [0d26527]
- Updated dependencies [550b55c]
- Updated dependencies [afe80eb]
- Updated dependencies [5e6bbc5]
- Updated dependencies [3f4aefe]
- Updated dependencies [9d0754c]
- Updated dependencies [af5b2b4]
- Updated dependencies [1d447f8]
- Updated dependencies [acb95fc]
- Updated dependencies [21d59f7]
- Updated dependencies [2912bdc]
- Updated dependencies [ec775df]
- Updated dependencies [197b737]
- Updated dependencies [24a6ae0]
  - @office-kit/pptx@0.24.0

## 0.14.0

### Minor Changes

- e860ae8: Preset shapes are now drawn from ECMA-376's own preset definitions instead of hand-written approximations. The preview, the editor canvas and its Shapes gallery icons showed many presets wrongly — hearts, lightning bolts, suns, moons, clouds, brackets and braces, bent and curved arrows, the equation shapes, stars with 7 to 32 points — and now match the reference desktop app, including the lit and shaded faces of cubes, cans and curved arrows, adjust handles on every preset, and elliptical arcs in custom geometry (whose `arcTo` angles were read as parametric instead of visual angles). Text in a preset now wraps inside the rectangle the definition gives it, such as an ellipse's inscribed rectangle, as the reference desktop app does.

  - `@office-kit/pptx`: new `getPresetGeometry(preset, size, adjustValues?)` evaluates any preset's paths and text rectangle in the same form as `getShapeCustomGeometry`.
  - `@office-kit/pptx-preview` (breaking): `shapeCustomTextRect(custom, extent)` is replaced by `shapeTextRect(shape)`, which reads custom and preset geometry alike, and `resolveTextBodyRect` now takes `(bounds, margins, region)` with the region from `shapeTextRect`.

- 089140c: The preview now draws text bevels (`<a:bodyPr>` `<a:sp3d><a:bevelT>`, as the reference desktop app's Soft Bevel and Sharp Bevel text styles write them): beveled glyphs are shaded with a light and dark edge sized from the bevel's width and height and lit from the light rig's direction, in shape text, table cells and the editor while text is being edited. The 12 bevel presets get different edge profiles and shiny materials add a highlight; this is a 2-D approximation, so the camera, extrusion and contour are not drawn. Text without a bevel renders exactly as before. New `getTableCellText3D` reads a table cell's text 3-D, like `getShapeText3D` does for shapes.

### Patch Changes

- Updated dependencies [e860ae8]
- Updated dependencies [089140c]
- Updated dependencies [8db2360]
  - @office-kit/pptx@0.23.0

## 0.13.0

### Minor Changes

- e18d19d: Support the reference desktop app's Top, Middle and Bottom Centered text anchors through the existing `setShapeTextAnchor` API's `centered` option. Resolve inherited centering through `getShapeBodyPrEffective`, preserve paragraph alignment, and expose all six anchor choices in the editor with preview and text-editing support.

  Preview integrations can use `shapeTextAnchorOffset` to position editable text consistently with the rendered text block.

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

- e18d19d: `getShapeCustomGeometry` now reports the `<a:rect>` a custom-geometry shape
  states for its text, as `textRect`, with its guide formulas already evaluated.

  The preview lays a custGeom shape's text in that rectangle instead of the whole
  bounding box, through the new `shapeCustomTextRect` in `@office-kit/pptx-preview`,
  and inline editing in the dev editor puts the caret in the same place. A custom
  shape that states no rectangle still gets its whole box — a preset's
  approximated region is never substituted for a shape that describes itself.

- e18d19d: Show bullet and numbered-list markers during inline editing, including nested numbering, without changing copied text or selection offsets. Share the preview renderer's numbering through `paragraphNumberLabels` so editing surfaces use the same counter and restart rules. Build inline paragraph text in one pass instead of copying the whole shape for every paragraph.
- e18d19d: Inline editing now reads vertical text the way the preview paints it, and
  shrinks autofit text by the same factor — a `<a:normAutofit/>` title no longer
  jumps back to its authored size the moment the caret appears.

  Two new exports carry the shared rules: `verticalTextStyle` / `textColumnsStyle`
  turn `<a:bodyPr vert=… numCol=… spcCol=…>` into the CSS both surfaces use, and
  `shapeAutoFitScale` reports the shrink factor the renderer applies to a shape,
  for the box it is actually laid out in.

- e18d19d: Keep character reflection and inner-shadow effects visible while shape or table-cell text is being edited, including tables in scaled groups, and align rotated editing text around the text body's inner rectangle.

  Add `renderTextEffectsSvg` to render character reflection and inner-shadow overlays in local text-box coordinates, including text-body rotation and vertical layout.

- e18d19d: Add setShapePreset and read picture presets through getShapePreset. Clip picture previews to their preset shape while preserving source-image crop and effects. Add bilingual image-shape controls for ellipse, rounded rectangle, triangle, diamond, pentagon, hexagon, star, and heart masks, with undo and saved persistence.
- e18d19d: Share the preview's preset text rectangle calculation through `resolveTextBodyRect`. Inline editing now retains the preview's constrained text region in triangles, diamonds, pentagons, stars and double arrows, and respects default centered autoshape paragraphs.
- e18d19d: Expose optional glyph ink ascent and descent measurements separately from text line metrics. Ink bounds include shaped glyph placement, preserve signed distances from the baseline, and remain unavailable for estimated glyphs.
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

- e18d19d: Extend the existing paragraph formatting API to table cells, render their bullets and paragraph spacing, and add bilingual editor controls for individual cell paragraphs and selected cell ranges. Preserve rich text and support undo, redo, and save/reload.
- e18d19d: Paint text runs with a gradient fill (`<a:gradFill>`) or pattern fill (`<a:pattFill>`), such as the reference desktop app's gradient and pattern text-art presets. These runs were previously drawn in the default text color. A gradient spans the whole text block, including every line, as in the reference desktop app. Theme colors and their tints are resolved. Pattern fills use the same tiles as shape pattern fills. This works in both the SVG and the browser (`foreignObject`) text layouts. The editor canvas now shows these fills, also while the text is being edited.
- e18d19d: Keep heavy, long-dash, dash-dot, double-wave, and words-only underlines visible when entering text editing, without moving the text or changing solid strikethroughs. Expose `textUnderlineStyle` so custom HTML editors can share the preview's DrawingML underline rendering.

### Patch Changes

- e18d19d: Render DrawingML picture biLevel effects as black and white, including saturated colors, and preserve the requested threshold without rounding to coarse steps.
- e18d19d: Preview background images with their stretch offsets, tile scale, alignment, offsets, and alternating reflections, including inherited backgrounds. Expose natural background image dimensions using embedded resolution and fill DPI, so tiling matches the image size instead of stretching it across the slide.
- e18d19d: Render background picture cropping for stretched and tiled images, including inherited backgrounds. Add `getSlideBackgroundImageCrop` to read the effective source rectangle without changing the deck.
- e18d19d: Read and edit slide background picture opacity, including inherited backgrounds. Format Background now provides a transparency slider and percentage field, and the preview displays the selected opacity.
- e18d19d: Render character-level outer shadows and glows in the SVG preview path, including mixed formatted runs and theme colors.
- e18d19d: Support `TextFormat.underlineColor` for independently colored underlines. Set it to `null` to follow the text color. Preserve the color when reading, editing, copying, and saving text, and expose underline color and Automatic in the editor's Font dialog.
- e18d19d: Render custom paragraph tab stops in browser previews, including left, center, right and decimal alignment using the browser's font measurements.
- e18d19d: Preserve tab characters and paragraph default tab spacing in browser previews and direct text editing, including scaled text.
- e18d19d: Display the borders and transparent cells of the built-in No Style, Table Grid style when a presentation stores only its style ID.
- e18d19d: Support reading and changing transparency on image-filled shapes as well as pictures, including preview rendering. Invalid opacity values now leave the existing transparency unchanged.
- e18d19d: Keep chart titles, axes, data labels, trendline names, and legends readable when a chart or parent group is flipped. Preserve label rotation and its side of the anchor.
- e18d19d: Follow custom slide show links during playback and in presenter view, with the option to return to the calling slide when the linked show ends.
- e18d19d: Expose inherited bullet color, size, and font details through `getParagraphPropertiesEffective`.

  Preserve bullet formatting inherited from layouts and masters in previews and while editing text, including explicit follow-text overrides.

- e18d19d: Hide bullet markers on empty paragraphs in the browser preview, matching the reference desktop app while preserving their list formatting.
- e18d19d: Add bilingual image upload and replacement dialogs, crop and appearance controls,
  and alternative text editing to the development preview. Preserve picture geometry
  and crop during replacement and support undo and persisted reloads.

  Correct signed image contrast rendering so zero is neutral and negative contrast
  reduces color separation without inverting the picture.

- e18d19d: Keep linear gradients fixed to the slide when Rotate with shape is disabled, including on wide or tall rotated shapes.
- e18d19d: Preserve gradient stop transparency when reading shapes and expose theme-resolved stop colors. Shape previews now render the brightness and transparency saved by the reference desktop app for linear and radial gradient stops.
- e18d19d: Add image-fill placement readers and setters for tile alignment, scale, offsets, mirroring, stretch offsets and rotation with the shape. Placement changes preserve the embedded image, crop and effects. Preview now renders stretch offsets and clips the image to the shape.
- e18d19d: Keep shape text readable inside flipped groups, including nested rotated groups, while preserving the transformed text position and baseline direction.
- e18d19d: Add `isSlideBackgroundGraphicsHidden`, `setSlideBackgroundGraphicsHidden`, and `isSlideLayoutBackgroundGraphicsHidden` to inspect and control inherited decoration without deleting template content. Apply to All also copies the graphics visibility setting to slides and layouts. The preview honors both levels, and the editor's Format Background pane supports changing selected slides with undo and save/reload.
- e18d19d: Recognize `true` and `false` as well as numeric DrawingML flip flags in imported presentations. Flipped shapes and groups now retain their orientation in the preview and when ungrouped, including after saving and reloading.
- e18d19d: Preserve the opening text size of placeholders that inherit automatic fitting without a saved shrink factor. Titles now retain their intended line breaks instead of shrinking prematurely.
- e18d19d: Resolve inherited text autofit and columns through `getShapeBodyPrEffective`, and use them consistently in previews and editing controls. Allow `setShapeTextColumns` to author one column explicitly, overriding inherited columns; invalid column settings leave the previous values intact.
- e18d19d: Keep text at the same layout scale in the preview and editing view so entering text editing no longer shifts wrapped lines at fractional zoom levels. Preserve the size, spacing, and font of large or explicitly styled bullets while editing.
- e18d19d: Resolve internal slide link targets by package and part name so preview links retain the correct destination after reading, reordering, and saving a presentation.
- e18d19d: Apply OOXML kerning thresholds consistently in SVG and HTML preview layout, fontkit and browser measurement, and inline text editing.
- e18d19d: Preserve the font size of leading empty lines in shapes and table cells, preventing text from jumping when editing begins.
- e18d19d: Apply last-column formatting to horizontally merged table cells that reach the right edge, matching the reference desktop app even when the cell starts in an earlier column.
- e18d19d: Apply total-row formatting to vertically merged table cells that reach the last row, matching the reference desktop app even when the cell starts in an earlier row.
- e18d19d: Fix existing text in shapes and table cells incorrectly adopting the font, size, or emphasis saved for newly inserted text at the end of a paragraph.
- e18d19d: Keep ordinary Latin words intact in SVG previews unless the paragraph explicitly allows mid-word wrapping, while still emergency-splitting words that cannot fit on an empty line as the reference desktop app does. This matches the DrawingML `latinLnBrk` setting and the reference desktop app's handling of narrow preset shapes.

  When mid-word wrapping is enabled, use the remaining line width while keeping East Asian closing punctuation attached to its preceding character.

- e18d19d: Render picture outlines along the image shape, including cropped masks and rotation. Add bilingual image-border color, width, and line-style controls with undo, redo, and saved persistence.
- e18d19d: Preserve image detail when applying Washout and match the reference desktop app's brightness and contrast calculation in slide previews, live video, and correction thumbnails.
- e18d19d: Preserve different outline colors and widths on adjacent text runs in SVG previews instead of applying the first run's outline to the entire word.
- e18d19d: Preserve double strikethrough in browser and SVG previews instead of displaying it as a single line, including text with underlines.
- e18d19d: Render picture duotone recolor colors without unintentionally brightening their RGB values.
- e18d19d: Render radial gradient focus positions using the reference desktop app's edge insets, so corner directions no longer appear centered.
- e18d19d: Preserve custom superscript and subscript offsets in previews instead of drawing all offsets at the same height. Keep the run's own font size when shrinking script text in HTML previews.
- e18d19d: Align browser-preview custom tabs with the painted width of text that uses expanded or condensed character spacing, including combining characters.
- e18d19d: Render character-level inner shadows in both SVG and editable foreignObject preview text paths.
- e18d19d: Render PNG and JPEG image fills as repeating tiles with alignment, offsets, independent scaling and alternating horizontal/vertical reflections. Use embedded PNG/JFIF resolution or an explicit fill DPI to size the tiles. Add `getShapeImageIntrinsicSize` to read the image's unscaled physical size.
- e18d19d: Preserve fonts, text colors, bold and italic formatting inherited from embedded table styles when displaying, editing and copying table text. Apply header, footer, banded row and column, and corner formatting in the reference desktop app's precedence order while keeping explicitly formatted cell text unchanged.
- e18d19d: Render character-level DrawingML reflections in the SVG preview using the laid-out glyph positions, including per-run opacity, fade, blur, and vertical scale.
- e18d19d: Render rectangular gradient fills with rectangular contours and the correct center or corner direction. Correct radial gradient color order so the first stop appears at the focus, matching the reference desktop app.
- e18d19d: Apply gradient stop brightness and theme colors when previewing slide, layout and master backgrounds. Background gradient readers now include resolved stop colors while preserving their original color tokens for editing.
- e18d19d: Preserve the axis-swapped text rectangle used by pure SVG vertical layout, including asymmetric text margins.
- e18d19d: Open the editor's Selection Pane to select nested objects, rename them, and show or hide individual objects or the whole slide's objects with undo and autosave.

  Fix renaming and visibility changes on group shapes. Omit hidden objects and hidden group descendants from previews and canvas hit targets while preserving their editable content.

- e18d19d: Support null in getShapeRunFormatEffective to resolve paragraph end formatting. Preserve empty shape paragraph sizes when displaying and editing text so the following text does not shift when editing begins.
- e18d19d: Allow getShapeRunFormatEffective to resolve fields with a fieldIndex selector. Preserve inherited fonts, sizes, colors, and emphasis for shape fields when displaying, editing, inspecting, and copying their text.
- e18d19d: Read and preview theme-referenced shape fills, including solid colors, transparency, and gradient stops, using each shape's slide master theme. Direct shape fills continue to override theme references. The fill opacity reader accepts an optional presentation argument to resolve theme transparency.

  Use a shape style's text color and theme font defaults when its existing text formatting does not supply them, so text in the reference desktop app's colored shape styles keeps its intended color.

- e18d19d: Resolve text fonts from each slide's own master theme, and honor shape-style font and color defaults ahead of inherited placeholder formatting. Shapes using their group's fill now inherit solid colors, transparency, and gradient details instead of falling back to their own theme style.
- e18d19d: Preserve and play links that return to the last viewed slide or end the slide show. Offer both destinations in the link editor for shapes, selected text, and table cells.
- e18d19d: Resolve next, previous, first, and last slide links against the active slide show's order, including repeated slides in custom shows. SVG navigation links now retain their action as a `#pptx-*` fragment so playback can choose the correct destination.

  Keep presenter playback active when it is opened immediately after leaving a fullscreen presentation.

- e18d19d: Preserve and render negative image crop offsets, including when switching away from a picture fill and restoring it in the editor.
- e18d19d: Set solid slide background transparency with the optional opacity argument to `setSlideBackground`. Background readers and previews retain the alpha channel, and the editor provides transparency controls for selected slides with Undo and save/reload support.
- e18d19d: Keep Latin words together across formatting changes when breaks within words are disabled, so making part of a word bold or italic does not introduce an unwanted line break.
- e18d19d: Render paragraph tab stops and default tab spacing in SVG text and Node image output, including left, center, right, and decimal alignment in text boxes and table cells.
- e18d19d: Preserve explicit cell-side borders in banded table styles, including interior cells, while retaining interior borders when a side is unspecified.
- e18d19d: Correct default table cell margins in SVG and browser previews so text wrapping and padding remain proportional when fitting content to a new page size. Explicit zero margins remain supported.
- e18d19d: Preserve empty table paragraphs' authored font size in the preview and text editor, so blank lines no longer collapse to a default size. The font controls now resolve inherited formatting at an empty cell paragraph's caret. Pass a null run index to getTableCellRunFormatEffective to read the effective paragraph end format.
- e18d19d: Table fields such as dates and slide numbers now inherit cell text formatting in previews, text editing, and copied text. The existing getTableCellRunFormatEffective API accepts a fieldIndex selector to resolve field formatting.
- e18d19d: Add effective table cell appearance resolution for embedded table styles and the built-in Medium Style 2 – Accent 1 style. Resolve theme fill and line references, preview solid cell fills and styled borders, preserve explicitly transparent cells, and stop adding white table backgrounds or gray borders that are absent from the presentation.
- e18d19d: Keep table cell text readable when a table or its containing group is flipped, in both browser preview and SVG output.
- e18d19d: Honor table-cell text direction in previews and inline editing. Keep vertical text in place when editing cells with asymmetric margins, including bottom-to-top and upright right-to-left text.
- e18d19d: Display theme-referenced background gradients instead of a solid color, including radial backgrounds selected with the reference desktop app's Background Styles gallery. Resolve gradient colors through the owning slide master's theme and color map while preserving the original theme and background XML on save.
- e18d19d: Resolve shape style effect references against the owning slide master's theme, including placeholder colors and explicit empty effect lists in the layout-to-master cascade.
- e18d19d: Preserve theme-based shape outlines in previews, including line colors and widths from the reference desktop app's Quick Styles. Direct line formatting now retains theme properties that it does not override.
- e18d19d: Add slide-background shape fills through `setShapeSlideBackgroundFill` and the editor's Fill pane, including multiple selection, undo and saved reloads. Fill readers expose the new `background` kind. Preview paints the slide background through these shapes while keeping it aligned through shape and group transforms.
- e18d19d: Keep Japanese punctuation attached to neighboring text when a font or formatting change splits the text into separate runs in SVG previews.
- e18d19d: Support reading and editing source crops on image-filled shapes with the existing image crop functions. Preview cropped image tiles at their cropped size, including mirrored tiles.
- e18d19d: Render DrawingML small caps in SVG previews with reduced-size capitals for lowercase source letters while preserving authored line metrics.
- e18d19d: Read table-cell paragraph and outline-level character defaults with `getTableCellRunFormatEffective`. Preserve inherited fonts, sizes, emphasis, and colors when displaying, editing, inspecting, and copying table text.

  Resolve inherited shape and table text colors through the slide color map, including detached editing previews.

- e18d19d: Render double, dotted, dashed, and heavy underline styles distinctly in previews, including dash-dot and double-wave patterns. Keep strikethrough solid when combined with a patterned underline, and omit spaces from words-only underlines.

  Enable Home font controls for selected table cells, including the Font dialog, while preserving mixed formatting and single-step undo.

- e18d19d: Show imported inner shadows over opaque shapes instead of hiding them behind the shape fill.
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
  - @office-kit/pptx@0.22.0

## 0.12.0

### Minor Changes

- 3c326b7: feat: the SVG ties clicks back to the deck — each slide shape (group members included) is wrapped in a `<g data-pptx-shape-id>` carrying `getShapeId`, paragraphs carry `data-pptx-paragraph`, and table cells carry `data-pptx-cell="row,col"`. Rendering is otherwise unchanged.

## 0.11.0

### Minor Changes

- Updated dependencies [28731c2]
  - @office-kit/pptx@0.21.0

## 0.10.0

### Minor Changes

- 4235418: fix: narrow the `@office-kit/pptx` peer range to `^0.20.0`

  The previous `>=0.13.0` range accepted core versions this renderer no longer
  matches, and it kept every future core release in range so the renderer was never
  republished alongside a breaking one. `@office-kit/pptx` 0.13–0.19 are no longer
  accepted; upgrade the core package alongside this one.

## 0.9.10

### Patch Changes

- a51fa10: fix: highlighted points lost their colors in stacked chart previews (#350)

  Stacked and percent-stacked column and bar previews now preserve per-point highlight colors. They use the same point-color, single-series varying-color, series-color, and palette fallbacks as clustered charts.

## 0.9.9

### Patch Changes

- bb5a337: fix: transparent column and bar series appeared opaque in previews (#349)

  Column and bar previews now respect series fill opacity in clustered, stacked, and percent-stacked charts. Transparent bases remain part of the stack, preserving the position of visible segments in waterfall-style charts.

- 3c9850f: fix: chart previews ignored axis visibility and category tick settings (#351)

  Hiding an axis line now preserves its labels and tick marks, and category axes use the requested major and minor tick marks. Hiding the category axis or its line also removes the zero baseline in column, bar, line, and area charts. Scatter and bubble previews now hide the entire corresponding axis, including its labels and ticks, when the value-axis or category-axis visibility setting requests it.

- 55fc76b: fix: chart previews ignored manually positioned inner plot areas (#348)

  Plotted data now uses the authored inner plot position and size. Negative sizes collapse to zero, and plot rectangles are clipped to the chart frame so malformed layouts cannot produce negative SVG dimensions or extend the plot over neighboring shapes. Outer layouts retain automatic axis gutters.

## 0.9.8

### Patch Changes

- 8a8ac43: Paint an unstyled bullet in its paragraph's first-run colour. Without `<a:buClr>` anywhere in the cascade the renderer fell back to the deck's body-text colour, so a numbered agenda whose runs carry their own light colour drew its `1.` / `a.` markers in the default black and they sank into a dark background. The reference desktop app and LibreOffice take the first run's colour, which is the same rule an un-sized bullet already followed for its size.
- 8a8ac43: Keep a resized group's text at its authored size in the preview. A group whose `<a:ext>` differs from its `<a:chExt>` scales its children, and the renderer was applying that scale to their glyphs as well, so a group squashed vertically (what Google Slides writes for a hand-resized group) drew stretched, half-height letters. The reference desktop app and LibreOffice resize only the geometry, so the text now lays out inside the group-scaled rect at its authored point size and aspect.

## 0.9.7

### Patch Changes

- 978c317: fix: auto-numbered lists restarted at 1 after a nested list

  The preview kept a single numbering counter and reset it on every indent-level change, so a top-level list with nested items between its entries rendered as `1. / a. / b. / 1. / 1.`. The reference desktop app keeps one counter per level: the outer list continues (`2.`, `3.`) and only deeper levels reset when a shallower paragraph starts. A non-numbered paragraph still restarts its own level, and a different numbering scheme at the same level starts over.

## 0.9.6

### Patch Changes

- 71deec5: feat: author every chart type — scatter, bubble, radar, stock, surface, the 3-D variants and pie-of-pie

  `addSlideChart` / `setChartSpec` now write all sixteen plot types of ECMA-376, and `getShapeChartSpec` reads every one of them back.

  - `kind: 'scatter' | 'bubble' | 'radar'` are authorable (they used to throw "read-only"). Scatter and bubble series carry their own `xValues` (and `bubbleSizes`); the embedded workbook lays them out per series so "Edit data" opens onto the right cells.
  - New kinds `'stock'` (3 series: high, low, close; or 4 with open first, drawn as candlesticks) and `'surface'` (`surfaceContour` for the top-down contour plot, `surfaceWireframe` for the mesh).
  - 3-D is a modifier: `view3D: { rotX, rotY, rightAngleAxes, perspective, depthPercent, heightPercent }` on `bar` / `column` / `line` / `area` / `pie` selects the 3-D element, with `bar3DShape` (cylinder, cone, pyramid, …), `gapDepthPct` and `seriesAxis`. `bubble3D` shades bubbles as spheres.
  - `ofPie` turns a pie into a pie-of-pie / bar-of-pie chart.
  - Per series: `errorBars` / `xErrorBars` (fixed, percentage, standard deviation, standard error, custom), `fillOpacity`.
  - Per chart: `dataTable`, `upDownBars`, `categoryAxisDate` (date axis with time units), `categoryGroupLevels` (multi-level category labels), `categoryAxisScaling` (the x axis of scatter / bubble charts, or a date axis' range), `plotAreaLayout` / `titleLayout` / `legend.layout` (manual placement), `valueAxisLineHidden` / `categoryAxisLineHidden`, `valueAxis.displayUnitsLabel`.
  - Data labels gain `showBubbleSize`, `showLegendKey`, `fillColor`, and per-point `text` (a literal label, e.g. naming one scatter point).
  - A spec whose fields contradict each other (`view3D` on a scatter chart, error bars on a pie, a date axis with non-numeric categories, …) throws with a message naming the field, instead of writing a chart the reference desktop app would repair.

  Behavior changes when reading existing decks:

  - A stock chart now reads as `kind: 'stock'` (was `'line'`) and a surface chart as `kind: 'surface'` (was `'column'`). `getPresentationChartKindCounts` has the two new keys.
  - 3-D charts still read as their flat kind, now with `view3D` set, so `setChartSpec(chart, getShapeChartSpec(shape))` no longer flattens a 3-D chart.
  - fix: for scatter / bubble charts `valueAxis` (and the other `valueAxis*` fields) described the **x** axis. They now describe the y axis; the x axis reads into `categoryAxis*` and `categoryAxisScaling`.
  - fix: charts with more than 25 series wrote workbook references past column `Z` as invalid cell addresses.

  `@office-kit/pptx-preview` draws stock charts as lines and surface charts as columns (as before), and honors `categoryAxisScaling` on scatter / bubble charts.

## 0.9.5

### Patch Changes

- a13afc2: fix: bar and column charts ignored per-point colors (`pointColors` / `<c:dPt>`), so a "one bar highlighted, the rest grey" chart rendered in a single color. The preview now paints them, as the reference desktop app does.

## 0.9.4

### Patch Changes

- f7a710f: feat: format the secondary value axis, author per-point data labels, and write mixed-format paragraphs

  - `ChartSpec.secondaryValueAxis` sets scaling, number format, title, label style, gridlines, line color, tick marks, `tickLabelPos` and `crossBetween` on the right-hand axis that `secondaryAxis` series use; `getShapeChartSpec` reads it back. Its colors are validated like the primary axis colors, and its unset label / title colors take the deck's body-text color like the primary axes do, on `setChartSpec` as well as `addSlideChart` (the two now bake unset text colors identically). A chart whose every series is `secondaryAxis` is rejected: at least one series must plot on the primary axis.
  - `ChartSeries.pointDataLabels` authors per-point `<c:dLbl>` overrides (content, format, font) the way pie / doughnut exporters do, and the reader surfaces them. A `<c:dLbls>` holding only per-point overrides no longer reads back as an all-false series-level `dataLabels`. On read, toggles a point override leaves out inherit the series-level values, a deleted point label reads as an override with every toggle off, and overrides past the series' point count are dropped.
  - `setShapeParagraphs(shape, paragraphs)` and `setTableCellParagraphs(cell, paragraphs)` replace a shape's or cell's text with paragraphs that each carry their own alignment and several differently formatted runs; run text is written verbatim. An empty paragraph list is rejected (use `[{ runs: [] }]` for an empty body).
  - `ChartSpec.valueAxisTickLabelPos` positions the primary value-axis tick labels (`none` / `low` / `high` / `nextTo`), mirroring `categoryAxisTickLabelPos`.
  - `ChartSeries.lineColor` sets the series outline / line color separately from its fill (a doughnut's slice borders, for example), and the reader returns the `<a:ln>` color of every series.
  - `ChartSeries.markerColor` / `markerLineColor` set the marker fill and outline colors of a line / scatter / radar series, and the reader returns both. The builder now always writes them (the fill defaults to the series color, the outline to the fill), because the reference desktop app paints a marker without `<c:spPr>` in the theme's automatic color rather than the series color.
  - `@office-kit/pptx-preview` paints chart markers (data points and legend swatches) in `markerColor` / `markerLineColor` and strokes line / scatter / radar series in `lineColor`, instead of always using the series color.
  - Axis lines and gridlines carry a width next to their color: `valueAxisLineWidthEmu` / `categoryAxisLineWidthEmu`, `valueAxisMajorGridlineWidthEmu` / `valueAxisMinorGridlineWidthEmu`, `categoryAxisMajorGridlineWidthEmu` / `categoryAxisMinorGridlineWidthEmu`, and `secondaryValueAxis.lineWidthEmu` / `majorGridlineWidthEmu` (EMU, 12700 = 1 pt). The reader returns them, and a width of 0 — valid in the schema — now reads back on these and on a series' `lineWidthEmu` instead of turning into `undefined`. Without a width the reference desktop app draws these lines at 0.75 pt, so a 1 pt line from another writer used to come back thinner.
  - `ChartDataLabels.showLeaderLines` writes and reads `<c:showLeaderLines>` on series- and chart-level labels (per-point overrides have no such setting).
  - Chart, axis and secondary-axis titles no longer force a 14 pt default size or a horizontal layout: without `sizePt` the application default applies, and a title without a rotation writes neither `rot` nor `vert="horz"` (the reference desktop app reads a lone `vert="horz"` as "not rotated"), so a value-axis title takes the reference desktop app's default (vertical) — pass `0` to keep it horizontal. The reader merges a title's paragraph defaults under its run style.
  - Combo series are shaped by their own `chartKind`: a line series in a column chart keeps its markers and `smooth`, and line / area plot groups write a schema-valid `<c:grouping>` (`standard` instead of the bar-only `clustered`).
  - Horizontal bar charts (`kind: 'bar'`) now put the category axis on the left and the value axis at the bottom, matching the reference desktop app; the hidden companion category axis of a secondary axis follows the primary one.
  - The chart reader honors `<c:ptCount>` (including on multi-level category caches), so series with empty trailing points keep their full length on read-back; cache points past the authored count are dropped.
  - The chart reader identifies the secondary value axis by which plot groups reference it, not by its position, so a scatter chart's X axis at the top no longer reads as secondary and decks that list the secondary axis pair first read back correctly.
  - The XML parser applies XML 1.0 end-of-line handling: a raw CR LF or CR in a part reads as LF, while a `&#13;` character reference still yields a CR.

## 0.9.3

### Patch Changes

- 104f7d2: Read and render the transparency of solid fills and outlines.

  `@office-kit/pptx` adds `getShapeFillOpacity`, `getShapeStrokeOpacity`, and
  `resolveDrawingColorOpacity`, which resolve `<a:alpha>` / `<a:alphaMod>` /
  `<a:alphaOff>` on a color to a 0–1 opacity. `getShapeFillColorResolved` and
  `getShapeStrokeColorResolved` keep returning the plain `#RRGGBB`.

  `@office-kit/pptx-preview` now emits `fill-opacity` / `stroke-opacity` for
  translucent solid fills and outlines, including artwork inherited from slide
  layouts and masters. A semi-transparent shape layered over a gradient or the
  slide background previously rendered as an opaque block that hid whatever sat
  beneath it.

## 0.9.2

### Patch Changes

- f89ac2b: Fix missing EMF artwork and abrupt background color changes in imported slide previews.

  EMF pictures made of solid-filled line and Bézier paths now render as transparent
  vector images, including artwork inherited from slide layouts. Images with
  unsupported drawing commands retain their placeholder instead of rendering only
  part of the artwork. Gradient colors are ordered by their positions before
  rendering, so out-of-order stops no longer introduce flat bands or color jumps.

## 0.9.1

### Patch Changes

- 1020b9a: East Asian line breaking and measured legend packing.

  The text layout engine tokenized wrapped text by whitespace only, so a
  space-free CJK clause travelled as one unbreakable "word" — the greedy
  wrapper pushed the whole clause to the next line, leaving artifacts like a
  lone bullet glyph on its own line. CJK runs now break between any two
  characters with simple kinsoku (closing punctuation glued to its
  predecessor, opening brackets to their successor), matching the reference desktop app's
  East Asian line breaking.

  Chart legends previously packed items into fixed-width slots
  (`min(140px, frameW / n)`), so long CJK series names overflowed into the
  neighbouring item. Horizontal legends ('b'/'t') now pack items by an
  estimated per-label width and shrink the font when the row exceeds the
  frame; vertical legends ('r'/'tr') size the right column to the widest
  label instead of a fixed 100px.

## 0.9.0

### Minor Changes

- 6294c52: Combo charts: per-series `chartKind` overrides and a secondary value axis.

  `ChartSeries` gains `chartKind` (`'bar' | 'column' | 'line' | 'area'`) to
  overlay e.g. a line series on a column chart, and `secondaryAxis: true` to
  plot a series against a right-hand secondary value axis — the standard
  the reference desktop app's combo layout for series with mixed units (counts vs. rates).
  The builder splits series into plot groups (`<c:barChart>` + `<c:lineChart>`
  …) and emits the secondary `<c:valAx>`/`<c:catAx>` pair on demand;
  `getShapeChartSpec` round-trips both fields. The preview renderer paints
  bars below line/area overlays, scales each axis from its own series, and
  draws the secondary ticks on the plot's right edge.

## 0.8.1

### Patch Changes

- 553e3d9: fix(preview): stop shrinking text in shapes without `<a:normAutofit>`

  The reference desktop app never shrinks text in shapes that lack `<a:normAutofit>`:
  `<a:noAutofit>` (or no autofit element) simply overflows the box, and
  `<a:spAutoFit>` grows the box to fit the text. The preview applied a
  heuristic shrink-to-fit to such shapes, which rendered template
  placeholders (font size inherited from the layout/master, box authored
  tightly around the sample text) at down to 0.4× of their size in the
  reference desktop app. The heuristic estimator is removed; only an authored
  `<a:normAutofit>` (with or without a baked `fontScale`) shrinks text,
  matching the reference desktop app.

## 0.8.0

### Minor Changes

- a4f03a4: feat: let `renderSlideToImage` / `renderSlideToRgba` load extra rasterizer fonts

  `RenderImageOptions` gains `fontFiles` (extra ttf / otf / ttc paths handed to
  resvg alongside the bundled Latin faces) and `loadSystemFonts` (opt-in OS font
  fallback, default `false` to keep output deterministic). The bundled set has
  no CJK coverage, so decks with Japanese / Chinese / Korean text previously
  rasterized as missing-glyph boxes with no way to fix it — now callers can
  supply a CJK face, pairing it with `buildFontkitMeasurer({ fonts })` so wrap
  math and painted glyphs use the same metrics.

### Patch Changes

- dc6b2eb: fix: render group children exactly once

  `getSlideShapes` / `getSlideLayoutShapes` / `getSlideMasterShapes` flatten
  group descendants into their result, while `renderSlideToSvg` already recurses
  into groups and draws every child with the group transform applied. Rendering
  the flat list verbatim painted each group child a second time, untransformed —
  visibly offset whenever the group's `chOff` differs from its `off` (common in
  Google Slides exports, e.g. shape-built charts showing doubled axis labels).

  `auditTextLayout` had the same double-enumeration and reported each group
  child's issues twice; it now audits every shape exactly once.

## 0.7.0

### Minor Changes

- 2715289: feat: add `auditTextLayout` — detect text overflowing its box (はみ出し) and unintended soft wraps (段落ち)

  `auditTextLayout(pres, options)` measures every shape's text with the same
  layout engine the preview renders with and reports `overflow-x` / `overflow-y`
  issues (plus opt-in `soft-wrap` reports via `reportSoftWraps`). Results carry
  an `approximate` flag when widths were estimated rather than measured.

  `buildFontkitMeasurer` (the `/node` entry) now accepts `{ fonts }` to register
  the deck's actual font files by their authored family names; registered fonts
  also serve as glyph fallbacks, and glyphs no font covers are estimated
  per-character (CJK ≈ 1em) instead of measured against missing-glyph advances.

## 0.6.2

### Patch Changes

- e09b698: Relax the `engines.node` floor from `>=24.16.0` to `>=22.18.0` on both `pptx-kit` and `pptx-kit-preview` so the maintained LTS lines — Node 22 and Node 24 — are supported, and restore Node 22 to the CI test matrix. The published runtime bundles are unchanged; the previous floor reflected the dev toolchain's pin and needlessly blocked `pnpm install` (under `engine-strict`) on still-supported LTS releases such as Node 22.x and earlier Node 24 LTS patches (e.g. 24.13.x).

## 0.6.1

### Patch Changes

- 4f943a5: Fix several text-on-shape and shape-geometry rendering gaps found by comparing
  output against LibreOffice ground truth on a corpus of realistic, multi-feature
  decks:

  - **Preset pattern fills** (`pct5`–`pct90`, `smGrid`/`lgGrid`, and the
    horizontal/vertical/diagonal hatch families) now match LibreOffice's actual
    substitution for these fills — a density-scaled diagonal hatch for the
    percentage family, and correctly differentiated tile pitches for the "small"
    vs "large" grid variants — instead of a uniformly-dense ordered-dither screen.
  - **Multi-column text bodies** (`numCol` with `noAutofit`/`spAutoFit`) now wrap
    into a new row of columns once the last column also overflows, instead of
    piling all remaining text into the final column forever.
  - **`u="wavy"`** (and `wavyDbl`/`wavyHeavy`) now renders as an actual wavy
    underline in the SVG/raster path (drawn as an explicit path — resvg has no
    `text-decoration-style` support, and the path now scales correctly for a
    superscript/subscript run) and as real CSS in the browser path (without
    also waving a strikethrough on the same run — CSS's `text-decoration-style`
    is a single value for the whole underline + line-through shorthand, and
    the reference desktop app always draws strikethrough solid regardless of underline style).
  - **Rotated + vertically-flipped shape text** no longer renders upside-down;
    the reference desktop app adds a compensating 180° turn to the text specifically for
    `flip.vertical`, independent of `flip.horizontal`. (A shape nested inside a
    vertically-flipped group is not yet covered by this — see the `KNOWN GAP`
    comment on the group-rendering path in `render-slide.ts`.)
  - **Diagonal connectors that need both a horizontal and vertical flip** (e.g. a
    line drawn from bottom-right to top-left) no longer render reversed — the
    flip was applied twice, once via the endpoint swap and again via a leftover
    transform, cancelling itself out and pointing any arrowhead away from its
    target.
  - **Pie/doughnut data labels** now join as `<category> — <value/percent>`
    (e.g. "Web — 48%"), matching the order of the reference desktop app and LibreOffice instead of the
    reverse.
  - **Overlapping (non-stacked) area charts** now paint each series' fill AND
    outline back-to-front as one unit, so the first-authored series stays fully
    on top, matching the reference desktop app; the category-axis title no longer collides
    with the tick-label row.
  - **Table row/column banding** now uses a pale tint (not a near-solid accent
    color), alternates between two tints across every body row (previously every
    other row was left unshaded), and starts the alternation at the first body
    row rather than one row late.

## 0.6.0

### Minor Changes

- e9eae5c: Sharpen preview-renderer fidelity across the sample corpus. Block arrows now use
  the correct OOXML shaft/head proportions (head length scales with `min(w, h)`),
  text in non-rectangular autoshapes (triangle, diamond, pentagon, star, double
  arrow) wraps inside the shape's inscribed text rectangle, vertical text honours
  the rotated text-box insets, glow effects render as a saturated ring instead of
  a pale haze, hyperlink runs take the theme `hlink` colour, bullets size to the
  paragraph's first run and follow centred/right-aligned text, line breaking is
  space-inclusive (matching LibreOffice and the reference desktop app), the first text baseline gets
  the same leading drop for every anchor, and category line charts plot at band
  centres with title/axis text sized in pixels. Overall mean fg-SSIM rises from
  ≈0.82 to ≈0.87.

### Patch Changes

- e9eae5c: Close three renderer correctness gaps surfaced by the expanded corpus:

  - **Preset pattern fills** now render the real ECMA-376 `ST_PresetPatternVal`
    tokens (`horz`/`vert`/`ltHorz`/`ltVert`/`dotGrid` etc.) instead of falling
    through to a 50%-coverage checker — the old matcher keyed on GDI HatchStyle
    names no valid OOXML emits.
  - **`wordArtVert` / `wordArtVertRtl`** stack glyphs upright (one per line) per
    `ST_TextVerticalType`, instead of rotating the run 90°, matching the reference desktop app
    and the browser (`text-orientation:upright`) path.
  - **`<a:normAutofit/>` without a baked `fontScale`** now shrinks text to fit the
    box, so overflowing bodies render at the reduced size the reference desktop app and LibreOffice
    compute at display time rather than spilling past the box. The shrink factor is
    computed once and shared, so the server (SVG) and browser (`foreignObject`)
    previews agree.

- e9eae5c: Fix a crash when rendering a line or connector that sets an explicit line
  cap or join (`setShapeStrokeCap` / `setShapeStrokeJoin`). The renderer
  emitted both its default `stroke-linecap="round"` and the shape's explicit
  cap on the same element, producing a duplicate SVG attribute that aborted
  the render ("attribute 'stroke-linecap' is already defined"). The default
  is now only applied when the shape does not specify its own.

## 0.5.0

### Minor Changes

- 7200690: Drop Node.js 22 support. The minimum supported version is now Node 24.16. The published runtime bundles are unchanged; this only raises the `engines` floor and the CI/test matrix to Node 24.

## 0.4.0

### Minor Changes

- 3ba1e3d: Drop Node.js 20 support. The minimum supported version is now Node 22.18. The build toolchain moved to tsdown, whose current release requires Node 22.18+; the published runtime bundles are unchanged.

## 0.3.2

### Patch Changes

- 333b19f: fix: chart rendering now matches the reference desktop app. The renderer drew an invented
  light-gray chart-area frame, omitted axis spines, used faint inward tick stubs,
  defaulted value-axis gridlines on, rendered every line/scatter marker as a
  circle, and drew bar charts with the category axis upside-down. Now:

  - The chart-area border is drawn only when the chart authors one.
  - Value and category axes draw their spine and outward major tick marks.
  - Major gridlines render only when authored (`<c:majorGridlines>`).
  - Line / scatter / radar markers follow the reference desktop app's automatic symbol
    rotation (diamond, square, triangle, x, …) when no symbol is authored.
  - Bar charts order categories bottom-to-top, matching the reference desktop app.

- 333b19f: fix: percentage pattern fills (`pct5`…`pct90`) now render at the requested
  coverage. They were drawn as a sparse 1–4 dot grid that read far too light —
  `pct50` looked like ~5% ink instead of a 50% screen. They now use an ordered
  (Bayer) dither so the tone matches the reference desktop app.

## 0.3.1

### Patch Changes

- 0f7c538: Preview: take the default text color from the deck's body style, not the `tx1` token

  The preview used `scheme:tx1` as the fallback color for runs without an authored color. On a template with an inverted color map (`tx1 → lt1`) that resolves to the light slot, so body text was painted white on the white background — the whole slide looked blank. The reference desktop app instead takes the fallback from the master `bodyStyle` (e.g. `schemeClr bg1`). The preview now does the same via the newly exported `resolveDeckBodyTextColor(slide)`, so default-colored text and table-cell text resolve to the color the reference desktop app actually paints.

  - New export **`resolveDeckBodyTextColor(slide)`** — the deck's resolved body-text color (master `bodyStyle`, run through the effective color map + theme). This is the color `addSlideTable` / `addSlideChart` bake in, now reusable by renderers.

## 0.3.0

### Minor Changes

- 4a2ede1: Resolve scheme colors through the slide's color map so inverted-map templates render correctly

  Templates whose slide master inverts the color map (`<p:clrMap bg1="dk1" tx1="lt1">`, common in Google Slides / Canva exports) previously rendered with swapped light/dark colors: slide backgrounds came out black in the preview while the reference desktop app paints them white, and generated tables and charts came out with invisible text (the default `tx1` token resolved to the same color as the background).

  - **`getEffectiveColorMap(slide)`** — new export returning the slide's effective color map (the master's `<p:clrMap>` overlaid by a per-slide `<p:clrMapOvr>`). Color resolution and renderers apply it to `schemeClr` tokens before indexing the theme.
  - **`resolveDrawingColor(colorEl, theme, clrMap?)`** — accepts an optional color map; scheme tokens are remapped through it before the theme lookup. Omitting it preserves the previous behavior (correct for the standard map).
  - **`addSlideTable` / `addSlideChart`** now bake the deck's resolved body-text color onto table cells and chart text (axis labels, legend, data labels) so generated tables and charts stay readable regardless of the template's color map. Authored colors still win; override table cells afterwards with `setTableCellTextFormat`.
  - **`pptx-kit-preview`** resolves `schemeClr` tokens through the effective color map, so previews of inverted-map decks match what the reference desktop app paints.

## 0.2.0

### Minor Changes

- b03e0cb: feat: fidelity calibration sweep — measured against LibreOffice ground truth,
  mean fg-SSIM rose from ≈0.66 to ≈0.78 (≈0.81 excluding documented
  divergences). Body placeholders now inherit the master `bodyStyle` bullet and
  hanging indent through the paragraph cascade (new `bullet` field on
  `ParagraphProperties` from `getParagraphPropertiesEffective`); charts no
  longer invent a legend when the XML authors no `<c:legend>`, and the value
  axis gets spreadsheet-style headroom above the data max with the tick step
  preserved; the chart builder writes `<c:smooth val="0"/>` explicitly on line
  series (the schema default for an absent element is smooth=1, which made
  LibreOffice draw unauthored lines as curves); and the pure-SVG text layer is
  nudged 0.75px left to land on LibreOffice's pixel grid.

## 0.1.0

### Minor Changes

- e65228f: feat: read custom geometry. New `getShapeCustomGeometry(shape)` returns a
  shape's `<a:custGeom>` (ECMA-376 §20.1.9) as a fully-evaluated path list —
  guide formulas (`avLst`/`gdLst`, all §20.1.9.11 operators) are resolved
  against the shape extents so the returned `moveTo`/`lnTo`/`arcTo`/
  `quadBezTo`/`cubicBezTo`/`close` commands carry only numbers. The preview
  renderer now draws custom geometry as a real SVG path (arcs converted to
  cubic Béziers) instead of a labelled rectangle placeholder; only a custGeom
  that fails to evaluate still falls back, marked `data-pptx-fallback`.
- acc9b15: feat: effects & fills polish. The reflection effect (`a:reflection`) now
  renders as a vertically mirrored, gradient-masked copy honoring start/end
  alpha, distance, and the signed `sy` scale; picture bullets (`a:buBlip`)
  render as real inline images in both text layout modes via the new core
  reader `getParagraphBulletImageBytes` (the "■" fallback remains only when
  bullet bytes are genuinely unavailable); and gradient fills inherited
  through the placeholder layout/master cascade resolve via the new
  `getShapeGradientFillEffective` instead of painting a hardcoded orange
  tint.
- 008e7c1: feat(preview): vertical text (`vert`, `vert270`, `eaVert`, `mongolianVert`,
  `wordArtVert`) and multi-column bodies (`numCol`/`spcCol`) now render in the
  pure-SVG text mode (`textLayout: 'svg'`) used by server-side rasterization.
  Previously they fell back to horizontal single-column layout; now the server
  output matches the browser (`foreignObject`) path: rotated line stacking for
  vertical text and desktop-app-style sequential column fill for multi-column
  bodies. Browser rendering is unchanged.
- 2207ed1: feat: scatter, radar, and bubble charts are now modeled as their own
  `ChartKind`s instead of being folded into `line`. `ChartSeries` gains
  `xValues` (`<c:xVal>`) and `bubbleSizes` (`<c:bubbleSize>`); `ChartSpec`
  gains `scatterStyle`, `radarStyle`, `bubbleScale`, and
  `bubbleSizeRepresents`. Read + render only: the preview draws real
  scatter (two value axes + markers), radar (polar spokes/rings), and
  bubble (area-proportional circles) plots, and the write path now rejects
  these kinds loudly — previously a read-modify-write silently corrupted a
  scatter chart into a line chart.

### Patch Changes

- ba94f5e: fix(preview): table cell text now renders with its real per-run formatting —
  font size, bold, italic, color, typeface, and paragraph alignment — and wraps
  within the cell width, in both the browser (`foreignObject`) and server
  (`svg`) text modes. Previously every cell was drawn at a flat 18 pt with no
  styling. Cells with no explicit run size still fall back to the reference desktop app's 18 pt
  default in the theme's body font.
