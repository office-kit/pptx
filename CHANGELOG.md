# @office-kit/pptx

## 0.23.0

### Minor Changes

- e860ae8: Preset shapes are now drawn from ECMA-376's own preset definitions instead of hand-written approximations. The preview, the editor canvas and its Shapes gallery icons showed many presets wrongly — hearts, lightning bolts, suns, moons, clouds, brackets and braces, bent and curved arrows, the equation shapes, stars with 7 to 32 points — and now match PowerPoint, including the lit and shaded faces of cubes, cans and curved arrows, adjust handles on every preset, and elliptical arcs in custom geometry (whose `arcTo` angles were read as parametric instead of visual angles). Text in a preset now wraps inside the rectangle the definition gives it, such as an ellipse's inscribed rectangle, as PowerPoint does.

  - `@office-kit/pptx`: new `getPresetGeometry(preset, size, adjustValues?)` evaluates any preset's paths and text rectangle in the same form as `getShapeCustomGeometry`.
  - `@office-kit/pptx-preview` (breaking): `shapeCustomTextRect(custom, extent)` is replaced by `shapeTextRect(shape)`, which reads custom and preset geometry alike, and `resolveTextBodyRect` now takes `(bounds, margins, region)` with the region from `shapeTextRect`.

- 089140c: The preview now draws text bevels (`<a:bodyPr>` `<a:sp3d><a:bevelT>`, as PowerPoint's Soft Bevel and Sharp Bevel WordArt write them): beveled glyphs are shaded with a light and dark edge sized from the bevel's width and height and lit from the light rig's direction, in shape text, table cells and the editor while text is being edited. The 12 bevel presets get different edge profiles and shiny materials add a highlight; this is a 2-D approximation, so the camera, extrusion and contour are not drawn. Text without a bevel renders exactly as before. New `getTableCellText3D` reads a table cell's text 3-D, like `getShapeText3D` does for shapes.

### Patch Changes

- 8db2360: `setShapeText`, `setTableCellText` and `setSlideNotes` with a `range` no longer leave a separate `<a:r>` for every insertion: text inserted next to a run with the same formatting extends that run, and deleting the text between two halves of a run joins them again, as PowerPoint does. Typing in the editor previously saved one run per keystroke.

## 0.22.0

### Minor Changes

- e18d19d: `getSlideAnimations` reads what a timing tree says becomes of an effect's value once it has run. Each step reports it as `valueAfterEnd`: `'held'` when every behaviour that animates something carries a `fill` that keeps its value, `'removed'` when one says it is taken away again, and `'unstated'` when the tree does not say — the schema gives `fill` no default, so an absent one is not the same as `'hold'`.

  A step is now `playable` (and so `editable`) only when the tree states that the effect leaves the shape where it puts it: the effect node's own `fill`, and that of the `<p:set>` that flips visibility — or, for an effect that only fades, the fade's. An effect that says its value is taken away there, or says nothing there, is still listed and still saved exactly as it arrived, but is reported as one this library reads rather than plays.

  `valueAfterEnd` itself decides nothing: an effect whose fade is removed or unstated still plays, because on its own the shape ends up where its own visibility puts it. It is what a player needs when two effects animate one object at the same time and the shorter one ends first — whether its value stays above the longer one or gives way to it. Decks this library authored are unaffected: it has always written `fill="hold"`.

- e18d19d: `setShapeAnimation` and `updateSlideAnimation` take five more effects, and
  `getSlideAnimations` reads them back: `flyIn` / `flyOut` travel in from, or out
  through, one edge of the slide, `zoomIn` / `zoomOut` grow from nothing and
  shrink back to it about the shape's centre, and `spin` turns the shape one
  clockwise turn.

  `spin` is the first emphasis effect. It animates a shape that is already on the
  slide and leaves it exactly where it was, so unlike an entrance or an exit it
  never decides whether the shape is shown — a shape no entrance has revealed
  stays unrevealed while it turns.

  A fly takes a `direction`: `'top'`, `'right'`, `'bottom'` (the default) or
  `'left'`, naming the edge it comes from or leaves by. Steps report it as
  `direction`, `null` for every effect that does not fly, and passing it for one
  of those is an error rather than a silently dropped field.

  An effect is only named when the tree states what it does. A preset is matched
  on all three of `presetClass`, `presetID` and `presetSubtype`, so an imported
  entrance that shrinks in from four times its size is no longer read as one that
  grows from nothing; and a rotation is only reported as `spin` when it really is
  a single full clockwise turn, so a half turn or one the other way is listed and
  saved as it arrived rather than rewritten as a full one.

  Fixes: an exit fade hid its shape the moment it started instead of fading it
  out. The `<p:set>` that takes the shape off the slide now trails the motion the
  way PowerPoint writes it, and moves when the effect's duration changes.

- e18d19d: `setShapeAnimation` accepts `byParagraph`, which reveals a shape's text one paragraph at a time instead of animating the shape as a whole — PowerPoint's and Google Slides' "By paragraph". Each paragraph gets its own effect with the `start` you asked for, so the default advances a paragraph per click, and they share the one `<p:bldP build="p">` that makes PowerPoint treat them as a single build. `getSlideAnimations` reports each paragraph as its own step, targeting a paragraph range. A shape with no text is refused rather than silently animated as a whole, and a call that cannot be completed leaves the slide's existing timing exactly as it was.
- e18d19d: `setShapeAnimation` accepts `start` and `delayMs`. `'withPrevious'` runs an effect alongside the one before it and `'afterPrevious'` once that one has finished, both off the click that started their predecessor, so several shapes can animate from a single click. As a slide's first effect neither has a predecessor to follow, so both run as the slide appears.

  `'afterPrevious'` needs to know when the effect before it ends. On a slide whose timing states that in a form this library does not model — an effect that runs indefinitely or states no duration, one that repeats or is rescaled, or one that starts from another node rather than at a fixed offset — the call throws and leaves the slide's timing untouched rather than placing the effect at a guessed moment. `'click'` and `'withPrevious'` need no such measurement and still work there.

  Existing calls are unaffected: the defaults still write the click-triggered, zero-delay tree they wrote before, byte for byte.

- e18d19d: Apply the active slide's background to the whole presentation with `applySlideBackgroundToAll` or the editor's Apply to All button. Backgrounds move to slide masters and individual slide/layout overrides are cleared, preserving theme colors and image references. The editor supports undoing the whole operation together.
- e18d19d: Add `setSlideMasterBackgroundStyle` to select PowerPoint's twelve theme background styles. The selected slide's master receives the theme fill reference and light or dark text/background mapping, preserving accent colors and existing slide overrides.
- e18d19d: Read and edit PowerPoint’s Rewind After Playing setting for audio and video, preserving it when saving and reopening presentations.
- e18d19d: Preview background images with their stretch offsets, tile scale, alignment, offsets, and alternating reflections, including inherited backgrounds. Expose natural background image dimensions using embedded resolution and fill DPI, so tiling matches the image size instead of stretching it across the slide.
- e18d19d: Read and change slide background image placement with `getSlideBackgroundImageFillLayout` and `setSlideBackgroundImageFillLayout`. Stretch offsets and tiled scale, position, alignment, and mirroring preserve the image data. Editing inherited images creates a slide override without changing other slides.
- e18d19d: Render background picture cropping for stretched and tiled images, including inherited backgrounds. Add `getSlideBackgroundImageCrop` to read the effective source rectangle without changing the deck.
- e18d19d: Edit slide background patterns with a pattern gallery and foreground/background theme color controls. Selected slides support undo, reset and saving. Editing an inherited pattern creates a slide override while preserving the layout or master and the original color transforms.

  Background pattern readers accept `preserveTheme` to retain unmodified theme color references.

- e18d19d: Add `setSlideBackgroundPatternFill` to edit background patterns and colors. Partial edits preserve unspecified pattern colors and their original DrawingML transforms.
- e18d19d: Restore picture backgrounds when returning from another fill type during editing, preserving placement, transparency, cropping and imported effects. Add `copySlideBackground` to copy an effective background between slides or presentations while retaining its XML and referenced media.
- e18d19d: Read and edit slide background picture opacity, including inherited backgrounds. Format Background now provides a transparency slider and percentage field, and the preview displays the selected opacity.
- e18d19d: Add a Background Styles gallery to the editor's Design ribbon, with theme color previews, master-wide application, undo, and links to background formatting and reset. Expose `getSlideMasterBackgroundStyles` to read the available presets and master selection without modifying a presentation.
- e18d19d: Insert several slides at one position with `addSlideAt` and distribute formatted text ranges to multiple shapes with `setShapeParagraphs`. Batch text copies preserve paragraph properties, fields and hyperlinks, including when a destination is also the source. Slide insertion rejects layouts from another presentation before changing the deck.

  Read all paragraph inline elements in one pass by omitting the paragraph index from `getShapeParagraphElements`.

- e18d19d: Allow setShapeTextFormat to target a UTF-16 text range while retaining surrounding
  run and paragraph properties. Reject invalid ranges without changing the document.

  Add a bilingual selected-text formatting bar to the development preview with
  bold, italic, underline, fonts, size and color, backed by persisted edits and undo.

  Keep property controls synchronized with canvas edits and preserve mixed text
  formatting when editing in the properties pane. Build the preview editor into
  the development package regardless of the working directory.

- e18d19d: Support `TextFormat.underlineColor` for independently colored underlines. Set it to `null` to follow the text color. Preserve the color when reading, editing, copying, and saving text, and expose underline color and Automatic in the editor's Font dialog.
- e18d19d: Support reading and changing transparency on image-filled shapes as well as pictures, including preview rendering. Invalid opacity values now leave the existing transparency unchanged.
- e18d19d: Support PowerPoint's Top, Middle and Bottom Centered text anchors through the existing `setShapeTextAnchor` API's `centered` option. Resolve inherited centering through `getShapeBodyPrEffective`, preserve paragraph alignment, and expose all six anchor choices in the editor with preview and text-editing support.

  Preview integrations can use `shapeTextAnchorOffset` to position editable text consistently with the rendered text block.

- e18d19d: Text can now carry its own outline, shadow and glow — the WordArt half of a
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

- e18d19d: Add Clear text formatting to the preview's English and Japanese text and table-cell toolbars, with Ctrl/Command+Backslash for selected text. Clearing preserves hyperlinks and paragraph settings and supports Undo/Redo.

  The existing `setShapeTextFormat` and `setTableCellTextFormat` APIs now accept `{ reset: true }` to restore inherited run appearance before applying a new format, optionally limited to a character range.

- e18d19d: Add comment replies through `addSlideComment({ replyTo })` and `getCommentParent`, preserving PowerPoint p15 threading extensions through save and reload. Removing a comment also removes its descendant replies.

  The preview comments dialog supports creating and editing replies in English and Japanese, shows each reply's parent, and deletes threads with undo/redo support.

- e18d19d: `setShapeShadow` and `setShapeGlow` no longer erase each other. Each now
  replaces only the effect of its own kind and leaves the rest of the shape's
  `<a:effectLst>` in place, written in the order `CT_EffectList` states — so a
  shape can carry a shadow and a glow at once, as PowerPoint routinely writes
  them. Setting the same effect twice still replaces it, because the schema
  allows each one only once. `clearShapeEffects` remains the way to empty the
  list.

  `getShapePatternFill` now reports `preset` as the `PatternPreset` union rather
  than a bare `string`, so what it reads can be handed straight back to
  `setShapePatternFill`. An unrecognised `prst` in the file reads as `'pct50'`,
  which is what PowerPoint paints when the attribute is absent.

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

- e18d19d: Add `setShapeCustomGeometry(shape, { paths })`, the writer counterpart of `getShapeCustomGeometry`. It replaces a shape's preset or custom geometry with `<a:custGeom>` paths (moveTo, lnTo, arcTo, quadBezTo, cubicBezTo, close), rounding coordinates to whole numbers and rejecting non-finite values, empty paths and paths that do not start with moveTo before touching the shape.
- e18d19d: Add custom show destinations to shape and text-range click actions.
- e18d19d: Add public Custom Show APIs for reading, replacing, and validating named slide sequences, including repeated slides and safe cleanup when a slide is removed.
- e18d19d: Set slide background gradients with `setSlideBackgroundGradientFill`, including theme colors, stop opacity and brightness, and linear or radial direction settings. Invalid gradient settings leave the previous background unchanged.
- e18d19d: Edit individual objects inside groups in the development preview, with English and Japanese group navigation, transformed pointer/keyboard movement, and undoable deletion and duplication.

  Allow `copyShape` to preserve ancestor group transforms with `preserveGroupTransform`. Deletion and z-order operations now handle nested shapes in their owning group and preserve trailing extension metadata.

- e18d19d: Animations on a slide can now be edited, not only added. `updateSlideAnimation` changes an effect's preset, duration, delay, start condition or paragraph build; `removeSlideAnimation` takes one out; `moveSlideAnimation` changes its place in the click order. Every effect is addressed by the `<p:cTn id>` `getSlideAnimations` reports, and that id survives each call — including a `byParagraph` toggle, where the addressed effect keeps its place and the rest of the build is new. Reordering and retiming recompute the offsets that make `afterPrevious` mean it, so the effects after the one you changed still start when the one before them ends, and an effect that starts as the slide appears is never quietly turned into one that waits for a click.

  Copying a shape now brings its animations across as the effects they are: a build with a paragraph deleted stays deleted, per-paragraph timings survive, and a preset this library can only read comes along rather than being dropped. Removing a shape takes its effects with it and moves up whatever started when they ended.

  An edit, copy or removal this library cannot reproduce faithfully is refused before anything changes, with a message naming what is in the way — a `<p:tn>` condition that would be left waiting on nothing, an effect shared with a shape that is staying behind, an interactive sequence, or a main sequence laid out in a way it could not write again. In particular, `copyShape` refuses before adding the shape, its relationships or its parts, and `removeShape` refuses before taking the shape out, so the presentation is left exactly as it was.

- e18d19d: Expose inherited bullet color, size, and font details through `getParagraphPropertiesEffective`.

  Preserve bullet formatting inherited from layouts and masters in previews and while editing text, including explicit follow-text overrides.

- e18d19d: Allow `setShapeTextFormat` and `setTableCellTextFormat` to target a paragraph end mark with `paragraphEnd`, preserving existing runs and paragraph properties.

  Fix Increase/Decrease Font Size leaving empty paragraphs unchanged when formatting an entire text box or table cell.

- e18d19d: Read and write PowerPoint's Equalize character height setting with `TextFormat.normalizeHeight`, preserving explicit enabled and disabled values through editing and saving.
- e18d19d: Add an optional proportional content fit when changing slide size, with English and Japanese page setup controls. Scale object coordinates, explicit text formatting, table dimensions and outlines, and center the original page on the new canvas. Invalid scaled dimensions leave the presentation unchanged.
- e18d19d: Read and update slide show mode, slide range, looping, narration, animation, and timing settings while preserving other presentation properties.
- e18d19d: Preserve gradient stop transparency when reading shapes and expose theme-resolved stop colors. Shape previews now render the brightness and transparency saved by PowerPoint for linear and radial gradient stops.
- e18d19d: Allow gradient stops to specify brightness and opacity, and gradients to preserve scaling and rotation settings. Reject invalid stop edits without discarding the shape's existing fill.

  Edit gradient stop color, position, brightness and transparency in the Format Shape pane, with add/remove controls and Undo support. Drag stops directly, commit each drag as one undoable edit, and cancel an in-progress drag with Escape.

- e18d19d: Expose gradient tile bounds through `GradientFillOptions.tileRect` and gradient readers. Preserve PowerPoint's corner-gradient tile bounds when changing stops or rotation and saving the presentation.
- e18d19d: Add image-fill placement readers and setters for tile alignment, scale, offsets, mirroring, stretch offsets and rotation with the shape. Placement changes preserve the embedded image, crop and effects. Preview now renders stretch offsets and clips the image to the shape.
- e18d19d: Add `isSlideBackgroundGraphicsHidden`, `setSlideBackgroundGraphicsHidden`, and `isSlideLayoutBackgroundGraphicsHidden` to inspect and control inherited decoration without deleting template content. Apply to All also copies the graphics visibility setting to slides and layouts. The preview honors both levels, and the editor's Format Background pane supports changing selected slides with undo and save/reload.
- e18d19d: Fix horizontal text overrides on placeholders inheriting vertical text from a layout or master. Explicit `horz` values now stop the inheritance cascade. Behavior change: `setShapeTextDirection(shape, 'horz')` writes an explicit override; use `null` to clear the override and restore inheritance.
- e18d19d: Add `setShapeImageRecolor` for PowerPoint-compatible grayscale, duotone, threshold, and washout image effects.
- e18d19d: Allow effective run formatting to resolve layout and master inheritance from an original shape while reading a detached editing preview. Show inherited character styles during inline shape editing, including pending text and paragraph changes, without baking those styles into saved text or clipboard data.
- e18d19d: Resolve inherited text autofit and columns through `getShapeBodyPrEffective`, and use them consistently in previews and editing controls. Allow `setShapeTextColumns` to author one column explicitly, overriding inherited columns; invalid column settings leave the previous values intact.
- e18d19d: Allow effective paragraph properties to retain original placeholder and slide inheritance when reading a detached shape preview. Display paragraph alignment, line spacing, spacing before and after, indentation and text direction during inline editing, with canvas zoom applied to absolute dimensions.
- e18d19d: Add character-level inner shadow formatting to the text API, including PowerPoint OOXML read, write, inheritance, and independent removal.
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

- e18d19d: Resolve line-break formatting with `{ breakIndex }` in getShapeRunFormatEffective and getTableCellRunFormatEffective. Preserve inherited font, size, and emphasis when selecting, editing, or copying line breaks.
- e18d19d: Add PowerPoint link destinations for the last slide viewed and ending the slide show.
- e18d19d: `SlideShowProperties` gains an optional `showMediaControls` (PowerPoint's Show Media Controls, stored as `p14:showMediaCtrls`); `getSlideShowProperties` always reports it. The editor's Slide Show tab adds Rehearse Timings (time each slide while presenting, then keep the times as slide timings), Record, and the Use Timings, Play Narrations and Show Media Controls options; the Microsoft 365-only commands are shown disabled.
- e18d19d: Add read and write support for named bookmarks in embedded media playback.
- e18d19d: Say how a clip plays: `getShapeMediaPlayback` and `setShapeMediaPlayback`.

  A deck could embed a video or a sound but not state anything about playing it,
  so every clip waited for a click at PowerPoint's default volume. The new pair
  reads and writes autoplay, loop, volume, mute, hide-when-stopped and (video
  only) full screen, from the clip's media time node. Omitted properties keep
  their current value.

  Trimming a clip is still not supported: PowerPoint stores it in a 2010
  extension rather than in the core schema.

  The editor exposes the settings through its properties panel and command
  palette, in English and Japanese.

- e18d19d: Read and edit the OOXML `numSld` setting for audio and video playback across slides.
- e18d19d: Preserve automatic media start delays through `MediaPlayback.delayMs`. Presentation preview now waits for the saved delay before playing audio or video and cancels pending playback when leaving a slide or starting media manually.
- e18d19d: Comments written by PowerPoint 2021 and Microsoft 365 are now read and edited,
  not merely carried along. Those decks keep their comments in a different format
  from the one ECMA-376 defines ([MS-PPTX] §2.16.1): a thread owns its replies
  instead of the replies pointing back at a parent, the text is a DrawingML body
  rather than a string, authors are identified by GUID in their own
  `/ppt/authors.xml`, and — the part that has no equivalent at all in ECMA-376 — a
  thread can be **resolved**.

  `getSlideComments` returns both kinds. `getCommentFormat` says which one a
  comment came from, `getCommentStatus` reports `'active'`, `'resolved'` or
  `'closed'`, and `setCommentStatus` resolves a thread or reopens it. Asking for
  a status on an ECMA-376 comment gives `null`, and setting one throws rather than
  quietly doing nothing: that file has nowhere to keep it.

  `setCommentText`, `removeSlideComment` and `clearSlideComments` work on either
  format. `addSlideComment` follows the file rather than the caller — a reply
  joins its own thread, and a new thread goes where the slide already keeps its
  comments, so a deck that has never had one still starts an ECMA-376 list that
  every reader understands. Everything the library does not model is left exactly
  where it was: the slide or shape anchor, extension lists, reactions, and the
  pin.

  Fixes: a comment added without a pin produced invalid XML. `<p:pos>` is
  required by `CT_Comment`, and nothing in the API made a caller pass one — so
  every comment the editor added was rejected by the schema. Such a comment is now
  pinned to the slide's origin, and `getCommentPosition` reports that rather than
  `null`. Comments read from a file that omits the element still report no
  position, and are still saved the way they arrived.

- e18d19d: Move all selected objects together with Bring to Front, Send to Back, Bring Forward and Send Backward, preserving their relative stacking order and supporting a single undo step in the editor.

  The four existing stacking APIs also accept an array of sibling shapes. Selections from different parent containers are rejected before any shape is moved.

- e18d19d: Preserve shadow alignment and rotation when reading and editing PowerPoint text. Shadow formatting now accepts `alignment` and `rotateWithShape`, including native WordArt anchors.
- e18d19d: Add character-level reflection formatting to `TextFormat`, including native PowerPoint read, write, merge, removal, and round-trip support.
- e18d19d: Group sibling shapes and ungroup groups inside an existing group without removing the outer group or changing its transform. Grouping preserves the members' stacking order regardless of selection order. Preview group commands support nested selections with Undo/Redo and reject selections spanning different parents.
- e18d19d: Format selected speaker-note text from the Home ribbon, preserving mixed formatting when editing and saving notes. Add APIs to read and update notes character formatting and change letter case.
- e18d19d: Read and update paragraph levels across a UTF-16 text selection with `getParagraphLevel` and `setParagraphLevel`. Relative updates preserve differences between levels and keep text, formatting and links intact. Indenting a long outline selection now updates its text body once instead of once per paragraph.
- e18d19d: Persist outline collapse with `getCollapsedOutlineSlides` and `setSlideOutlineCollapsed`, preserving PowerPoint's per-slide view state. Deleting a slide removes its outline reference.

  In the editor, double-click a slide icon in Outline View to collapse or expand its body. The right-click menu also collapses or expands the selected slides or the whole outline. Each change supports Undo and survives saving and reopening.

- e18d19d: Allow `addSlidePlaceholder` to inherit a placeholder directly from the slide master without changing the slide layout.

  The editor now confirms deletion of additional objects when demoting an outline title, and restores a missing body placeholder even on Title Only slides. Cancel keeps the document unchanged; Undo restores the deleted slide and its objects.

- e18d19d: Move selected outline body paragraphs up or down from the context menu, retaining paragraph formatting and hyperlinks through save and Undo. The existing `setShapeParagraphs` API now accepts ordered source ranges for a single destination.
- e18d19d: Allow `setShapeZIndex` to insert an ordered batch of sibling shapes in one update, preserving non-shape XML and rejecting mixed containers before mutation.

  Add a layer preview for Arrange > Reorder Overlapping Objects. Drag or use arrow keys to stage an order, cancel without changing the document, or confirm with one undoable change.

- e18d19d: Add `setParagraphIndent` to edit left, right, and first-line indentation in text boxes and table cells, or remove individual overrides to restore inherited values.

  Expose paragraph indentation in the editor command palette.

- e18d19d: Add `setParagraphTabs` to edit custom tab positions, left/center/right/decimal alignment, and automatic tab spacing in text boxes and table cells. Effective paragraph properties now include inherited tab settings. Empty tab lists clear custom stops, while null removes a local override.
- e18d19d: Read and update paragraph Asian line breaking, Latin word wrapping, hanging punctuation, and font alignment with `setParagraphTypography`, including inherited settings and removal of local overrides.

  Add the Paragraph dialog's Line Breaks and Alignment tab. Apply changes to selected paragraphs or table cells together, preserve mixed values, and support undo and redo. These settings are saved in exported PowerPoint files; preview rendering does not yet reproduce all typography rules.

- e18d19d: Choose pattern foreground and background colors from theme and standard palettes. Base theme colors stay linked when switching fill types and after saving.

  `getShapePatternFill` accepts `preserveTheme: true` to return untransformed theme references; its default continues to return resolved RGB colors.

- e18d19d: Read and update shape aspect-ratio constraints with `isShapeAspectRatioLocked` and `setShapeAspectRatioLocked`. The editor now preserves the saved Lock aspect ratio setting when resizing videos and using the size pane.
- e18d19d: Add setShapePreset and read picture presets through getShapePreset. Clip picture previews to their preset shape while preserving source-image crop and effects. Add bilingual image-shape controls for ellipse, rounded rectangle, triangle, diamond, pentagon, hexagon, star, and heart masks, with undo and saved persistence.
- e18d19d: Read and edit presentation drawing guides, guide visibility, grid spacing and grid snapping while preserving unrelated view settings.

  The editor adds grid settings, draggable and colored guides, guide/grid snapping, Slide Sorter, a View ribbon, resizable slide thumbnails and zoom controls. Edit speaker notes directly below the slide, with autosave and undo/redo. Display preferences remain separate from presentation edits.

  Keep View submenus inside the window, including when space at the right edge is limited.

- e18d19d: Allow `getShapeImageDuotone` to preserve theme references and color transforms with `resolveColors: false`. Keep two-color image recoloring when switching a shape's fill away from an image and back in the editor.
- e18d19d: Preserve mixed text formatting during inline preview edits. `setShapeText` now accepts `preserveFormatting` for incremental changes, retaining unchanged runs, paragraph properties, fields and hyperlinks.
- e18d19d: `setShapePatternFill` now accepts partial settings, so changing a pattern preset or one color preserves the other color's theme reference and imported transforms. Invalid presets no longer remove the existing fill before throwing.
- e18d19d: Preserve existing speaker-note run formatting and fields when editing notes in the editor. `setSlideNotes` now accepts optional `range` and `preserveFormatting` options for targeted text edits; its default replacement behavior is unchanged.
- e18d19d: Add a bilingual find-and-replace dialog to the development preview, with match navigation, individual and bulk replacement, case sensitivity, current-slide scope, table cell support, keyboard shortcuts and undoable changes. Replacement values are literal text.

  Add an optional UTF-16 `range` to `setShapeText` and `setTableCellText` to replace an exact selection while retaining unaffected formatting, including when adjacent characters are identical. Invalid boundaries and split surrogate pairs are rejected.

- e18d19d: Add a canonical `setShapeStyle` API for applying native PowerPoint shape style references while preserving shape geometry and text formatting.

  Allow `setShapeHidden` to update multiple shapes in one call while preserving animation data.

- e18d19d: Render PNG and JPEG image fills as repeating tiles with alignment, offsets, independent scaling and alternating horizontal/vertical reflections. Use embedded PNG/JFIF resolution or an explicit fill DPI to size the tiles. Add `getShapeImageIntrinsicSize` to read the image's unscaled physical size.
- e18d19d: Add optional speaker-notes color resolution through the notes master theme and color map. The notes editor uses it for display while retaining literal color references for editing and round trips.

  Expose notes line-break metadata and paragraph/break-aware editing so soft breaks survive save/load and undo.

  Keep the notes caret after inserted paragraph breaks and preserve theme-based typing colors across consecutive empty paragraphs.

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

- e18d19d: `resetSlidePlaceholderTextFormatting` and `resetSlideLayout` now reach
  placeholders inside a group. A group scales and turns what is inside it; it
  does not decide what font the text is in, so a grouped placeholder inherits its
  layout's formatting and appearance like any other. Both functions count those
  placeholders in what they return.

  `resetSlidePlaceholderGeometry` still leaves them where they are, and so does
  the geometry half of `resetSlideLayout`. The layout states a rectangle on the
  slide, while a grouped shape's geometry is written in its group's coordinate
  space — restoring it would either tear the shape out of the arrangement it was
  grouped into, or invent a rectangle the layout never described.

- e18d19d: Add `resetShapeImageColorEffects` for clearing PowerPoint image recolor and brightness/contrast effects while preserving other image effects.
- e18d19d: Add resetSlidePlaceholderGeometry to restore top-level placeholder position, size, rotation and flips from the current layout or master while preserving content and formatting. The bilingual development editor exposes it for individual or selected slides with undo/redo.
- e18d19d: Add `resetSlidePlaceholderTextFormatting` to restore inherited text, paragraph and text-body formatting on layout-bound placeholders while retaining content, hyperlinks, language and outline levels. Expose the operation in English and Japanese in the preview, with selected-slide support and undo/redo.
- e18d19d: Add `resetSlideLayout` to restore missing placeholders, layout geometry and inherited shape/text formatting together while preserving content and image relationships. Expose Reset layout in the English and Japanese preview, with multi-slide selection and a single undo step.
- e18d19d: Add `addMissingSlidePlaceholders` to restore deleted layout slots without changing existing content or formatting. Expose the action in the preview editor in English and Japanese, with selected-slide support and undo/redo.
- e18d19d: Preserve imported gradient stop color adjustments when changing gradient position, direction, brightness, or transparency in the editor. Gradient stops now expose ordered `colorTransforms` for round-trip editing through the API.
- e18d19d: Support null in getShapeRunFormatEffective to resolve paragraph end formatting. Preserve empty shape paragraph sizes when displaying and editing text so the following text does not shift when editing begins.
- e18d19d: Allow getShapeRunFormatEffective to resolve fields with a fieldIndex selector. Preserve inherited fonts, sizes, colors, and emphasis for shape fields when displaying, editing, inspecting, and copying their text.
- e18d19d: Read and preview theme-referenced shape fills, including solid colors, transparency, and gradient stops, using each shape's slide master theme. Direct shape fills continue to override theme references. The fill opacity reader accepts an optional presentation argument to resolve theme transparency.

  Use a shape style's text color and theme font defaults when its existing text formatting does not supply them, so text in PowerPoint's colored shape styles keeps its intended color.

- e18d19d: Add `isShapeLocked` and `setShapeLocked` for PowerPoint-compatible object geometry locks, preserving text editing and unrelated drawing constraints.

  Add individual and all-object locks in the editor Selection Pane. Locked objects cannot be moved, resized, rotated, aligned or grouped; their lock state survives saving and undo/redo.

- e18d19d: Add `getSlideAnimations` to read a slide's animation effects in click order, with each effect's target (whole shape or paragraph range), preset, start condition, duration, delay and paragraph-build settings. Each effect also says which `<p:seq>` it belongs to, so a player can tell the slide's click order from a sequence a viewer triggers by clicking a shape, and carries separate `playable` and `editable` flags. Effects this library cannot author or render — custom presets, composite effects driving several shapes, interactive sequences — are reported rather than dropped, and flagged so they are never silently rewritten or folded into the normal progression.

  `slideHasAnimations` no longer reports a slide whose only `<p:timing>` holds a video or audio clip's play controls as animated.

- e18d19d: Add slide copy, cut and paste to the preview navigator's context menu and keyboard shortcuts, with independent snapshots, undo/redo and saved output. Allow importSlide to omit the target layout to preserve the source layout, master and theme, registering imported masters in the destination presentation.
- e18d19d: Allow `getSlideMasterPartName` to resolve the owning master from either a slide or a slide layout handle.
- e18d19d: Slide numbers, as live fields rather than typed text.

  `setShapeTextField(shape, type, { text })` writes an `<a:fld>` — the slide's
  number, a date, a footer — replacing the shape's text body the way PowerPoint
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
- e18d19d: Edit solid fill and outline opacity through the existing shape setters while preserving imported theme colors and color transforms. The editor adds percentage transparency controls with mixed-selection support, undo and saved reloads.
- e18d19d: Extend the existing paragraph formatting API to table cells, render their bullets and paragraph spacing, and add bilingual editor controls for individual cell paragraphs and selected cell ranges. Preserve rich text and support undo, redo, and save/reload.
- e18d19d: Preserve empty table paragraphs' authored font size in the preview and text editor, so blank lines no longer collapse to a default size. The font controls now resolve inherited formatting at an empty cell paragraph's caret. Pass a null run index to getTableCellRunFormatEffective to read the effective paragraph end format.
- e18d19d: Table fields such as dates and slide numbers now inherit cell text formatting in previews, text editing, and copied text. The existing getTableCellRunFormatEffective API accepts a fieldIndex selector to resolve field formatting.
- e18d19d: Add effective table cell appearance resolution for embedded table styles and the built-in Medium Style 2 – Accent 1 style. Resolve theme fill and line references, preview solid cell fills and styled borders, preserve explicitly transparent cells, and stop adding white table backgrounds or gray borders that are absent from the presentation.
- e18d19d: Convert shape and table-cell text with sentence, lowercase, uppercase, title, or toggle case through the existing text setters while preserving formatting, hyperlinks, and original run XML.

  Add a Change Case menu to the editor's Home ribbon for selected text, the word at the caret, and selected text shapes, including outline editing and Undo/Redo.

  Keep Home ribbon controls accessible without horizontal scrolling at intermediate window widths.

- e18d19d: Text formats can now carry PowerPoint's theme tints and WordArt bevels.

  - `TextFormat.colorTransforms`, and `colorTransforms` on `outline`, `shadow`, `innerShadow` and `glow`, write `<a:lumMod>`, `<a:lumOff>`, `<a:tint>`, ... on the run's colors — the same field gradient stops already take — so "Accent 2, Lighter 60%" stays a theme color instead of a fixed RGB. `getShapeRunFormat` reads them back beside the unresolved color and `toWritableTextFormat` carries them. The shared `ShadowOptions`, `InnerShadowOptions` and `GlowOptions` gain the field, so `setShapeShadow` and `setShapeGlow` accept it too.
  - New `setShapeText3D(shape, value | null)` / `getShapeText3D(shape)` write and read the text body's 3-D (`<a:scene3d>` camera and light rig, `<a:sp3d>` top bevel, extrusion height, material and contour color), which is where PowerPoint puts its WordArt bevels. Presets are typed by the schema's enums (`CameraPreset`, `LightRigType`, `LightRigDirection`, `BevelPreset`, `PresetMaterial`).
  - `setShapeTextAutoFit` now places its element ahead of any `<a:scene3d>` / `<a:sp3d>`, as the schema requires.
  - `getShapeRunFormat` reports a run effect's scheme color as its token (it returned an empty string without a theme), and a text gradient stop without alpha no longer reads back as `opacity: 1`, which wrote an `<a:alpha val="100000"/>` on the way back.

- e18d19d: Support paragraph-internal line breaks when replacing shape or table-cell text with `newlines: 'break'`. The editor preserves these breaks for Shift+Enter and for Enter in title placeholders, matching Mac PowerPoint.
- e18d19d: Add `setShapeTextLanguage(shape, lang)` and `getShapeTextLanguage(shape)` for the proofing language of a shape's text (`lang` on every run and paragraph end). The editor's Review tab gains Check Accessibility, Language, comment Delete/Previous/Next, Show Comments and Hide Ink; commands the library cannot support yet are shown disabled with the reason.
- e18d19d: Support gradient and pattern fills on formatted text runs through `textFill`,
  including theme color transforms and round-tripping native WordArt fills.
  The existing `color` property remains the solid text color shorthand; passing
  both `color` and `textFill` is rejected.
- e18d19d: Add slide-background shape fills through `setShapeSlideBackgroundFill` and the editor's Fill pane, including multiple selection, undo and saved reloads. Fill readers expose the new `background` kind. Preview paints the slide background through these shapes while keeping it aligned through shape and group transforms.
- e18d19d: Allow `copyShape` to copy between presentations while preserving related images, charts, embedded workbooks and unknown dependencies. Allocate unique IDs for every shape inside copied groups and preserve their internal connector references.

  Fix cut and paste in the development preview. Copied content now retains its state even after editing or deleting the source or opening another document, and undo/redo restores the pasted selection.

- e18d19d: Support reading and editing source crops on image-filled shapes with the existing image crop functions. Preview cropped image tiles at their cropped size, including mirrored tiles.
- e18d19d: Read table-cell paragraph and outline-level character defaults with `getTableCellRunFormatEffective`. Preserve inherited fonts, sizes, emphasis, and colors when displaying, editing, inspecting, and copying table text.

  Resolve inherited shape and table text colors through the slide color map, including detached editing previews.

- e18d19d: Transitions can now carry a duration and a sound. `setSlideTransition` accepts `durationMs`, written as PowerPoint 2010 writes it (`p14:dur` inside `mc:AlternateContent`, with the nearest `speed` as the fallback), and `getSlideTransition` reads it back, including from decks saved by PowerPoint. New `setSlideTransitionSound` / `getSlideTransitionSound` embed a WAV sound or stop earlier sounds; changing the effect keeps the sound. The editor's Transitions tab gains Preview, Duration and Sound, and the preview plays fast/medium/slow transitions at PowerPoint's 0.5/0.75/1 s.
- e18d19d: Add typed media trim and fade playback controls with PowerPoint OOXML round-trip support.
- e18d19d: Add video formatting reset to the editor ribbon and the `resetShapeVideoFormatting` API. Reset removes color corrections, borders and effects and restores a rectangular shape while preserving the video, poster image, crop and dimensions.
- e18d19d: Preserve and edit wheel transition spoke counts, with Japanese and English controls. Play clockwise wheel transitions in presentation mode, including default four spokes and a fade for zero spokes.

### Patch Changes

- e18d19d: Read imported HSL and scRGB background colors instead of displaying an empty color.
- e18d19d: Read transparency, brightness, focus position, tile bounds and scaling from slide, layout and master background gradients. These settings are no longer omitted when inspecting an imported presentation.
- e18d19d: Allow setting text formatting on blank autoshapes before typing. Preserve paragraph-end formatting through save/reload and subsequent text insertion, report zero paragraphs when a text-capable shape has no body, and show the preview formatting bar for empty shapes. Restore the Japanese text-format command label.
- e18d19d: Replace text across adjacent formatting runs in slide text, including table cells. Replacement text inherits the first matched run’s format, while surrounding text and run metadata stay intact. Matches stop at paragraph and explicit line-break boundaries. Regular expressions now operate on each complete text segment instead of each individual run, so anchors and replacement context follow the visible segment text. The return value remains the number of changed text nodes.
- e18d19d: Display the borders and transparent cells of the built-in No Style, Table Grid style when a presentation stores only its style ID.
- e18d19d: Read percent-form bullet sizes at their intended scale instead of rendering the list marker a thousand times too small.
- e18d19d: Add bilingual speaker-note and slide-transition dialogs to the development preview, including automatic advance and applying a transition to all slides in one undo step. Correct speaker-note reads and updates to use the body placeholder without overwriting footers, read soft breaks and field results, and retain automatic-advance settings when no transition effect is selected.
- e18d19d: Fix changing font colors on text with gradient, pattern, or other fills so the previous fill is replaced instead of producing conflicting fills in saved PowerPoint files.
- e18d19d: Fix reading and rendering of DrawingML preset colors such as `red`, including
  colors used by shape fills and text after a presentation is saved and reopened.
- e18d19d: Scale table text with no authored font size when fitting slide content to a new page size. Preserve explicit font-size inheritance and retain scaled defaults for empty cells and saved presentations.
- e18d19d: Fix explicit text-range replacements that switch between paragraph separators and soft line breaks without changing the visible text.
- e18d19d: Respect explicit shape, layout, and master fill choices when resolving effective gradient fills, so `noFill` and solid fills no longer inherit an ancestor gradient.
- e18d19d: Preserve preset and system colors when reading gradient stops from imported presentations.
- e18d19d: Preserve both recolor colors when imported pictures use HSL or scRGB duotone colors.
- e18d19d: Image brightness and contrast now work on image-filled shapes. The editor preserves these corrections when switching fill types and restoring the image.

  Image corrections, opacity, and DrawingML color transforms now read both fixed-point and percent-suffixed values correctly. Hue offsets use angle units, preventing incorrect colors in imported presentations.

- e18d19d: Preserve the actual package path when switching slide layouts, including imported layouts outside the conventional slideLayouts folder. This prevents broken or incorrect layout references after switching and saving.
- e18d19d: Recognize `true` and `false` as well as numeric DrawingML flip flags in imported presentations. Flipped shapes and groups now retain their orientation in the preview and when ungrouped, including after saving and reloading.
- e18d19d: Resolve internal slide link targets by package and part name so preview links retain the correct destination after reading, reordering, and saving a presentation.
- e18d19d: Read imported percentage line spacing at the correct scale, including paragraph and inherited list styles. Percent-form spacing such as `150%` no longer collapses the lines together.
- e18d19d: Keep event-triggered audio and video from being reported as automatic playback. Remove media timing nodes inside nested playback sequences when their shapes are deleted while preserving unrelated timing content.
- e18d19d: Apply last-column formatting to horizontally merged table cells that reach the right edge, matching PowerPoint even when the cell starts in an earlier column.
- e18d19d: Apply total-row formatting to vertically merged table cells that reach the last row, matching PowerPoint even when the cell starts in an earlier row.
- e18d19d: Convert PowerPoint's native Play in Background timing tree when changing an audio clip between automatic and click playback, preserving the existing timing IDs and media settings.
- e18d19d: Show and edit playback settings for audio and video nested inside animation timing groups. Preserve enclosing start conditions when changing volume, looping, or video display settings, and account for parent delays when reading and editing simple automatic playback. Reject unsupported start-condition edits before changing the document. Also reject non-finite volume values instead of writing invalid XML.
- e18d19d: Fix existing text in shapes and table cells incorrectly adopting the font, size, or emphasis saved for newly inserted text at the end of a paragraph.
- e18d19d: Read imported path-gradient focus rectangles correctly when PowerPoint omits zero insets. Preserve small integer percentage values and percentage-string values when editing and saving gradient fills.
- e18d19d: Read image crop offsets correctly when imported presentations store them as percentages. Include selected text and theme fonts in the editor's font picker, including custom font names.
- e18d19d: Preserve charts, embedded workbooks, speaker notes, media, and relationship-linked extension parts when importing slides. Retain shared dependencies and cycles, allocate independent copies for repeated imports, and reject missing dependencies before modifying the destination package.
- e18d19d: Keep nested audio and video playback settings when clearing a slide's animations, preserving their enclosing timing conditions.
- e18d19d: Keep existing outline color, theme references, and opacity when changing only line width. Image border width and dash controls preserve the existing color instead of replacing it with an opaque resolved color.
- e18d19d: Preserve original run XML on the unchanged portions of a text range replacement in shapes and table cells. Partially edited fields still become literal text without retaining field-only paragraph properties.
- e18d19d: Honor imported XML boolean spellings for through-black transitions and click advancement. Play cut-through-black transitions in presentation mode, respecting transition speed, reduced motion, and cleanup when interrupted or exited.
- e18d19d: Read explicit XML `false` values for bold and italic correctly, preventing imported text from appearing bold or italic when those styles are disabled.
- e18d19d: Preserve fonts, text colors, bold and italic formatting inherited from embedded table styles when displaying, editing and copying table text. Apply header, footer, banded row and column, and corner formatting in PowerPoint's precedence order while keeping explicitly formatted cell text unchanged.
- e18d19d: Read scRGB and HSL colors in imported fills, strokes, and gradient stops.
- e18d19d: Resolve scRGB and HSL color elements used by shape shadows and glow effects when reading imported presentations.

  Preserve opacity modifiers and offsets on imported effects, including percent-form shadow opacity.

- e18d19d: Recognize automatic playback for audio saved with PowerPoint's Play in Background command. Preserve its command timing when editing other playback settings, and reject unsupported start or delay changes instead of leaving conflicting playback instructions.
- e18d19d: Slide background readers now resolve imported DrawingML preset and system colors instead of treating them as an empty solid fill.
- e18d19d: Preserve the original outline after rejecting invalid width or opacity values, including when a later edit saves the shape.
- e18d19d: Resolve DrawingML percentage RGB and HSL colors when reading imported presentations.
- e18d19d: Resolve DrawingML color transforms and theme color-map overrides when reading slide, layout, and master backgrounds.
- e18d19d: Apply gradient stop brightness and theme colors when previewing slide, layout and master backgrounds. Background gradient readers now include resolved stop colors while preserving their original color tokens for editing.
- e18d19d: Resolve text and shape theme colors through layout color-map overrides. Preserve an explicit slide choice to use the master's color map.
- e18d19d: Preserve placeholder orientation and size categories when creating slides from layouts, including slides inserted in the development editor. New placeholders still inherit layout formatting and start with empty editable text.
- e18d19d: Preserve group rotation and reflections when ungrouping objects. Children retain their transformed positions and compose the group angle and flips with their own, including nested groups. The development preview supports this through the existing English and Japanese ungroup controls and Undo/Redo.
- e18d19d: Open the editor's Selection Pane to select nested objects, rename them, and show or hide individual objects or the whole slide's objects with undo and autosave.

  Fix renaming and visibility changes on group shapes. Omit hidden objects and hidden group descendants from previews and canvas hit targets while preserving their editable content.

- e18d19d: Allow text alignment and vertical anchoring on blank autoshapes. Add bilingual quick alignment controls that reflect mixed selections and apply to all selected text shapes with a single undo step.
- e18d19d: Read preset and system colors in imported shape fills and outlines, including inherited placeholder paint, instead of returning an empty color.
- e18d19d: fix: setShapeText / appendShapeText now add text to a shape that has no text body

  Previously, setting text on a shape authored without one (e.g. `addSlideShape`
  called without `text`) threw `shape "…" has no <p:txBody>`. PowerPoint always
  gives an autoshape a text body so you can click in and type, so these functions
  now create the body on demand and populate it, matching that behavior. Picture /
  table shapes still throw, since they are not text-bearing.

- e18d19d: Resolve text fonts from each slide's own master theme, and honor shape-style font and color defaults ahead of inherited placeholder formatting. Shapes using their group's fill now inherit solid colors, transparency, and gradient details instead of falling back to their own theme style.
- e18d19d: Preserve and render negative image crop offsets, including when switching away from a picture fill and restoring it in the editor.
- e18d19d: Preserve slides with absolute or normalized relative relationship targets when moving or sorting a presentation. Speed up moving multiple editor slides by applying the final order once.
- e18d19d: Preserve explicit cell-side borders in banded table styles, including interior cells, while retaining interior borders when a side is unspecified.
- e18d19d: Read small and percent-suffixed text baseline offsets correctly so imported superscript and subscript positions retain their intended values.
- e18d19d: Display theme-referenced background gradients instead of a solid color, including radial backgrounds selected with PowerPoint's Background Styles gallery. Resolve gradient colors through the owning slide master's theme and color map while preserving the original theme and background XML on save.
- e18d19d: Resolve shape style effect references against the owning slide master's theme, including placeholder colors and explicit empty effect lists in the layout-to-master cascade.
- e18d19d: Preserve theme-based shape outlines in previews, including line colors and widths from PowerPoint Quick Styles. Direct line formatting now retains theme properties that it does not override.
- e18d19d: Replacing a copied picture now changes only that picture, preserving other pictures that share its media or slide relationship. Repeated same-format replacements reuse the detached media part. Invalid image crop input now preserves the previous crop, including when a subsequent edit is saved.
- e18d19d: Reject invalid slide dimensions before modifying the presentation. Page setup now enforces PPTX’s 1–56 inch range in both inches and centimeters, with matching Japanese and English validation messages.

## 0.21.0

### Minor Changes

- 28731c2: feat(charts): `ChartSpec` is now a union over `kind`, so a chart cannot be given a field it has no element for

  Every chart kind carries only the fields its OOXML element defines. A
  `scatterStyle` on a column chart, an axis title on a pie, `bar3DShape`
  without `view3D`, a scatter series with no `xValues` — each of these used to
  either throw at save time or serialize into nothing. They are now type errors
  at the call site.

  `getShapeChartSpec` and `SlideChartData.spec` return the new `ReadChartSpec`,
  which keeps every field optional: a deck authored elsewhere can combine
  fields no kind draws, and dropping them on read would lose data. Narrow one
  back to the write side with `isChartSpec` before passing it to
  `addSlideChart` / `setChartSpec`.

  Two behaviour fixes came out of the change:

  - A 3-D line chart no longer reports `lineMarkers: false` on read.
    `<c:line3DChart>` has no `<c:marker>` element, so the value was meaningless.
  - Pie and doughnut charts no longer get the deck's body-text colour written
    onto axis-label styles they have no axes to show.

## 0.20.0

### Minor Changes

- 04eda82: Colors are now a type, not a string. `Color` (`#RRGGBB`, the `#RGB` shorthand, a theme token such as `accent1`, or the `scheme:`-prefixed form) replaces `string` on every authoring call that takes one — shape and table fills, slide backgrounds, strokes, shadows, glows, gradient stops, connectors and run formats — so `'reddish'` or a mistyped `'accnet1'` is a compile error instead of a run-time throw. Chart colors take `HexColor`, which enforces what the docs already said: a chart series must resolve to a concrete sRGB value and cannot carry a theme token.

  The `#` is now required. Bare `RRGGBB` was one of four accepted spellings of the same thing, and admitting it collapsed the union back to `string`, taking every other color check down with it. The runtime parser still accepts it, so untyped JavaScript callers are unaffected.

  `@office-kit/pptx-dsl` picks this up too: `fill`, `stroke`, `background`, `stripeFill` and cell borders are typed the same way, so a bad color in a TSX deck fails `npm run check` rather than the export.

  Reading stays permissive, since a deck authored elsewhere can carry a scheme token outside its theme. `getShapeRunFormat` and friends now return `ReadTextFormat`, and the gradient readers return `ReadGradientFill` / `ReadGradientStop` — same shapes, with colors as `string`. To write a color you read back, pass it through the new `asColor`, which returns `null` when the value is not one this library can emit.

## 0.19.0

### Minor Changes

- 3eecc92: Rename `setShapeBullets` to `setShapeBulletStyle`. The old plural name read like it set the list's _content_, so calls meant to create a bulleted list reached for it and silently got back an unchanged, text-less shape. The new name says what the function actually does: it restyles the bullet glyph on paragraphs that already exist. Bulleted text is still authored in one call — `setShapeText(shape, 'A\nB\nC', { bullets: 'bullet' })`.
- cb702c1: Placeholder lookups and slide transitions now reject an unknown token at compile time instead of at runtime.

  `findSlidePlaceholder`, `findSlidePlaceholders` and `findLayoutsWithPlaceholderType` take the closed `ST_PlaceholderType` set (exported as `PlaceholderType`) rather than `string`, so a shape's display name or a typo no longer type-checks and then silently finds nothing. `setSlidePlaceholders` keys its `byType` record the same way, so a misspelled key is no longer skipped in silence.

  `TransitionEffect` gains the five spec effects it was missing — `comb`, `pull`, `random`, `strips` and `wheel` — and `TransitionOptions.effect` drops its `| string` escape, so a mistyped effect is a type error rather than a runtime throw. Reading stays permissive: `getSlideTransition` returns the new `SlideTransition` type, whose `effect` is still `string`, because a deck authored elsewhere can carry an effect this library does not model.

## 0.18.5

### Patch Changes

- 73b763d: `removeSlide` now drops slide relationships in other slides that pointed at the removed slide, along with the `<a:hlinkClick>` / `<a:hlinkHover>` elements that carried them. A slide-jump click action stores its relationship on the _referring_ slide, so removing the target used to leave a dangling relationship: PowerPoint rejects the package and a later `duplicateSlide` of the referring slide throws `Cannot duplicate missing dependency`.

## 0.18.4

### Patch Changes

- 7b15e41: fix: removing or replacing a link left its relationship behind on the slide

  `setShapeClickAction(shape, null)`, `setShapeHyperlink(shape, null)` and `setShapeRunHyperlink(…, null)` deleted the `<a:hlinkClick>` but kept the `hyperlink` / `slide` relationship it pointed at, and replacing a link added a second one beside the first. The orphan stayed in `_rels`, so the URL remained in the saved file and a slide jump kept its target slide alive as a dependency: after removing that target, `duplicateSlide` failed with `Cannot duplicate missing dependency /ppt/slides/slideN.xml`.

  Each of the three setters now releases a link relationship once nothing on the slide points at it. A relationship other runs still use is kept, and relationships of every other type are untouched.

## 0.18.3

### Patch Changes

- 9b8f8a0: Reduce preview latency on large TSX decks by avoiding repeated parsing of existing slides, themes and relationships. Keep existing slide handles live when appending or duplicating slides.

  Verify builds when an embedded agent finishes and return failures to Claude Code or Codex for up to three automatic repair attempts. Support Shift+Enter in the Claude terminal and provide native TSX agenda examples and foreground/background contrast guidance.

## 0.18.2

### Patch Changes

- 78ee67f: fix: theme readers and the preview picked the first theme part by name instead of the slide master's theme

  `getPresentationTheme`, `getPresentationFonts`, `setPresentationTheme`, `setPresentationFonts` and every color resolver built on them (slide backgrounds, `scheme:*` fills, the deck body text color, the preview renderer) read `/ppt/theme/theme1.xml` whenever a package carried more than one theme. Notes and handout masters carry their own theme, and Google Slides exports write the notes theme as `theme1.xml` and the slide theme as `theme2.xml`, so a dark deck rendered on white with black text.

  The deck theme is now the one the first slide master (in `<p:sldMasterIdLst>` order) relates to, then the presentation part's own theme relationship, and only then the first theme part by name.

## 0.18.1

### Patch Changes

- c8bf75d: fix: `getParagraphAlignment` and `getTableCellAlignment` returned an invalid `algn` value from a file as if it were a valid alignment

  Both getters cast the attribute straight to `ParagraphAlignment`, so a hand-edited or third-party deck with `algn="bogus"` handed `'bogus'` to typed code. They now parse the value, and anything outside `ST_TextAlignType` reads as `null` (unset), like the other alignment readers already did.

  Their return type is now the new `ParagraphAlignmentToken` (`'l' | 'ctr' | 'r' | 'just' | 'dist' | 'justLow' | 'thaiDist'`) instead of the wider `ParagraphAlignment`. Runtime values are unchanged — these two getters have always returned the spec token, so `setParagraphAlignment(shape, 0, 'center')` reads back as `'ctr'` — but the old type also listed `'center'`, which let `getParagraphAlignment(…) === 'center'` compile although it can never be true. That comparison is now a type error; compare with `'ctr'`, or read the plain-English name from `getParagraphPropertiesEffective` / `getTableCellParagraphs`.

## 0.18.0

### Minor Changes

- 7e89e62: Validate enum inputs at authoring boundaries and throw a descriptive `RangeError`
  instead of writing invalid PowerPoint XML or silently selecting another mode.

  Affected APIs:

  - Shapes: `addSlideShape` preset and text anchor, including calls through the TSX DSL.
  - Strokes and fills: `setShapeStrokeDash`, `setShapeStrokeCap`, `setShapeStrokeJoin`,
    `setShapeStrokeCompound`, `setShapeStrokeArrow` (end, type, width, length), and
    `setShapeGradientFill` path.
  - Text: `setShapeAlignment`, `setParagraphAlignment`, `setShapeTextAnchor`,
    `setShapeTextAutoFit`, `setShapeTextWrap`, `setShapeTextDirection`, `setShapeBullets`,
    `setParagraphBullet`, and `setShapeText` bullets; `setShapeTextFormat`,
    `setShapeRunFormat`, and `setShapeParagraphs` formatting and alignment.
  - Tables: `setTableCellAlignment`, `setTableCellTextFormat`, `setTableCellParagraphs`,
    `setTableCellTextDirection`, `setTableCellAnchor`, and `setTableCellBorders` dash.
  - Charts: `addSlideChart` and `setChartSpec` kind, grouping, axes, legend, layout,
    data labels, series markers, line dash, trendlines, error bars, and chart styles.
  - Slides: `setSlideSize` type, `setSlideTransition` speed and split orientation,
    and `setShapeAnimation` effect.

  Previously accepted inputs such as preset `'rounded-rect'`, alignment `'middle'`,
  stroke dash `'dotted'`, text auto-fit `'shrink'`, and auto-numbering `'bogus'` now
  throw `RangeError`. Underline and strike strings must be exact XSD tokens:
  use `'sng'` / `'sngStrike'` (or `true`), not `'single'`. Existing JavaScript
  `null` values for underline, strike, and cap still remove the corresponding
  attributes; their TypeScript types are unchanged.

  Unknown chart kinds now throw `RangeError: addSlideChart: kind: … is not one of: …`
  instead of `Error: unsupported chart kind: …`; `setChartSpec` uses its own API name
  in the new message. Code that checks the exception class or message should adapt.

  Valid ECMA-376 shape presets remain supported, including presets outside the
  TypeScript autocomplete list. Rejected bullets and transitions preserve existing
  content, and rejected table-cell alignment or text-format enums do not create a
  text body.

### Patch Changes

- 7e89e62: fix: Preserve valid table XML when setting alignment on a covered merged cell whose text was dropped. Recreated text bodies now include the required paragraph, so the alignment is applied and survives saving and reloading.
- 64fa3c7: Fix nested list indentation when changing a paragraph's level after applying bullets or numbering, including after copying or reloading a presentation. Indentation that differs from the previous level's default pair is preserved.

## 0.17.0

### Minor Changes

- 71deec5: feat: author every chart type — scatter, bubble, radar, stock, surface, the 3-D variants and pie-of-pie

  `addSlideChart` / `setChartSpec` now write all sixteen plot types of ECMA-376, and `getShapeChartSpec` reads every one of them back.

  - `kind: 'scatter' | 'bubble' | 'radar'` are authorable (they used to throw "read-only"). Scatter and bubble series carry their own `xValues` (and `bubbleSizes`); the embedded workbook lays them out per series so "Edit data" opens onto the right cells.
  - New kinds `'stock'` (3 series: high, low, close; or 4 with open first, drawn as candlesticks) and `'surface'` (`surfaceContour` for the top-down contour plot, `surfaceWireframe` for the mesh).
  - 3-D is a modifier: `view3D: { rotX, rotY, rightAngleAxes, perspective, depthPercent, heightPercent }` on `bar` / `column` / `line` / `area` / `pie` selects the 3-D element, with `bar3DShape` (cylinder, cone, pyramid, …), `gapDepthPct` and `seriesAxis`. `bubble3D` shades bubbles as spheres.
  - `ofPie` turns a pie into a pie-of-pie / bar-of-pie chart.
  - Per series: `errorBars` / `xErrorBars` (fixed, percentage, standard deviation, standard error, custom), `fillOpacity`.
  - Per chart: `dataTable`, `upDownBars`, `categoryAxisDate` (date axis with time units), `categoryGroupLevels` (multi-level category labels), `categoryAxisScaling` (the x axis of scatter / bubble charts, or a date axis' range), `plotAreaLayout` / `titleLayout` / `legend.layout` (manual placement), `valueAxisLineHidden` / `categoryAxisLineHidden`, `valueAxis.displayUnitsLabel`.
  - Data labels gain `showBubbleSize`, `showLegendKey`, `fillColor`, and per-point `text` (a literal label, e.g. naming one scatter point).
  - A spec whose fields contradict each other (`view3D` on a scatter chart, error bars on a pie, a date axis with non-numeric categories, …) throws with a message naming the field, instead of writing a chart PowerPoint would repair.

  Behavior changes when reading existing decks:

  - A stock chart now reads as `kind: 'stock'` (was `'line'`) and a surface chart as `kind: 'surface'` (was `'column'`). `getPresentationChartKindCounts` has the two new keys.
  - 3-D charts still read as their flat kind, now with `view3D` set, so `setChartSpec(chart, getShapeChartSpec(shape))` no longer flattens a 3-D chart.
  - fix: for scatter / bubble charts `valueAxis` (and the other `valueAxis*` fields) described the **x** axis. They now describe the y axis; the x axis reads into `categoryAxis*` and `categoryAxisScaling`.
  - fix: charts with more than 25 series wrote workbook references past column `Z` as invalid cell addresses.

  `@office-kit/pptx-preview` draws stock charts as lines and surface charts as columns (as before), and honors `categoryAxisScaling` on scatter / bubble charts.

- 71deec5: feat: embed video, audio and online video with `addSlideMedia`, and read them back with `getShapeMedia`

  - `addSlideMedia(slide, { kind: 'video' | 'audio', data, ... })` embeds a clip from bytes; `{ kind: 'online', url }` links an online video. YouTube watch / `youtu.be` / shorts URLs are rewritten to the embed URL; any other `http(s)` URL is stored as given, and anything else throws.
  - The container is detected from the bytes (mp4, m4v, mov, webm, avi, wmv, mp3, wav, m4a, ogg, wma); pass `format` to override. An undetectable clip, an unreadable `poster`, or a bad URL throws before anything is added to the package.
  - The new shape is a picture showing the poster frame: `setShapeImage` / `getShapeImageBytes` replace and read it. Without `poster`, a small built-in play-button image is used.
  - The slide gets the `<p:video>` / `<p:audio>` time node PowerPoint writes itself, which is what makes the play controls appear in the slide show. Identical clip bytes are stored once per deck.
  - `getShapeMedia(shape)` returns `{ kind: 'video' | 'audio', partName, contentType, bytes }` or `{ kind: 'online', url }`, also for media authored by PowerPoint, PptxGenJS or python-pptx. `findShapesWithMedia(slide)` lists a slide's clips.
  - `@office-kit/pptx-dsl` gains a `Media` element with the same `kind` / `data` / `url` / `poster` props. It now requires `@office-kit/pptx` >= 0.17.0.
  - `copyShape` gives the copied clip its own time node, `removeShape` / `clearSlideShapes` remove it, and `importSlide` now carries video / audio parts (copied once per clip) and online-video links across decks — previously an imported slide with media was left with dangling relationships.
  - `setShapeAnimation` now works on a slide that holds a clip but no animation yet (it used to throw), and `clearSlideAnimations` keeps clips' time nodes instead of removing their play controls.
  - fix: `duplicateSlide` copied a slide's video / audio bytes for every duplicate, because the library's video / audio / media relationship-type constants did not match the URIs PowerPoint writes. Clips are now shared between the original and the duplicate, as documented.
  - `validatePresentation` reports a video / audio relationship whose part is missing, and a media time node whose shape is no longer on the slide.

## 0.16.1

### Patch Changes

- b8adc13: Duplicated slides now have independent charts, embedded workbooks and notes.
  Editing a duplicate no longer changes the original's chart or notes. Unknown
  owned dependency parts are copied too, including cyclic relationships. Layouts,
  masters, themes and media remain shared. Invalid missing dependencies fail before
  parts are added.

## 0.16.0

### Minor Changes

- 22f1f30: feat: author and read the complex-script typeface (`<a:cs>`), and align the blank deck's font scheme with Office's

  - `TextFormat.fontComplexScript` sets and reads `<a:cs typeface="…"/>`, the third typeface slot next to `font` (`<a:latin>`) and `fontEastAsian` (`<a:ea>`). It works everywhere the other two already did: `setShapeRunFormat` / `setShapeTextFormat`, `setShapeParagraphs` and `setTableCellParagraphs` (runs and `endFormat`), and it comes back from `getShapeRunFormat`, `getShapeRunFormatEffective` (including `+mj-cs` / `+mn-cs` and the theme's complex-script fallback), `getShapeParagraphElements`, `getParagraphEndFormat` and `getTableCellParagraphs`.
  - `ChartTextStyle.fontComplexScript` writes `<a:cs>` on a chart title / axis / data-label / legend style, and the chart reader returns it. `font` keeps filling the Latin and East Asian slots only, so a chart can name a complex-script face on the legend alone.
  - The three slots stay independent: setting one never changes the other two. Rewriting a slot writes its `typeface` attribute alone, so a `pitchFamily` / `charset` the source file carried on that element is dropped — as was already the case for `<a:latin>` and `<a:ea>`.
  - `createPresentation`'s blank deck now carries the standard per-script font list (`<a:font script="Thai" typeface="Cordia New"/>` and its 46 siblings) in both `majorFont` and `minorFont`, the same list Office's own default theme ships. It is what resolves a face for text in a script no run names one for. Every deck `createPresentation()` produces therefore has a larger `theme1.xml` than before — its bytes change — and text in a script with no explicit typeface now resolves through that list.
  - Chart titles and axis titles no longer carry a `lang="en-US"` language tag on their run. The chart-level tag is still `ChartSpec.language` (`<c:lang>`).

## 0.15.1

### Patch Changes

- d6036a5: Reject U+FFFE, U+FFFF, and unpaired surrogates in authored XML text and attributes, preventing malformed PPTX parts and silent replacement with U+FFFD. Valid supplementary characters, including emoji, remain unchanged.

## 0.15.0

### Minor Changes

- 96a37b5: feat: author and read the paragraph-end format, and merge table cells without keeping the covered cells' text

  - `setShapeParagraphs` / `setTableCellParagraphs` accept `endFormat` on each paragraph, written as `<a:endParaRPr>` after the runs. It is the place to give a paragraph with no runs a font size. Omitting `endFormat` writes nothing, as before.
  - `getParagraphEndFormat(shape, paragraphIndex)` reads a paragraph's end-mark format (`null` when absent), and every paragraph returned by `getTableCellParagraphs` now carries `endFormat`.
  - `mergeTableCells(table, block, { coveredText: 'drop' })` removes `<a:txBody>` from the cells the merge covers, the shape PptxGenJS writes. The default (`'keep'`) still leaves the covered cells' text in the XML. Covered cells without `<a:txBody>` read back as no paragraphs and can be written to again.
  - `setTableCellText` / `setTableCellParagraphs` on a cell that has no `<a:txBody>` now always insert it as the cell's first child, so a cell that carries only `<a:extLst>` stays schema-valid.
  - `setShapeParagraphs` / `setTableCellParagraphs` now validate every paragraph before they replace anything: a rejected format (for example a font size out of range) throws and leaves the existing text as it was.

## 0.14.0

### Minor Changes

- f7a710f: feat: format the secondary value axis, author per-point data labels, and write mixed-format paragraphs

  - `ChartSpec.secondaryValueAxis` sets scaling, number format, title, label style, gridlines, line color, tick marks, `tickLabelPos` and `crossBetween` on the right-hand axis that `secondaryAxis` series use; `getShapeChartSpec` reads it back. Its colors are validated like the primary axis colors, and its unset label / title colors take the deck's body-text color like the primary axes do, on `setChartSpec` as well as `addSlideChart` (the two now bake unset text colors identically). A chart whose every series is `secondaryAxis` is rejected: at least one series must plot on the primary axis.
  - `ChartSeries.pointDataLabels` authors per-point `<c:dLbl>` overrides (content, format, font) the way pie / doughnut exporters do, and the reader surfaces them. A `<c:dLbls>` holding only per-point overrides no longer reads back as an all-false series-level `dataLabels`. On read, toggles a point override leaves out inherit the series-level values, a deleted point label reads as an override with every toggle off, and overrides past the series' point count are dropped.
  - `setShapeParagraphs(shape, paragraphs)` and `setTableCellParagraphs(cell, paragraphs)` replace a shape's or cell's text with paragraphs that each carry their own alignment and several differently formatted runs; run text is written verbatim. An empty paragraph list is rejected (use `[{ runs: [] }]` for an empty body).
  - `ChartSpec.valueAxisTickLabelPos` positions the primary value-axis tick labels (`none` / `low` / `high` / `nextTo`), mirroring `categoryAxisTickLabelPos`.
  - `ChartSeries.lineColor` sets the series outline / line color separately from its fill (a doughnut's slice borders, for example), and the reader returns the `<a:ln>` color of every series.
  - `ChartSeries.markerColor` / `markerLineColor` set the marker fill and outline colors of a line / scatter / radar series, and the reader returns both. The builder now always writes them (the fill defaults to the series color, the outline to the fill), because PowerPoint paints a marker without `<c:spPr>` in the theme's automatic color rather than the series color.
  - `@office-kit/pptx-preview` paints chart markers (data points and legend swatches) in `markerColor` / `markerLineColor` and strokes line / scatter / radar series in `lineColor`, instead of always using the series color.
  - Axis lines and gridlines carry a width next to their color: `valueAxisLineWidthEmu` / `categoryAxisLineWidthEmu`, `valueAxisMajorGridlineWidthEmu` / `valueAxisMinorGridlineWidthEmu`, `categoryAxisMajorGridlineWidthEmu` / `categoryAxisMinorGridlineWidthEmu`, and `secondaryValueAxis.lineWidthEmu` / `majorGridlineWidthEmu` (EMU, 12700 = 1 pt). The reader returns them, and a width of 0 — valid in the schema — now reads back on these and on a series' `lineWidthEmu` instead of turning into `undefined`. Without a width PowerPoint draws these lines at 0.75 pt, so a 1 pt line from another writer used to come back thinner.
  - `ChartDataLabels.showLeaderLines` writes and reads `<c:showLeaderLines>` on series- and chart-level labels (per-point overrides have no such setting).
  - Chart, axis and secondary-axis titles no longer force a 14 pt default size or a horizontal layout: without `sizePt` the application default applies, and a title without a rotation writes neither `rot` nor `vert="horz"` (PowerPoint reads a lone `vert="horz"` as "not rotated"), so a value-axis title takes PowerPoint's default (vertical) — pass `0` to keep it horizontal. The reader merges a title's paragraph defaults under its run style.
  - Combo series are shaped by their own `chartKind`: a line series in a column chart keeps its markers and `smooth`, and line / area plot groups write a schema-valid `<c:grouping>` (`standard` instead of the bar-only `clustered`).
  - Horizontal bar charts (`kind: 'bar'`) now put the category axis on the left and the value axis at the bottom, matching PowerPoint; the hidden companion category axis of a secondary axis follows the primary one.
  - The chart reader honors `<c:ptCount>` (including on multi-level category caches), so series with empty trailing points keep their full length on read-back; cache points past the authored count are dropped.
  - The chart reader identifies the secondary value axis by which plot groups reference it, not by its position, so a scatter chart's X axis at the top no longer reads as secondary and decks that list the secondary axis pair first read back correctly.
  - The XML parser applies XML 1.0 end-of-line handling: a raw CR LF or CR in a part reads as LF, while a `&#13;` character reference still yields a CR.

## 0.13.0

### Minor Changes

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

## 0.12.0

### Minor Changes

- 6294c52: Combo charts: per-series `chartKind` overrides and a secondary value axis.

  `ChartSeries` gains `chartKind` (`'bar' | 'column' | 'line' | 'area'`) to
  overlay e.g. a line series on a column chart, and `secondaryAxis: true` to
  plot a series against a right-hand secondary value axis — the standard
  PowerPoint combo layout for series with mixed units (counts vs. rates).
  The builder splits series into plot groups (`<c:barChart>` + `<c:lineChart>`
  …) and emits the secondary `<c:valAx>`/`<c:catAx>` pair on demand;
  `getShapeChartSpec` round-trips both fields. The preview renderer paints
  bars below line/area overlays, scales each axis from its own series, and
  draws the secondary ticks on the plot's right edge.

## 0.11.1

### Patch Changes

- e09b698: Relax the `engines.node` floor from `>=24.16.0` to `>=22.18.0` on both `pptx-kit` and `pptx-kit-preview` so the maintained LTS lines — Node 22 and Node 24 — are supported, and restore Node 22 to the CI test matrix. The published runtime bundles are unchanged; the previous floor reflected the dev toolchain's pin and needlessly blocked `pnpm install` (under `engine-strict`) on still-supported LTS releases such as Node 22.x and earlier Node 24 LTS patches (e.g. 24.13.x).

## 0.11.0

### Minor Changes

- 5d881b6: Add `setShapeAdjustValues(shape, values)` to author a preset shape's adjust-handle guides (`<a:prstGeom><a:avLst>`). It's the mutating companion to the existing `getShapeAdjustValues` reader — pass raw ECMA-376 guide values keyed by guide name. The common use is the `roundRect` corner radius via the `adj` guide (`0..50000`; `setShapeAdjustValues(shape, { adj: 5000 })` gives a subtle 5% rounding). Throws when the shape has no preset geometry.

## 0.10.0

### Minor Changes

- 6060afc: `setShapeRunFormat` / `getShapeRunFormat` / `getShapeRunFormatEffective` now
  support `fontEastAsian`, a per-run East Asian typeface override (`<a:ea>`),
  alongside the existing `font` (`<a:latin>`). Previously a run's CJK glyphs
  always fell back to whichever East Asian font the theme's major/minor font
  scheme happened to carry, with no way to set a distinct typeface (e.g. a
  serif headline vs. a sans-serif body) on an individual run. `fontEastAsian`
  resolves theme `+mj-ea`/`+mn-ea` tokens and falls back to the theme's major/
  minor East Asian font the same way `font` already does for Latin text.
- fa5e53f: Added two authoring capabilities aimed at complex, dense slide layouts
  (process diagrams, KPI cards, branded decks):

  - `groupShapes` / `ungroupShapes` compose a selection of top-level shapes
    into a single `<p:grpSp>` (and reverse it). The group's bounds are the
    union of its members; moving or resizing the group afterwards rescales
    every member on ungroup, so a "KPI card" or diagram node built from a
    rectangle + label can be treated — and repositioned — as one unit.
  - `setPresentationTheme` / `setPresentationFonts` patch a deck's color
    scheme and font scheme in place, so a from-scratch deck can be branded
    with a custom palette and typography without hand-authoring a template.
    Only the slots passed in are overwritten; every other slot keeps its
    existing value.

## 0.9.0

### Minor Changes

- eeb8659: Validate authoring inputs at the API boundary so out-of-range values throw a
  clear `RangeError` instead of silently emitting a schema-invalid `.pptx` that
  PowerPoint marks corrupt and "repairs".

  A generative schema-validation sweep surfaced a whole class of defects where a
  caller-supplied number/string was serialized straight into a constrained
  ECMA-376 attribute. These now reject (or, for GUIDs, normalize) at the boundary:

  - Run formatting: `setShapeRunFormat` font `size` (ST_TextFontSize, 1..4000 pt)
    and `spc` (ST_TextPoint). It also now accepts the 3-digit hex shorthand for
    run `color` / `highlight`, matching `setShapeFill` / `setShapeStroke`.
  - Tables: `setTableStyleId` / `addSlideTable` `styleId` (ST_Guid — a lowercase
    GUID from `crypto.randomUUID()` is now accepted and normalized to uppercase;
    a non-GUID string throws); `setTableCellBorders` `widthEmu` (ST_LineWidth);
    `setTableColumnWidth` / `setTableRowHeight` / `addSlideTable` `w`/`h`
    (ST_PositiveCoordinate); `setTableCellMargins` (ST_Coordinate32).
  - Charts: bar/column `overlapPct` (ST_Overlap), `gapWidthPct` (ST_GapAmount),
    and series `lineWidthEmu` (ST_LineWidth).
  - Animations / transitions: `setShapeAnimation` `durationMs` and
    `setSlideTransition` `advanceAfterMs` (xsd:unsignedInt); the transition
    `effect` token is validated against the spec's effect set (an empty or unknown
    string previously produced non-well-formed or schema-invalid XML).
  - Connectors / strokes: `addSlideLine` and `setShapeStroke` `widthEmu`
    (ST_LineWidth) and `addSlideLine` endpoint coordinates.
  - Text boxes: `setShapeTextColumns` `count` (ST_TextColumnCount, 2..16) and
    `gapEmu` (ST_PositiveCoordinate32); `setShapeTextMargins` insets
    (ST_Coordinate32); `setShapeTextBodyRotationDeg` (guards the ST_Angle overflow).
  - Fills: `setShapePatternFill` `preset` is validated against ST_PresetPatternVal
    (and the `PatternPreset` type is now the exact token union, not `string`).
  - Shape / image / chart geometry: `addSlideShape` / `addSlideTextBox` /
    `addSlideImage` / `addSlideChart` and `setShapePosition` / `setShapeSize`
    `x`/`y` (ST_Coordinate) and `w`/`h` (ST_PositiveCoordinate).

- eeb8659: Fix a batch of "generates but is schema-invalid / wrong" authoring bugs, add an
  AI-agent authoring skill, and smooth several LLM-facing rough edges.

  Correctness fixes (output is now schema-valid in these cases):

  - Notes slides emitted a `<p:notesSlide>` root instead of the spec's `<p:notes>`,
    failing schema validation.
  - Combining a run's text `color` with `highlight` emitted them out of order.
  - Combining stroke dash / arrowheads / join, or a paragraph's bullet with
    `setParagraphSpacing`, or a table cell's fill with `setTableCellBorders`,
    emitted child elements out of their schema-mandated order.
  - Table cells containing leading/trailing spaces, tabs, or newlines emitted an
    illegal `xml:space` attribute (and a newline now correctly splits a cell into
    multiple lines).
  - `setSlideTransition({ effect: 'none' })` emitted an invalid `<p:none/>`; per-effect
    attributes (`direction`/`orientation`/`thruBlack`) are now only emitted on
    effects that allow them, and `direction` is validated against the effect's own
    value domain (e.g. `blinds` takes `horz`/`vert`, `push` takes `l`/`r`/`u`/`d`) —
    a mismatched pair like `{ effect: 'blinds', direction: 'l' }` now throws instead
    of emitting schema-invalid XML.
  - Charts emitted `<c:marker>`, `<c:smooth>`, `<c:invertIfNegative>`, and
    `<c:trendline>` on series kinds that don't permit them (e.g. a trendline on a
    `pie`/`doughnut`/`radar` series, which `CT_PieSer`/`CT_RadarSer` reject), and
    `valueAxis` `min`/`max` in the wrong order.
  - `setShapeImageBrightness` / `setShapeImageContrast` emitted `<a:lumOff>` /
    `<a:lumMod>`, which aren't valid `<a:blip>` children; both now write a single
    schema-valid `<a:lum bright/contrast>`.
  - `setShapeGradientFill` ignored its documented `path` / `focus` options, silently
    downgrading radial/shape gradients to linear.
  - `importSlide` could emit a duplicate `rId` when the source slide's layout
    relationship wasn't `rId1`.
  - `setShapeAnimation` wiped any pre-existing `<p:timing>` on the slide (losing a
    template's authored animations); it now merges, so multiple shapes can animate.
  - `compactPackage` / `readPackagePart` / `setMediaPartBytes` matched part names
    case-sensitively, unlike the rest of the package layer — a referenced image
    whose rel-target case differed could be wrongly deleted or missed.

  Behavior change:

  - `setShapeImageContrast` now takes a `[-1, 1]` offset (`0` = no change) instead
    of the previous `[0, 2]` multiplier, matching the underlying `<a:lum contrast>`.
  - Unstyled connectors (`addSlideLine` without an explicit stroke) previously
    emitted no line style and rendered invisibly; they now carry a default
    `<p:style>` (`lnRef`/`fillRef`/`effectRef`/`fontRef`) so the line is visible.

  Ergonomics:

  - Colors accept the CSS-style 3-digit hex shorthand (`#f0a` → `FF00AA`).
  - New `setParagraphLineSpacing(shape, p, { kind, value })` (the writer counterpart
    to the existing getter).
  - `setTableCellBorders` accepts a partial border per side, so `{ color, widthEmu }`
    type-checks without spelling out `dash` (the read type `getTableCellBorders`
    returns stays strict — all fields populated).
  - `addSlideShape` `textAnchor` is narrowed to the valid vertical anchors
    (`'t' | 'ctr' | 'b'`).

  Docs:

  - New `skill/SKILL.md` — a guide for driving pptx-kit from an AI agent (canonical
    calls, design rules, footguns, and a QA protocol), with a verified worked
    example.

- eeb8659: Close a final batch of correctness defects a generative schema sweep surfaced,
  where the writer emitted a `.pptx` PowerPoint marks corrupt:

  - **XML-illegal control characters** in any text field (shape text, table cells,
    notes, chart titles/categories/series, hyperlink tooltip/URL, section names,
    comments) used to serialize raw, producing a non-well-formed part that
    corrupts the whole package. They are now rejected at serialization with a
    clear error; the XML-legal whitespace controls (tab / LF / CR) still pass
    through. (XML 1.0 forbids the other C0 controls outright — they cannot even be
    escaped as numeric references.)
  - **Chart percentages** are now range-checked at the boundary: `gapWidthPct`
    (ST_GapAmount, 0..500 — the previous limit of 65535 let 501..65535 through),
    doughnut `holeSizePct` (ST_HoleSize, 1..90), and pie/doughnut
    `firstSliceAngleDeg` (ST_FirstSliceAng, 0..360).
  - **Shape effects**: `setShapeShadow` `blurEmu`/`offsetEmu` and `setShapeGlow`
    `radiusEmu` are validated as ST_PositiveCoordinate (fractional rounds;
    negative / non-finite / over-max throws) instead of emitting an invalid value.
  - **Scheme-color round-trip**: the read-back getters return `scheme:<token>`, but
    the setters rejected that string. `setShapeFill` / `setShapeStroke` /
    `setSlideBackground` now accept the `scheme:` prefix, so `setX(getX(...))`
    round-trips. (An unknown `scheme:` token still throws.)
  - **`importSlide` / `mergePresentations`**: importing a slide that contains a
    chart left a dangling `r:id` (the chart frame referenced a relationship the
    imported slide no longer carried), producing a corrupt package. The orphaned
    graphic frame is now dropped, matching the documented "charts are not imported
    in v1" behavior.
  - **`addSlideShape` presets**: the math-operator tokens were misspelled
    (`minus`/`mult`/`div`/`equal`/`notEqual`) and not in `ST_ShapeType`, so the
    shape silently vanished on open. They are now the spec names `mathMinus`,
    `mathMultiply`, `mathDivide`, `mathEqual`, `mathNotEqual` (plus `mathPlus`).

### Patch Changes

- e9eae5c: Fix `<a:tint>` / `<a:shade>` colour resolution to compute in linear-light RGB,
  matching PowerPoint and LibreOffice. A 75% tint of black now resolves to a mid
  grey (~#8B8B8B) instead of the too-dark #404040, so colours derived from theme
  scheme transforms (subtitle placeholders, table banding, chart fills) render at
  the right lightness.

## 0.8.0

### Minor Changes

- 7200690: Drop Node.js 22 support. The minimum supported version is now Node 24.16. The published runtime bundles are unchanged; this only raises the `engines` floor and the CI/test matrix to Node 24.

### Patch Changes

- ffeaef0: Numbered lists now carry a bullet font and explicit start number

  `setShapeBullets('number')` / `setParagraphBullet(..., 'number')` emitted only
  `<a:buAutoNum>`, with no `<a:buFont>` — so the auto-number glyph fell through to
  whatever font happened to apply, instead of the theme's major font that
  PowerPoint and PptxGenJS use. Numbered lists now emit
  `<a:buFont typeface="+mj-lt"/>` ahead of the number and write the default
  `startAt="1"` explicitly, matching PowerPoint-authored output.

## 0.7.0

### Minor Changes

- 3ba1e3d: Drop Node.js 20 support. The minimum supported version is now Node 22.18. The build toolchain moved to tsdown, whose current release requires Node 22.18+; the published runtime bundles are unchanged.

## 0.6.3

### Patch Changes

- 099d77b: Emit PowerPoint's default cell insets on table cells

  `addSlideTable` cells now carry the explicit default insets PowerPoint and
  PptxGenJS both write — `<a:tcPr marL="91440" marR="91440" marT="45720"
marB="45720">` — plus a `<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>` that
  suppresses any inherited list bullet on the cell paragraph. The table renders
  identically (these match the values PowerPoint applies when they're absent),
  but the cell is now self-describing, so the output matches a PowerPoint- or
  PptxGenJS-authored table byte-for-byte at the cell level.

  Note: `getTableCellMargins` now returns the explicit `91440 / 45720` defaults
  for a freshly-authored cell instead of `null`.

## 0.6.2

### Patch Changes

- c41d8f9: Fix jammed bullet lists and unreadable chart categories

  - **Bulleted text boxes now indent correctly.** `setShapeBullets` /
    `setParagraphBullet` added the bullet glyph but no hanging indent, so a bullet
    authored on a text box (which inherits the master's `otherStyle`, marL=0, not
    the body style) rendered with the glyph jammed against the text. They now
    write PowerPoint's per-level default `marL` / `indent` (unless the caller set
    their own), matching PowerPoint and PptxGenJS.
  - **Charts with multi-level category references now read back.** The chart
    reader handled `<c:strRef>` / `<c:strLit>` categories but not
    `<c:multiLvlStrRef>`, which is what PowerPoint and PptxGenJS emit — so
    `getShapeChartCategories` (and the full `getShapeChartSpec`) returned an empty
    category list for those charts. It now reads the level's points.

## 0.6.1

### Patch Changes

- 333b19f: fix: line-chart series colors now paint the line. The color was written only
  as a bare `<a:solidFill>`, which doesn't color a line series' stroke, so
  PowerPoint ignored it and fell back to its automatic palette (a 4-series line
  chart authored as accent1–4 rendered blue/red/green/purple instead of the
  requested colors). The color is now also emitted on `<a:ln>`.
- 665d4c2: Fix unstyled, broken-looking tables from `addSlideTable`

  `addSlideTable` set the `firstRow` / `bandRow` flags but never wrote a
  `<a:tableStyleId>`, and `createPresentation` shipped no `tableStyles.xml` part.
  With no style to resolve against, PowerPoint painted the table as a borderless,
  unstyled block — a "broken" grid with no rules.

  - **Tables now reference PowerPoint's "No Style, Table Grid" built-in**
    (`{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}`) via `<a:tableStyleId>`, the same
    default PptxGenJS and PowerPoint itself emit, so a table resolves to a clean
    ruled grid. Callers can override with the internal `styleId` option (or the
    existing `setTableStyleId`).
  - **`createPresentation` now ships `/ppt/tableStyles.xml`** (referenced from
    `presentation.xml.rels`), matching every PowerPoint-authored deck, so the
    `tableStyleId` always has a backing part.

## 0.6.0

### Minor Changes

- 0f7c538: Preview: take the default text color from the deck's body style, not the `tx1` token

  The preview used `scheme:tx1` as the fallback color for runs without an authored color. On a template with an inverted color map (`tx1 → lt1`) that resolves to the light slot, so body text was painted white on the white background — the whole slide looked blank. PowerPoint instead takes the fallback from the master `bodyStyle` (e.g. `schemeClr bg1`). The preview now does the same via the newly exported `resolveDeckBodyTextColor(slide)`, so default-colored text and table-cell text resolve to the color PowerPoint actually paints.

  - New export **`resolveDeckBodyTextColor(slide)`** — the deck's resolved body-text color (master `bodyStyle`, run through the effective color map + theme). This is the color `addSlideTable` / `addSlideChart` bake in, now reusable by renderers.

### Patch Changes

- 0f7c538: Fix corrupt files from fractional EMU and sideways-spreading stacked bar charts

  - **Whole-EMU coordinates.** `inches` / `cm` / `mm` / `pt` / `emu` now round to integer EMU, and every shape / table / text-box / connector / chart offset is rounded on serialization. Floating-point drift from unit math (e.g. `3090672.0000000005`) previously reached `<a:off>` / `<a:ext>`, which is invalid `ST_Coordinate` (xsd:long) — PowerPoint flagged the file as corrupt and "repaired" it by zeroing the offending offsets, collapsing shapes to the slide origin.
  - **Stacked bar/column charts** now emit `<c:overlap val="100"/>` by default (and for `percentStacked`). Without it PowerPoint draws each series in its own sub-slot so the "stack" spreads sideways across the category. An explicit `overlapPct` still wins; clustered charts are unchanged.

## 0.5.0

### Minor Changes

- 4a2ede1: Resolve scheme colors through the slide's color map so inverted-map templates render correctly

  Templates whose slide master inverts the color map (`<p:clrMap bg1="dk1" tx1="lt1">`, common in Google Slides / Canva exports) previously rendered with swapped light/dark colors: slide backgrounds came out black in the preview while PowerPoint paints them white, and generated tables and charts came out with invisible text (the default `tx1` token resolved to the same color as the background).

  - **`getEffectiveColorMap(slide)`** — new export returning the slide's effective color map (the master's `<p:clrMap>` overlaid by a per-slide `<p:clrMapOvr>`). Color resolution and renderers apply it to `schemeClr` tokens before indexing the theme.
  - **`resolveDrawingColor(colorEl, theme, clrMap?)`** — accepts an optional color map; scheme tokens are remapped through it before the theme lookup. Omitting it preserves the previous behavior (correct for the standard map).
  - **`addSlideTable` / `addSlideChart`** now bake the deck's resolved body-text color onto table cells and chart text (axis labels, legend, data labels) so generated tables and charts stay readable regardless of the template's color map. Authored colors still win; override table cells afterwards with `setTableCellTextFormat`.
  - **`pptx-kit-preview`** resolves `schemeClr` tokens through the effective color map, so previews of inverted-map decks match what PowerPoint paints.

## 0.4.0

### Minor Changes

- 9809441: Chart label fonts, table cell merging, and aspect-ratio-preserving image placement

  **`ChartTextStyle.font`** — chart titles, axis titles, axis tick labels,
  legends, and data labels now accept a font face
  (`titleStyle: { font: 'Yu Gothic' }`). The builder writes both
  `<a:latin typeface>` and `<a:ea typeface>` so CJK families render
  correctly in PowerPoint, and `getSlideCharts` / `getShapeChartSpec` read
  the face back for round-trips. Works through both `addSlideChart` and
  `setChartSpec`. (The SVG preview keeps its fixed substitution font set —
  authored chart faces affect the emitted PPTX, not the preview raster.)

  **`mergeTableCells(table, { row, col, rowSpan, colSpan })`** — the write
  counterpart to `getTableCellSpan`. Merges a rectangular block into its
  top-left anchor, emitting `gridSpan` / `rowSpan` on the anchor and
  `hMerge` / `vMerge` on the covered cells per ECMA-376 §21.1.3.18.
  Out-of-range blocks, 1×1 blocks, and overlaps with existing merges are
  rejected with descriptive errors before anything is mutated.

  **`fit: 'contain'` on `addSlideImage` / `setShapeImage`** — preserve the
  image's aspect ratio instead of stretching to the target box. `'contain'`
  inscribes and centers the image (natural size read from the PNG / JPEG
  header); other formats fall back to the default `'fill'` (the existing
  stretch behavior) rather than erroring. On `setShapeImage`,
  `fit: 'contain'` re-fits the replacement image inside the picture's
  current box.

## 0.3.0

### Minor Changes

- 3459aa5: `createPresentation()` now returns an immediately-authorable deck

  Previously `createPresentation()` returned an OPC package with only the OPC
  defaults — no slide master, layouts, theme, or slide size — so
  `getSlideLayouts()` came back empty and `addSlide({ layout })` was
  impossible. From-scratch authoring (a headline feature in the README) did
  not actually work without loading a template file.

  `createPresentation()` now ships a slide master, the Office theme, and three
  layouts — `Blank`, `Title Slide`, and `Title and Content` — so you can go
  straight to `addSlide` / `addTitleSlide` / `addContentSlide` and `savePresentation`.
  Every emitted part is validated against the ECMA-376 XSDs in CI. The slide
  size defaults to 16:9 and is selectable: `createPresentation({ size: '4:3' })`.

  Also in this release (input-validation hardening at the authoring boundary):

  - `addSlideChart` now rejects a series `color` (and `pointColors` /
    `trendline.color` / plot- and chart-area fills / axis & gridline colors)
    that isn't an sRGB hex (`#RRGGBB` or `RRGGBB`) with a clear error, instead
    of silently emitting an invalid `<a:srgbClr val="…"/>` that PowerPoint
    dropped or repaired. Bare `RRGGBB` (no `#`) is accepted and normalized;
    scheme tokens like `accent1` are correctly rejected, since charts emit
    `srgbClr`.
  - `addSlideTable` with empty `rows: []` (or a row with no cells) now throws
    an actionable `addSlideTable: …` error at the boundary rather than
    producing a grid-less `<a:tbl>` that triggers PowerPoint's repair dialog.
    (The error message previously named the old internal `addTable` path.)
  - `findSlideLayout`'s case-sensitive, locale-dependent name matching is now
    documented in its JSDoc and the README, pointing readers to the
    locale-stable `findSlideLayoutByType` and to `RegExp`/`i` for
    case-insensitive name lookups. No behavior change.

  No breaking changes. `createPresentation()` keeps its zero-argument call
  signature; the new `{ size }` options object is optional.

## 0.2.0

### Minor Changes

- b03e0cb: feat: fidelity calibration sweep — measured against LibreOffice ground truth,
  mean fg-SSIM rose from ≈0.66 to ≈0.78 (≈0.81 excluding documented
  divergences). Body placeholders now inherit the master `bodyStyle` bullet and
  hanging indent through the paragraph cascade (new `bullet` field on
  `ParagraphProperties` from `getParagraphPropertiesEffective`); charts no
  longer invent a legend when the XML authors no `<c:legend>`, and the value
  axis gets Excel-style headroom above the data max with the tick step
  preserved; the chart builder writes `<c:smooth val="0"/>` explicitly on line
  series (the schema default for an absent element is smooth=1, which made
  LibreOffice draw unauthored lines as curves); and the pure-SVG text layer is
  nudged 0.75px left to land on LibreOffice's pixel grid.

## 0.1.0

### Minor Changes

- 263bf52: feat: `getShapeAdjustValues(shape)` returns the `<a:prstGeom><a:avLst>
<a:gd name=… fmla="val N"/></a:avLst>` map (preset adjust-handle
  values). Only literal `val` formulas are surfaced; computed formulas
  (`pin`, `+-`, etc.) reference the preset's built-in guides and
  aren't useful without them.

  Playground reads `adj` on `roundRect` to project the authored corner
  radius — previously every rounded rectangle painted at a hard-coded
  18%. Other presets (callouts, arrows, etc.) can adopt the same getter
  as their renderers grow.

- a944352: feat(site/playground): render auto-numbered bullets. Paragraphs with
  `bulletStyle === 'number'` or `{ autoNum: '…' }` now emit the next
  number in sequence (1., 2., 3., …; A., B., C., …; i., ii., iii., …)
  rather than a generic dot. Counter resets on a non-numbered paragraph
  or a level change, matching PowerPoint's behaviour.

  Covers the common `ST_TextAutoNumberScheme` tokens — arabicPeriod /
  ParenR / ParenBoth, romanUc / Lc with Period / ParenR / ParenBoth,
  alphaUc / Lc with Period / ParenR / ParenBoth.

- 63c4453: feat: chart axis _tick labels_ honor authored `<c:txPr>` font / color.
  `ChartSpec.categoryAxisLabelStyle` and `valueAxisLabelStyle` carry the
  font / color extracted from `<c:catAx><c:txPr>` and `<c:valAx><c:txPr>`.
  A shared `axisTickAttrs` helper composes the SVG `font-*` / `fill`
  attributes; the value-axis renderer and category-axis renderer both
  project it onto every tick label.
- e60dc26: feat: chart value-axis tick marks. `ChartSpec.valueAxisMajorTickMark`
  and `categoryAxisMajorTickMark` carry `<c:majorTickMark val="in|out|
cross|none"/>`. The playground value-axis renderer draws short stubs
  on the appropriate side of the plot edge (default `out` matches
  PowerPoint's stock look); `none` suppresses them entirely.
- 3a2b974: feat: chart axis titles honor authored `<a:rPr>` font / color.
  `ChartSpec.categoryAxisTitleStyle` and `valueAxisTitleStyle` carry the
  same `ChartTextStyle` shape as `titleStyle`. The playground renderer
  projects size / bold / italic / fill onto both axis title labels,
  sharing the helper that drives the chart title.
- b8c676d: feat(site/playground): bar chart category-axis labels honor
  `categoryAxisLabelRotationDeg`. The horizontal-value renderer used
  the rotation field only for column charts (categories along the
  x-axis); now the bar variant (categories down the y-axis) also
  rotates each label around its anchor and widens its ellipsis budget
  for tilted labels.
- 0b61a96: feat(site/playground): apply the 3-level gradient bg cascade. When
  the slide reports `'gradient'` but doesn't author the actual stops,
  the renderer walks slide → layout → master gradient-fill readers
  to find the inherited gradient definition.
- 2896447: feat: `getSlideLayoutBackgroundImageBytes(pres, layout)` and
  `getSlideMasterBackgroundImageBytes(pres, layout)` complete the
  picture-background cascade. The slide reader already returned bytes
  for slide-level `<a:blipFill>` backgrounds; the new readers resolve
  the same shape on layouts and masters via their own rel lists. The
  playground renderer threads slide → layout → master fallback, so
  template-defined photo backgrounds finally show on inheriting slides.
- 8a4dafb: feat(site/playground): apply the 3-level pattern bg cascade. When the
  slide reports `'pattern'` but doesn't author the actual pattern preset,
  the renderer now walks slide → layout → master pattern-fill readers,
  paralleling the gradient cascade.
- 4b1fd76: feat: `getSlideLayoutBackgroundPatternFill(pres, layout)` and
  `getSlideMasterBackgroundPatternFill(pres, layout)` complete the
  pattern-background cascade. Slides reporting `'pattern'` can now
  resolve the actual preset / colors by walking slide → layout →
  master, paralleling the gradient / solid / blip cascades.
- 63dee61: feat: `getShapeBodyPrEffective(pres, shape)` — `<a:bodyPr>` cascade
  covering anchor, wrap, vertical-text direction, and inset margins.
  Walks shape → layout placeholder → master placeholder bodyPr the same
  way the rPr / pPr cascades do. Playground uses it so placeholders
  inherit text alignment / margins from the layout / master without
  each slide having to re-author them.
- 36bd14d: feat(site/playground): shape text honors `<a:bodyPr wrap="none"/>`.
  The reader's effective wrap value was already threaded through; the
  renderer now emits `white-space:nowrap` when wrap is `'none'`,
  keeping single-line text frames (vertical labels, breadcrumbs,
  fixed-width badges) from wrapping into multiple lines.
- e31b414: feat(site/playground): paragraphs with image bullets (`<a:pPr><a:buBlip>`)
  render a filled-square glyph (■) instead of inheriting the default
  round bullet. The reader already exposed `isParagraphBulletPicture`;
  the playground now threads it through paragraph metadata so the
  visual cue lands.
- e30c9f3: feat: `isParagraphBulletPicture(shape, p)` returns `true` when the
  paragraph uses an image as its bullet (`<a:pPr><a:buBlip>`).
  Renderers without image-bullet support can fall back to a generic
  glyph; UIs that want to indicate the bullet is custom have a clean
  yes/no signal.
- 263bf52: feat: `getParagraphBulletStyle(pres, shape, p)` returns the
  paragraph-level bullet overrides — color (theme-resolved), percent
  size, fixed-point size, font face — from `<a:buClr>` / `<a:buSzPct>`
  / `<a:buSzPts>` / `<a:buFont>`. Playground projects each onto the
  bullet `<span>`, so decks that style bullets in an accent color or
  sized-up font (a common branding move) render correctly instead of
  falling back to the body's color.
- b53b420: feat: chart category-axis tick labels honor `<a:bodyPr rot="N"/>`.
  `ChartSpec.categoryAxisLabelRotationDeg` carries the authored rotation
  (converted from OOXML's 60000ths-of-a-degree to plain degrees). The
  playground renderer rotates each tick label around its anchor and
  shifts the text-anchor side based on the sign of the rotation so dense
  charts with 45°/-45°/90° rotated labels render the way PowerPoint
  shows them. Rotated labels also get a longer truncation budget before
  ellipsization.
- 4f5cc4c: feat: round out gridline color round-trip with 3 more fields —
  `valueAxisMinorGridlineColor`, `categoryAxisMajorGridlineColor`, and
  `categoryAxisMinorGridlineColor`. Previously only the value-axis major
  color was carried. All four now share a new chart-builder
  `gridlinesElement(local, color?)` helper and a chart-reader
  `readGridlineColor(gl)` helper; the existing major-gridline color
  inline parse was replaced with a call to the shared reader for
  consistency.
- 9fba547: feat(chart): `ChartSpec.chartAreaFill` and `plotAreaFill` read
  `<c:chartSpace><c:spPr><a:solidFill>` and `<c:plotArea><c:spPr>
<a:solidFill>`. Playground paints the chart-area backdrop in the
  authored color (replacing the hard-coded white) and adds a tinted
  rect behind the plot area when `plotAreaFill` is authored. Common
  on branded dashboards that paint a subtle background behind the
  chart bars.
- 9d79d53: feat: chart-area / plot-area authored outline strokes.
  `ChartSpec.chartAreaStrokeColor` reads `<c:chartSpace><c:spPr><a:ln>`;
  `ChartSpec.plotAreaStrokeColor` reads `<c:plotArea><c:spPr><a:ln>`.
  The playground renderer projects them onto the chart-area card
  border and the plot-area inner rect — branded charts with thick / no
  / colored card borders finally render the way PowerPoint shows them.
- eb4159e: feat(chart): `ChartSpec.valueAxisHidden` and `categoryAxisHidden`
  read `<c:valAx><c:delete val="1"/>` and `<c:catAx><c:delete val="1"/>`.
  Playground skips rendering the axis when hidden — common on KPI tile
  charts that show just the data points without axis labels.
- 2710f56: feat: `ChartSpec.categoryAxisLineColor` and `valueAxisLineColor` —
  authored stroke color on the axis line itself
  (`<c:catAx|valAx><c:spPr><a:ln><a:solidFill><a:srgbClr val=…/>`).
  `undefined` falls back to the renderer's default. Read by chart-reader,
  written by chart-builder in the correct CT_CatAx / CT_ValAx schema
  order (after the tick-mark elements, before `<c:txPr>`).
- c5761f4: feat(chart): `ChartSpec.categoryAxisOrientation` and
  `valueAxisOrientation` read `<c:catAx>/<c:valAx><c:scaling>
<c:orientation val="minMax|maxMin"/>`. Tools and renderers that
  care about category render order (typically bar charts emit
  `maxMin` so the first category sits at the top) can act on these
  without dropping to XML.
- 45e889d: feat: `ChartSpec.valueAxis` reports the authored
  `<c:valAx><c:scaling>` min / max. Playground respects them when
  computing axis ranges, so charts with a fixed authored scale (e.g.
  percentage charts pinned to 0..100) render with the same scale the
  deck author saw instead of auto-fitting to the data.

  Adds the `ChartAxisScaling` interface to the public type surface.

- 6dd7281: feat: `ChartSpec.categoryAxisTitleRotationDeg` and
  `ChartSpec.valueAxisTitleRotationDeg` — rotation in plain degrees
  (clockwise) on the per-axis title. Maps to
  `<c:catAx|valAx><c:title><c:tx><c:rich><a:bodyPr rot="N"/>` (60000ths
  of a degree on the wire). PowerPoint often emits `-90` on the value-
  axis title; the field now survives round-trip. Read by chart-reader
  via a new `readTitleRotationDeg` helper; written by chart-builder
  through an extended `titleElement(title, style?, rotationDeg?)`
  signature.
- 2bf32ca: feat(chart): `ChartSpec.categoryAxisTitle` and `valueAxisTitle` read
  the per-axis `<c:title>` rich text on `<c:catAx>` (or `<c:dateAx>` /
  `<c:serAx>`) and `<c:valAx>`. Playground paints the value-axis title
  rotated -90° along the y-axis and the category-axis title centered
  below the x-axis.
- db36287: feat: chart builder writes back plot-area / chart-area fill + stroke
  colors. A new `spPrChildren(fill, stroke)` helper emits
  `<c:spPr><a:solidFill><a:srgbClr/></a:solidFill><a:ln>…</a:ln>`. The
  builder appends it under `<c:plotArea>` when `plotAreaFill` or
  `plotAreaStrokeColor` is set, and under `<c:chartSpace>` (root) when
  `chartAreaFill` or `chartAreaStrokeColor` is set. Round-trip test
  verifies all four survive.
- 02434bb: feat: chart builder writes back value-axis extras and tick marks.
  `<c:valAx>` now emits `<c:scaling><c:logBase>`, `<c:majorTickMark>`,
  and `<c:dispUnits><c:builtInUnit>` when authored on `ChartSpec`;
  `<c:catAx>` emits `<c:majorTickMark>`. Round-tripping a chart with
  these fields no longer drops them. Covered by a new round-trip test
  in `fn-chart-readback`.
- 0afa65b: feat: chart builder writes back full value-axis scaling. `<c:valAx>`
  now emits `<c:scaling><c:min/>/<c:max/>`, `<c:numFmt formatCode>`,
  `<c:majorUnit>`, and `<c:minorUnit>` when authored — completing the
  read/write parity for `ChartSpec.valueAxis`. Round-trip test added.
- 7bee159: feat: chart builder writes back axis titles, hidden flags, and
  category-axis tick-label config. `<c:valAx>` / `<c:catAx>` now emit:

  - `<c:title>` with style (from `valueAxisTitleStyle` /
    `categoryAxisTitleStyle`) when an axis title is authored
  - `<c:delete val="1"/>` when `valueAxisHidden` / `categoryAxisHidden`
  - `<c:tickLblPos>` and `<c:tickLblSkip>` when authored on the
    category axis

  Closes the read/write gap for these `ChartSpec` fields. Round-trip
  test added.

- bddd838: feat: chart builder writes back chart-level data-label config. A new
  `dLblsElement` helper builds `<c:dLbls>` with `showVal` / `showCatName`
  / `showSerName` / `showPercent` toggles plus optional `<c:numFmt>`,
  `<c:dLblPos>`, and `<c:separator>`. Wired into every chart variant
  (bar / column / line / pie / doughnut / area), so round-tripping a
  chart with authored data labels preserves them.
- 3f6f848: feat: chart builder writes back per-data-point `<c:dPt>` overrides.
  New `dPtElements(colors, explosions)` helper emits sparse
  `<c:dPt><c:idx><c:bubble3D val="0"/>[<c:explosion>]
[<c:spPr><a:solidFill><a:srgbClr/>]</c:dPt>` entries from
  `ChartSeries.pointColors` and `ChartSeries.pointExplosions`.
  Round-trip test asserts both sparse arrays survive.
- 8a9bab4: feat: chart builder writes back a wide slate of optional chart fields.
  The chart-builder now emits, when authored on `ChartSpec`:

  - `<c:varyColors>` (per chart kind), `<c:gapWidth>`, `<c:overlap>`
    on bar / column
  - `<c:grouping>` honors `ChartSpec.grouping` (`'clustered' | 'stacked'
| 'percentStacked' | 'standard'`) on bar / column / line / area
  - `<c:dropLines>`, `<c:hiLowLines>` on line
  - `<c:firstSliceAng>` on pie / doughnut; `<c:holeSize>` on doughnut
    honors `holeSizePct`
  - `<c:majorGridlines>` (with optional `<c:spPr><a:ln><a:solidFill>`
    color), `<c:minorGridlines>` on the value axis
  - `<c:title><c:overlay>` honoring `titleOverlay`

  Round-trip test asserts the additions all survive read → save →
  reload.

- e7078d7: feat: chart builder writes back legend.textStyle + axis orientation
  reversals. `<c:legend><c:txPr>` now carries the authored font / color
  from `legend.textStyle` (via the existing `axisTxPrElement` helper);
  `<c:scaling><c:orientation>` honors `categoryAxisOrientation` and
  `valueAxisOrientation` (defaulting to `minMax`). Round-trip test
  asserts all four survive read → save → reload.
- 034207d: feat: chart builder writes back `<c:legend>` and `<c:dispBlanksAs>`.
  The chart-root previously emitted only the default legend / blanks
  behavior; the builder now:

  - emits `<c:legend>` with `legendPos`, `overlay`, and one
    `<c:legendEntry><c:idx><c:delete val="1"/></c:legendEntry>` per
    hidden series index — or skips the element when
    `spec.legend.position === null` (author wants no legend)
  - threads `spec.dispBlanksAs` (`'gap' | 'zero' | 'span'`) into
    `<c:dispBlanksAs>`

  Round-trip test added.

- f643b29: feat: chart builder writes back per-series `<c:dLbls>` overrides.
  `dLblsElement` is refactored to take the labels arg directly via
  `buildDLblsFromLabels(dl)`; `seriesElement` now emits per-series
  `<c:dLbls>` when authored, so charts with per-series label toggles /
  numberFormat / position survive round-trip. Round-trip test covers
  all four fields plus the no-override case.
- cf2d02b: feat: chart builder writes back series-level optional fields. Each
  `<c:ser>` now emits:

  - richer `<c:spPr>` with `<a:ln w="…"><a:prstDash/>` when
    `series.lineWidthEmu` or `lineDash` is authored
  - `<c:invertIfNegative val="1"/>` when set
  - `<c:marker><c:symbol/><c:size/></c:marker>` from `markerSymbol` /
    `markerSizePt`
  - `<c:smooth val="1"/>` when set

  Round-trip test covers all five fields.

- 7f0b8ee: feat: chart builder writes back `ChartSpec.titleStyle`. Previously the
  reader picked up authored `<a:rPr sz/b/i><a:solidFill>` on chart
  titles but the builder dropped any incoming style, so round-tripping
  (read → save → reload) lost the title font / color. The builder now
  emits `<a:rPr>` attributes and an inner `<a:solidFill><a:srgbClr/>`
  when a `titleStyle` is provided. New round-trip test
  (`fn-chart-readback`) covers this; total tests 801.
- bf62577: feat: chart builder writes back per-series `<c:trendline>`. A new
  `trendlineElement(tl)` helper emits `<c:trendlineType>`,
  `<c:period>` (movingAvg), `<c:order>` (poly), `<c:forward>` /
  `<c:backward>`, and `<c:spPr><a:ln><a:solidFill>` color when
  authored. Closes the read/write gap for `ChartSeries.trendline`;
  round-trip test covers type / period / forward / backward / color.
- 925fd6f: feat: chart builder writes back axis tick-label style + rotation via
  `<c:txPr>`. New `axisTxPrElement(style, rotationDeg)` helper emits the
  `<c:txPr><a:bodyPr rot/><a:lstStyle/><a:p><a:pPr><a:defRPr…/></a:pPr></a:p></c:txPr>`
  payload from `categoryAxisLabelStyle` / `categoryAxisLabelRotationDeg`
  and the value-axis counterparts. Closes the read/write gap for these
  fields; round-trip test added.
- ba80399: feat: `ChartSpec.categoryAxisMajorGridlines` and
  `ChartSpec.categoryAxisMinorGridlines` — companions to the existing
  `valueAxis*` pair. Bar charts (where the category axis sits on the
  vertical edge) actually use these as horizontal guide lines per
  category band. Mapped to `<c:catAx><c:majorGridlines/>` /
  `<c:minorGridlines/>`. Read by chart-reader, written by chart-builder
  in the correct CT_CatAx schema order (right after `<c:axPos>`).
- 2d61d26: feat: `ChartSpec.categoryAxisLabelOffset` and
  `ChartSpec.categoryAxisLabelAlign` — two more category-axis tuning
  knobs from ECMA-376. `<c:catAx><c:lblOffset val="N"/>` (0..1000, default 100) controls the distance from the axis line to the labels as a
  percent of text size; `<c:catAx><c:lblAlgn val="ctr|l|r"/>` controls
  how multi-line category labels align relative to their tick mark. Both
  are read by chart-reader and written by chart-builder in the correct
  CT_CatAx schema order.
- 21f58cb: feat: `ChartSpec.categoryAxisNoMultiLevelLabel` — toggle multi-level
  (hierarchical) category labels via `<c:catAx><c:noMultiLvlLbl val/>`.
  PowerPoint defaults to `0` (multi-level labels stack); set to `true`
  to flatten hierarchical categories into a single row. Read by
  chart-reader, written by chart-builder at the schema-required last
  position inside `<c:catAx>`.
- c561df4: feat: `ChartSpec.categoryAxisNumberFormat` — number-format code for the
  category-axis tick labels (`<c:catAx><c:numFmt formatCode="…"/>`). Most
  useful on date-style categories (`"mm/dd/yyyy"`, `"mmm-yyyy"`) but
  accepts any Excel format string. Independent of `valueAxis.numberFormat`.
  Read by chart-reader, written by chart-builder in the correct CT_CatAx
  schema order (after `<c:title>`, before `<c:majorTickMark>`).
- 3ecc11b: feat(chart): category-axis label-skip + position. `ChartSpec.categoryAxisTickLabelSkip`
  reads `<c:catAx><c:tickLblSkip val="N"/>` (render every Nth label),
  and `categoryAxisTickLabelPos` reads `<c:tickLblPos val="…"/>`
  (`'none'` hides labels but keeps the axis line; `'low'`/`'high'`/
  `'nextTo'` are the other tokens). Playground honors both — dense
  time-series charts with `tickLblSkip="5"` no longer overlap their
  labels.
- e04ccff: feat: `ChartSpec.dataLabels` carries the chart-level `<c:dLbls>` toggles
  — `showValue`, `showCategory`, `showSeriesName`, `showPercent` — read
  from each plotted-kind element. Playground projects them onto bar /
  column tops (numeric value above each bar) and pie / doughnut slices
  (value, percent, and / or category text painted at the slice mid-arc).

  Adds the `ChartDataLabels` interface to the public type surface.

- 199031b: feat: chart data labels honor `<c:dLbls><c:numFmt formatCode="…"/>`.
  `ChartDataLabels.numberFormat` exposes the format code on both
  chart-level and per-series toggle groups, and the playground renderer
  projects value labels through the same Excel-format subset the value
  axis already supports (`"0%"`, `"$#,##0"`, `"0.00"`, etc). Per-series
  formats win over the chart-level default.
- b77c0ed: feat(chart): `ChartSpec.dispBlanksAs` reads `<c:dispBlanksAs>`
  (`'gap' | 'zero' | 'span'`). Playground line / area renderer:

  - `gap` (default): breaks the path on null values
  - `zero`: substitutes 0 so the line dips to the baseline
  - `span`: connects the surrounding points across the gap

  Previously every null value was coerced to 0, which silently
  flattened the chart whenever the deck had genuine missing data.

- 0eee0da: feat: `ChartDataLabels.textStyle` — the default-run text style for chart
  data labels is now read and written. `<c:dLbls><c:txPr><a:defRPr/>`
  is parsed into `ChartTextStyle` (sizePt / bold / italic / color) and
  emitted in CT_DLbls schema order (after `<c:numFmt>`, before
  `<c:dLblPos>`). Both the chart-level `dataLabels` and per-series
  `series[i].dataLabels` honor the field.
- 17f57b3: feat: `ChartSeries.pointColors` — sparse map of per-data-point fill
  overrides read from `<c:ser><c:dPt><c:spPr><a:solidFill>`. Pie /
  doughnut decks almost always emit one of these per slice; the playground
  now paints each slice in its authored color (and reflects it in the
  legend swatches) rather than cycling through the accent palette.
- 4d2cecb: feat(chart): `ChartSpec.dropLines` and `hiLowLines` read
  `<c:dropLines>` and `<c:hiLowLines>` on line / area / stock plots.
  Playground renders drop lines from each first-series data point down
  to the value axis (dashed gray) and hi-low lines as a vertical span
  between the highest and lowest series value at each category
  (solid darker gray). The latter is the canonical OHLC pattern.
- 263bf52: feat: chart reader now recognises scatter, bubble, radar, stock, and
  (2D / 3D) surface charts and degrades them to the closest already-
  modelled kind so renderers paint something useful instead of the
  "unsupported chart kind" placeholder. Scatter / bubble series read
  their `<c:yVal>` channel; their `<c:xVal>` / `<c:bubbleSize>` are
  not yet surfaced.
- 0a9236f: feat(chart): `ChartSpec.gapWidthPct` and `overlapPct` read from
  `<c:gapWidth>` and `<c:overlap>` on bar / column plots. Playground
  sizes bars per ECMA-376 §21.2.2.75 — `barW = groupW / (clusterUnits +
gapWidth/100)` with `clusterUnits = 1 + (S - 1)(1 - overlap/100)` —
  so authored bar spacing matches PowerPoint instead of the hard-coded
  0.8 / 0.7 ratios.
- b88dbb8: feat(chart): `ChartSpec.valueAxisMajorGridlines` / `valueAxisMinorGridlines`
  read the presence of `<c:majorGridlines/>` / `<c:minorGridlines/>`
  under `<c:valAx>`. Playground hides gridlines when `majorGridlines`
  is explicitly `false` (absent in the source XML) — common on KPI
  charts that show clean bars / lines without horizontal rules behind
  them. Tick labels still render.
- 4caa5ad: feat(chart): `ChartSeries.invertIfNegative` reads `<c:ser>
<c:invertIfNegative val="1"/>`. Playground's bar / column renderer
  paints negative bars in a darker shade of the series color when the
  flag is set — matching PowerPoint's profit/loss visualization.
- b603115: feat: `ChartSpec.language` (`<c:chartSpace><c:lang val=…/>`) and
  `ChartSpec.date1904` (`<c:date1904 val=…/>`) — chartSpace-level Office
  metadata round-tripped for parity. `language` is the Office UI
  language code (e.g. `'en-US'`, `'ja-JP'`); `date1904` selects the
  1904 date epoch (default `false` = Excel 1900 epoch, surface only
  when explicitly true). pptx-kit's renderers don't act on either yet.
- 028e3b7: feat: chart `<c:legend><c:legendEntry><c:delete val="1"/>` honored.
  `ChartSpec.legend.hiddenIndices` carries the series indices the
  author wants suppressed from the legend (typically trendline series).
  The playground filters the parallel legend arrays (names, colors,
  marker glyphs) in lock-step so the remaining entries stay aligned,
  without affecting plotted data.
- c141173: feat(chart): `ChartSpec.legend` carries the `<c:legend><c:legendPos>`
  token — `'r' | 't' | 'b' | 'l' | 'tr'`. Playground projects each
  onto the appropriate edge (horizontal row for top / bottom, vertical
  stack for the side / corner positions). Charts whose `<c:legend>`
  sets `position` to `null` paint without a legend.
- 9a49faf: feat(chart): `ChartAxisScaling.majorUnit` and `minorUnit` read
  `<c:valAx><c:majorUnit>` / `<c:minorUnit>` tick spacing. Playground's
  value-axis renderer emits ticks at each multiple of the authored
  majorUnit instead of nice-rounded auto-ticks when present.
- f99d548: feat: `ChartSpec.valueAxisMinorTickMark` and `categoryAxisMinorTickMark`
  — minor-tick-mark mode siblings of the existing `*MajorTickMark` pair.
  Maps to `<c:catAx><c:minorTickMark val="in|out|cross|none"/>` and the
  value-axis equivalent. Read by chart-reader, written by chart-builder
  in the correct schema order (right after `<c:majorTickMark>`).
- 28d77ea: fix: chart categories accept `<c:cat><c:numRef>` (numeric / date
  categories). Previously the category-axis dropped to an empty
  labels array when the chart authored a numeric category channel
  (common for date-axis line charts authored in Excel). Falls back
  to formatting each cached numeric value as a string so date /
  number cats appear on the axis instead of disappearing.
- 7b3ba0a: feat(chart): axis number formats now accept Excel's `"$"#,##0`
  quoted-literal prefix / suffix syntax. PowerPoint typically emits
  currency as `"$"#,##0` (or `"\$"#,##0`) rather than the bare `$`
  form, so the previous detection missed it.
- d2f86d2: feat(chart): `ChartAxisScaling.numberFormat` reads `<c:valAx><c:numFmt
formatCode="…"/>`. Playground projects the most common Excel format
  codes to axis labels — percent (`'0%'`, `'0.0%'`), thousand
  separator (`'#,##0'`, `'#,##0.0'`), and currency prefixes
  (`'$#,##0'`, `'¥#,##0'`). Other codes fall through to the generic
  auto-formatted label.
- 3efdbeb: feat(chart): `ChartSpec.titleOverlay` and `ChartSpec.legend.overlay`
  read `<c:title><c:overlay>` / `<c:legend><c:overlay>`. When `true`,
  the title / legend sits on top of the plot area instead of taking a
  horizontal strip. Playground sizes the plot area accordingly — gives
  the chart back the extra vertical real estate when overlay is set.
- 4cde872: feat: `ChartSpec.plotVisibleCellsOnly` — toggle `<c:plotVisOnly val/>`.
  PowerPoint's default is `true` (only plot visible cells); the field
  exists to let authors opt into `false` (plot hidden rows / columns too).
  The reader surfaces `false` only when the wire is explicitly `0` so
  round-tripping the common default doesn't drag a redundant explicit
  `true` into the spec.
- 693ba3e: feat: `ChartSpec.roundedCorners` — round-trip the chartSpace-level
  `<c:roundedCorners val>` toggle. PowerPoint's default is `false`; the
  reader surfaces `true` only when the wire is explicitly `1` and the
  builder emits the element only when authored, so common defaults stay
  clean. Schema position is BEFORE `<c:chart>` (per CT_ChartSpace).
- 733120a: feat(chart): `ChartSeries.smooth` reads `<c:smooth val="1"/>`. Playground
  line / area renderer interpolates a cubic-Bézier curve through the
  data points (Catmull-Rom-to-Bezier with 0.5 tension) when `smooth` is
  true, matching PowerPoint's "smooth line" preset visually.
- d581121: feat(site/playground): bar (horizontal), line, and area charts now
  honour `ChartSpec.grouping` for stacked / percentStacked layouts —
  matching the column-chart treatment added previously. Data labels
  render inside the stacked segments for bar (white bold), at the
  appropriate cumulative position for line / area, and percent-stacked
  versions normalize each category to 100%.
- 33c2c11: feat: `ChartSpec.grouping` carries the `<c:grouping>` token —
  `'clustered' | 'stacked' | 'percentStacked' | 'standard'`. Playground
  column chart renders stacked / percent-stacked layouts: series stack
  within each category, and percent-stacked normalises to 0..100% per
  column with in-bar value labels.

  Adds the `ChartGrouping` type to the public surface.

- 53148e4: feat: `ChartSpec.chartStyle` — round-trip the chartSpace-level
  `<c:style val="N"/>` PowerPoint chart-style preset (1..48). Encodes a
  curated combo of theme accent colors, gradients, effects, and font
  sizes from the PowerPoint "Chart Styles" gallery. Read and written for
  round-trip parity; pptx-kit's renderers don't interpret the preset
  yet, but the field survives save/reload.
- b1cfda3: feat: `ChartSpec.categoryAxisTickMarkSkip` — the second half of the
  ECMA-376 `<c:catAx>` skip pair. `<c:tickLblSkip>` (already supported)
  controls label-skip stride; `<c:tickMarkSkip val="N"/>` independently
  draws every Nth tick mark. Useful when you want fewer label collisions
  but the same dense tick lattice. Read by chart-reader and written by
  chart-builder.
- 2599a46: feat: chart titles read `<c:tx><c:strRef>` workbook-cell references.
  Previously only literal `<c:rich>` titles surfaced; titles authored
  via Excel's "Link to source cell" wizard (which emits `<c:strRef>`
  with a `<c:strCache>` of the resolved text) now flow through to
  `ChartSpec.title` as the cached value. Affects the title shown above
  the chart and, transitively, axis-title rendering.
- da1e50d: feat: chart titles honor `<a:rPr>` font / color overrides.
  `ChartSpec.titleStyle` carries the authored size (in pt), bold, italic,
  and fill color extracted from the title's first `<a:r><a:rPr>` (or
  `<a:pPr><a:defRPr>` as fallback). The playground renderer projects
  those through to the SVG `<text>`. Templates that brand their chart
  titles to a non-default size / color finally render with the authored
  look.
- 5f84cfc: feat: `ChartTrendline.displayEquation` and `ChartTrendline.displayRSquared`
  — two booleans that toggle the regression-equation label and R²
  coefficient overlay next to a trendline. Map to
  `<c:trendline><c:dispEq val="1"/>` and `<c:dispRSqr val="1"/>`. Read by
  chart-reader; written by chart-builder in the correct CT_Trendline
  schema order (after `<c:backward>`, before any `<c:trendlineLbl>`).
- a978251: feat: `ChartTrendline.name` — round-trip a custom trendline label
  (`<c:trendline><c:name>…`). PowerPoint auto-generates a label like
  "Linear (X)" or "MA(5) (X)" when this element is omitted; authors who
  want a different label (or who imported one from another tool) now
  have the field. Read by chart-reader; written by chart-builder at the
  CT_Trendline schema-required first position (before `<c:spPr>`).
- 57eeffa: feat(chart): `ChartSeries.trendline` reads `<c:trendline>` —
  regression type (linear / exp / log / poly / power / movingAvg),
  moving-average period, polynomial order, and the trendline's stroke
  color. Playground overlays a dashed trendline on bar / column / line
  charts; linear / log / exp use fitted regressions, movingAvg uses a
  rolling mean.

  Adds the `ChartTrendline` type to the public surface.

- 24b2794: feat: `ChartSpec.valueAxisCrossBetween` — controls whether the value
  axis crosses the category axis _between_ tick marks (the default for
  bar/column/area) or _at_ each tick mark (the default for line/scatter).
  Maps to `<c:valAx><c:crossBetween val="between|midCat"/>`. Read by
  chart-reader, written by chart-builder in the correct CT_ValAx schema
  order (after `<c:crossesAt>`).
- 6a236be: feat: `ChartSpec.valueAxisCrosses` — controls where the category axis
  crosses the value axis. Accepts either an enum keyword
  (`'autoZero' | 'min' | 'max'` → `<c:valAx><c:crosses val=…/>`) or a
  numeric tagged form (`{ at: N }` → `<c:valAx><c:crossesAt val=N/>`).
  The two forms are mutually exclusive per the schema; `crossesAt` wins
  when both are present on read. Read by chart-reader, written by
  chart-builder in the correct CT_ValAx schema order (after `<c:crossAx>`).
- b2c1304: feat: chart `<c:varyColors>` for single-series bar / column.
  `ChartSpec.varyColors` carries the `<c:plottedKind><c:varyColors val="1"/>`
  flag. When set and the chart has exactly one series, the renderer
  assigns each data point a distinct accent color (mirroring
  PowerPoint's "Vary colors by point" toggle for column / bar). Pies
  already varied colors implicitly.
- 69431a9: feat: `getSlideColorMapOverride(slide)` returns the slide's
  `<p:clrMapOvr><a:overrideClrMapping/>` token-remap, or `null` when the
  slide inherits the master's color map. Returned as a plain `Record`
  with the eight stable tokens (`bg1`, `tx1`, `bg2`, `tx2`, `accent1`-
  `accent6`, `hlink`, `folHlink`) keyed to their override targets.
  Useful for renderers that need to know when a slide reinterprets the
  theme's color story.
- 263bf52: feat: apply ECMA-376 §20.1.2.3.x color transforms when resolving colors.

  - New `resolveDrawingColor(colorEl, theme)` resolves any DrawingML color
    element (`<a:srgbClr>` / `<a:schemeClr>` / `<a:sysClr>` / `<a:prstClr>`)
    with all transform children (`<a:lumMod>`, `<a:lumOff>`, `<a:shade>`,
    `<a:tint>`, `<a:satMod>` / `Off`, `<a:hueMod>` / `Off`, `<a:gray>`,
    `<a:inv>`, `<a:comp>`) applied. Scheme tokens are looked up against
    the supplied theme.
  - New `getShapeFillColorResolved(pres, shape)` and
    `getShapeStrokeColorResolved(pres, shape)` return the exact `#RRGGBB`
    PowerPoint paints — useful for renderers / exporters where the legacy
    `getShapeFillColor` / `getShapeStrokeColor` strings (`#RRGGBB` or
    `scheme:<token>`) miss both scheme resolution and color transforms.
  - `getShapeRunFormatEffective` now applies the same pipeline at every
    layer of the rPr cascade, so a run inheriting `accent1 lumMod=40000
lumOff=60000` (PowerPoint's "Accent 1, Lighter 60%") resolves to the
    concrete tinted hex instead of leaking the raw token through.

- 263bf52: feat(site/playground): bent / curved connector routing.

  `bentConnector{2,3,4,5}` render as the matching L / Z / two-step /
  three-step paths, and `curvedConnector{2,3,4,5}` render as quadratic
  / cubic Bézier curves between the connector's bounding-box endpoints.
  Previously every connector preset projected to a straight line; flow-
  chart and diagram decks now show the right cadence.

- e65228f: feat: read custom geometry. New `getShapeCustomGeometry(shape)` returns a
  shape's `<a:custGeom>` (ECMA-376 §20.1.9) as a fully-evaluated path list —
  guide formulas (`avLst`/`gdLst`, all §20.1.9.11 operators) are resolved
  against the shape extents so the returned `moveTo`/`lnTo`/`arcTo`/
  `quadBezTo`/`cubicBezTo`/`close` commands carry only numbers. The preview
  renderer now draws custom geometry as a real SVG path (arcs converted to
  cubic Béziers) instead of a labelled rectangle placeholder; only a custGeom
  that fails to evaluate still falls back, marked `data-pptx-fallback`.
- 1e774b8: feat(site/playground): pie / doughnut / line / area honor
  `<c:dLblPos>` data-label positions. Pie supports `ctr` (default
  midline), `inEnd` (just inside the rim), and `outEnd` (outside the
  pie, with a darker fill so it shows on the chart-area backdrop). Line
  and area chart per-point labels honor `ctr`, `t` (default), `b`, `l`,
  `r`. Column / bar already shipped in the prior batch.
- fc019d1: feat: chart data label position. `ChartDataLabels.position` carries the
  `<c:dLbls><c:dLblPos val="…"/>` token (typed as
  `ChartDataLabelPosition`). The reader extracts it at both chart-level
  and per-series scope. The playground renderer projects `ctr`, `inEnd`,
  `outEnd`, `inBase` onto clustered column and bar labels — outside-end
  remains the default, but authored positions now move labels inside the
  bar or to the base as PowerPoint shows them.
- 3cfba8d: feat: chart data label separator. `ChartDataLabels.separator` carries
  the `<c:dLbls><c:separator>…</c:separator>` text used to join
  multiple label parts (value + percent + category etc.). The pie /
  doughnut renderer threads the per-series override, falling back to
  the chart-level separator and finally to a single space. Common
  values: `", "`, `"\n"`, `"; "`.
- 003e7b5: feat: density-array companions for tables and images —
  `getPresentationTableCountsBySlide(pres)` and
  `getPresentationImageCountsBySlide(pres)`. Both return a dense
  per-slide count array (0 for slides without that asset kind),
  matching the shape / chart / comment / text-length counterparts.
  Completes the deck-density family.
- fd1519a: feat(site/playground): `<c:dispUnits>` value-axis label. When the
  chart authors a display-units token (`thousands`, `millions`, etc.)
  the value-axis now emits an italic "Thousands" / "Millions" /
  … label rotated alongside the axis (vertical orientation) or to
  the right of the rightmost tick (horizontal). Completes the
  display-units rendering — values are already divided, and now the
  scale self-describes.
- c5f0b60: feat: chart value-axis honors `<c:dispUnits><c:builtInUnit/>`.
  `ChartAxisScaling.displayUnits` carries the authored scale token
  (`hundreds`, `thousands`, `millions`, etc.). The playground divides
  each value-axis tick by the corresponding divisor before formatting,
  so charts authored "in millions" finally render as `10` / `20` /
  `30` instead of `10000000`.
- 263bf52: feat: add `getShapeRunFormatEffective(pres, shape, p, r)` — resolves a
  run's character properties (font, size, color, bold, italic, underline)
  through the full ECMA-376 §21.1.2.4.7 inheritance chain: run `<a:rPr>` →
  `<a:endParaRPr>` → paragraph `<a:defRPr>` → text-body `<a:lstStyle>` →
  matching layout placeholder → matching master placeholder → master
  `<p:txStyles>` → theme `<a:fontScheme>`. Theme tokens like `+mj-lt` are
  expanded to the deck's major/minor typefaces. The existing
  `getShapeRunFormat` still returns the literal `<a:rPr>` only.
- 25654cf: feat: `getShapeEffectsEffective(pres, shape)` walks the layout →
  master placeholder cascade for `<a:effectLst>`. Effect lists override
  rather than compose (matching PowerPoint's behaviour), so the first
  layer that supplies any effects wins. Playground uses it so
  placeholder shadows / glows / soft edges inherited from the master
  finally render on slides that don't repeat the effect list.
- acc9b15: feat: effects & fills polish. The reflection effect (`a:reflection`) now
  renders as a vertically mirrored, gradient-masked copy honoring start/end
  alpha, distance, and the signed `sy` scale; picture bullets (`a:buBlip`)
  render as real inline images in both text layout modes via the new core
  reader `getParagraphBulletImageBytes` (the "■" fallback remains only when
  bullet bytes are genuinely unavailable); and gradient fills inherited
  through the placeholder layout/master cascade resolve via the new
  `getShapeGradientFillEffective` instead of painting a hardcoded orange
  tint.
- 263bf52: feat: `getShapeEffects(pres, shape)` returns every effect on the
  shape's `<a:effectLst>` (`outerShdw`, `innerShdw`, `glow`, `reflection`,
  `softEdge`, `blur`) in document order, with each effect's color
  (transform-resolved against the theme), opacity, blur radius, distance,
  and angle. PowerPoint composes multiple effects in a single filter
  stack — the existing `getShapeEffect` only surfaced the first one.

  The playground renderer now emits an SVG `<filter>` chain that
  composes the same effects, including a synthesized inner shadow
  (SVG has no `feInnerShadow` primitive — built via offset + composite).

  Also adds the `ShapeEffectAny` type union to the public surface.

- 18d3ceb: feat: `getShapeFillEffective(pres, shape)` walks the layout → master
  placeholder cascade when the shape's own fill is `'inherit'`. Returns
  the first non-inherit fill found. Playground reaches for it as its
  primary fill source so placeholder fills authored on the master /
  layout finally show through.
- 81ce680: feat: `findShapesByPreset(slide, preset)` returns every shape whose
  `<a:prstGeom prst="…"/>` matches. Useful for diagram introspection:
  find all `'leftArrow'`s for a workflow swap, replace every `'cloud'`
  with `'rect'`, etc. Shapes without a preset (custGeom / pictures /
  charts / tables / connectors / groups) are filtered out.
- 019a934: feat: `findChartsWithDataLabels(slide)` — slide-scoped auditor for
  charts whose chart-level or per-series `dataLabels` enable at least
  one of `showValue` / `showCategory` / `showSeriesName` / `showPercent`.
  Purely presence-based; doesn't validate numberFormat or position.
  Charts whose kind isn't modeled are skipped.
- cb5d037: feat: `findChartsWithTrendlines(slide)` — slide-scoped finder for
  charts that carry at least one `<c:trendline>` on any of their
  series. Useful for deck-audit reports — trendlines are easy to add
  and easy to forget. Charts whose kind isn't modeled are skipped.
- 1653804: feat: `findCommentsByAuthor(pres, authorName)` and
  `findSlidesWithCommentsByAuthor(pres, authorName)` now accept a
  `RegExp` as well as a literal string. Useful for "every comment from
  review bots" (`/^review-bot/`) or "every comment from anyone with a
  given email domain" patterns. Backward compatible — string callers
  still get exact-equality matching.
- b6d9ea4: feat: `findShapeByName(slide, name)` now accepts a `RegExp` as well
  as a literal string. Mirrors the RegExp support just landed on
  `findShapesByName` (multi-match). Returns the first match in document
  order; backward compatible with existing string callers.
- 7e59ac4: feat: `findShapeInPresentation(pres, name)` now accepts a `RegExp` as
  well as a literal string. Mirrors the RegExp support on the
  slide-scoped `findShapeByName`. Backward compatible — string callers
  still get exact-equality.
- 70a2327: feat: `findShapesByEffect(pres, slide, kind)` — returns every shape on
  the slide whose `<a:effectLst>` carries an effect of the given `kind`
  (`'outerShdw'`, `'innerShdw'`, `'glow'`, `'reflection'`, `'softEdge'`,
  `'blur'`). Pure presence check; doesn't walk the layout / master
  cascade. Useful for "which shapes have a shadow / glow on this
  slide?" audits.
- 94c4480: feat: `findShapesByHyperlink(slide, url)` — slide-scoped finder that
  returns every shape whose hyperlink target matches `url` (substring or
  `RegExp`). Pairs the existing presentation-level
  `findSlidesByHyperlink` for cases where the caller already has a
  specific slide and wants the linking shapes inside it.
- e71664d: feat: `findShapesByName(slide, name)` now accepts a `RegExp` as well
  as a literal string. Useful when template-cloned shapes share a
  prefix (`'TextPlaceholder1'`, `'TextPlaceholder2'`, …). Backward
  compatible — string callers still get exact-equality matching.
- f57a4ab: feat: `findShapesInRect(slide, x, y, w, h)` — marquee-style region
  finder. Returns every shape whose bounds overlap the rectangle
  (touching edges count). Shapes with no resolvable bounds are skipped.
  Companion to `findShapesAtPoint(slide, x, y)` for cases where the
  caller wants a region of the slide rather than a single point.
- a7c00cb: feat: `findShapesWithAnimation(slide)` — returns every shape on the
  slide whose `getShapeAnimation` is not `null`. Pair to
  `slideHasAnimations`. Useful for "which shapes on this slide actually
  animate?" audits before exporting to a video pipeline that doesn't
  honor PowerPoint's timing tree.
- 87d7fbb: feat: `findShapesWithHyperlinks(slide)` — every shape on the slide
  that carries any hyperlink, regardless of target. Counterpart to
  `findShapesByHyperlink(slide, url)` (which requires a matching URL)
  for "audit every clickable shape on this slide" workflows.
- 5cc4f75: feat: `findSlideByTitle(pres, title)` now accepts a `RegExp` as well
  as a literal string. Pairs the RegExp support on
  `findSlidesByText` / `findShapeByName` / `findCommentsByAuthor`.
  Backward compatible — string callers still get exact-equality.
- 134943d: feat: `findSlidesByLayoutPartName(pres, layoutPartName)` — finds every
  slide whose resolved layout part name matches `layoutPartName` (e.g.
  `'/ppt/slideLayouts/slideLayout3.xml'`). Pair to the existing
  `findSlidesByLayoutName` / `findSlidesByLayoutType`. Keyed on the
  actual package path, so it's stable across template-name collisions
  and PowerPoint UI locales.
- 20613b5: feat: `findSlidesWithChartKind(pres, kind)` — kind-filtered variant of
  the existing `getSlidesWithCharts`. Returns every slide carrying at
  least one chart of the given `ChartKind` (`'bar'`, `'column'`,
  `'line'`, `'pie'`, `'doughnut'`, `'area'`). Built on `getSlideCharts`
  so the predicate respects the spec the renderers actually see.
- 7c545c9: feat: `findSlidesWithChartTrendlines(pres)` — deck-level variant of
  the slide-scoped `findChartsWithTrendlines`. Returns every slide
  carrying at least one chart with a trendline on any series. Useful
  for "audit every trendline in this deck" workflows before publishing.
- 666343d: feat: `getEmptySlides(pres)` — returns every slide whose `<p:spTree>`
  carries no shapes (per `getSlideShapes`). Useful as a pre-publish
  "find the section dividers I forgot to fill in" check.
- b4dbcc0: feat: `getPresentationChartCountsBySlide(pres)` — dense per-slide chart
  count array. Counts every chart returned by `getSlideCharts` regardless
  of whether its spec parsed; pair with `getPresentationChartKindCounts`
  for kind-level totals. Rounds out the density-array family alongside
  `getPresentationCommentCountsBySlide`,
  `getPresentationShapeCountsBySlide`, and
  `getPresentationTextLengthsBySlide`.
- a4ca6ca: feat: `getPresentationChartKindCounts(pres)` — deck-wide histogram of
  `ChartKind` → count. Returns a frozen `Record` with every kind
  present (zeros for absent kinds), so destructuring and chart-style
  audits stay typed without runtime checks. Charts whose spec doesn't
  parse are skipped, matching `findChartByKind` / `findSlidesWithChartKind`.
- f7dbcc4: feat: `getPresentationCommentCountsByAuthor(pres)` — deck-wide
  histogram of comment counts keyed by author display name. Useful for
  "who reviewed this deck the most?" audits. Authors sharing a display
  name get merged into the same bucket; pair with
  `getPresentationCommenters` when authors with identical names need to
  be kept separate by `id`.
- be1a608: feat: `getPresentationCommentCountsBySlide(pres)` — dense per-slide
  comment count array. Every slide appears as an element (count `0`
  when the slide has no comments), so callers can chart comment
  density per slide without re-indexing.
- 2789bb3: feat: `getPresentationHyperlinkCountsBySlide(pres)` — dense per-slide
  hyperlink count array. Counts shapes whose `getShapeHyperlink` is
  non-null. Cheaper than `getAllHyperlinks` when the caller only wants
  per-slide counts. Rounds out the deck-density family.
- 2f67bd7: feat: `getPresentationNotesLengthsBySlide(pres)` — dense per-slide
  speaker-notes length array. Pair with
  `getPresentationTextLengthsBySlide` for handout / talk-track audits —
  slides with little on-screen text but heavy notes are usually the
  slow part of a presentation.
- 1974b01: feat: `getPresentationShapeCountsBySlide(pres)` — dense per-slide
  shape count array. Counts whatever `getSlideShapes` flattens (top-
  level + group-children). Useful for charting shape density per slide
  and identifying outliers for cleanup.
- 6569f9d: feat: `getPresentationTextLengthsBySlide(pres)` — dense per-slide
  visible-text length array. Counts code points (surrogate pairs as 1)
  per `getSlideTextLength`. Pair with `getPresentationShapeCountsBySlide`
  for slide-density audits.
- b793c74: feat: `getSlideLayoutUsageCountsByType(pres)` — companion to
  `getSlideLayoutUsageCounts`, but keyed on the OOXML layout-type enum
  token (`title`, `obj`, `twoObj`, `blank`, …) instead of the user-
  visible name. Stable across PowerPoint UI locales. Useful for "how
  many content slides vs. dividers vs. title slides?" audits.
- 3891fa2: feat: `getSlideLayoutUsageCounts(pres)` — layout name → number-of-slides
  histogram. Every layout enumerated by `getSlideLayouts` appears as a
  key (count `0` for unreferenced layouts), so the function surfaces
  unused layouts directly — useful for trimming template decks that
  ship with placeholder layouts the working deck never picks up.
- aad46e4: feat: `getSlideMasterUsageCounts(pres)` — master part name → number of
  slides chaining to that master. Every master in the package appears as
  a key (count `0` for unreferenced masters), so it surfaces unused
  masters directly. Pair with `getSlideLayoutUsageCounts` for the
  layout layer in multi-master template decks.
- 02fe159: feat: `getSlideTables(slide)` — returns every table graphic-frame
  shape on the slide, in document order. Pair to `getSlideCharts` for
  cases where the caller wants just the tables; convenience over
  `getSlideShapes(slide).filter(isTableShape)`.
- acf5880: feat: `getUnusedSlideLayouts(pres)` — returns the layouts in the
  package that no slide references. Useful when trimming a template
  deck — unused layouts contribute parts and rels without ever
  rendering. Iteration order matches `getSlideLayouts`.
- eca84ce: feat: `getUnusedSlideMasters(pres)` — master part names that no slide
  chains to (count of `0` in `getSlideMasterUsageCounts`). Pair to
  `getUnusedSlideLayouts`. Useful when trimming multi-master template
  decks of dead theme variants.
- 263bf52: feat: `getShapeGradientFill` now surfaces non-linear gradient paths
  (`<a:path path="circle|rect|shape">`) and the `<a:fillToRect>` focus
  rectangle. `GradientFillOptions` gains `path` and `focus` fields so
  renderers can reproduce radial, rectangular, and shape-following
  gradients instead of falling back to a linear approximation.

  The playground renderer emits an SVG `<radialGradient>` for the
  non-linear paths, with reversed stop offsets so the first ECMA-376
  stop sits at the focus center (matching PowerPoint's outward
  painting order).

- 855076d: feat: chart value-axis major gridlines honor authored stroke color.
  `ChartSpec.valueAxisMajorGridlineColor` extracts the
  `<c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr/>` color and
  the playground renderer paints gridlines with it (falls through to the
  existing light-gray default when no color is authored). Branded
  templates with custom gridline tints finally render correctly.
- 74b227e: feat(site/playground): hyperlink tooltips. Shape and per-run
  hyperlinks now surface their `<a:hlinkClick tooltip="…"/>` text —
  shapes get an SVG `<title>` child on the `<a>` wrapper, runs get a
  `title=` attribute on the HTML anchor. PowerPoint shows these on
  hover during the slideshow; the playground now does too.
- a610e82: feat: `getShapeHyperlinkTooltip(shape)` and
  `getShapeRunHyperlinkTooltip(shape, p, r)` return the
  `<a:hlinkClick tooltip="…"/>` text. Tooltips show up in PowerPoint
  when the user hovers a linked shape in slide-show mode — useful for
  accessibility and link-preview surfaces.
- cbdda7c: feat(site/playground): render `<a:duotone>` image recolor. The filter
  pipeline desaturates the picture to luminance, then samples a
  two-color gradient (firstColor → secondColor) via a 16-step
  `feComponentTransfer` table. Pictures with PowerPoint's Color >
  Recolor preset finally render in their authored two-color tint.
- c4a89c1: feat: `getShapeImageDuotone(pres, shape)` reads the picture's
  `<a:blip><a:duotone>` two-color recolor effect — the typical
  "Picture Tools > Recolor" output. Returns the two hex-resolved
  colors (or `null` for each that the duotone didn't author). Lets
  downstream renderers project the duotone via SVG `<filter>` or
  inform consumers that the picture has a color-replacement applied.
- 99fcb65: feat: image color-effect readers — `isShapeImageGrayscale(shape)`
  detects `<a:blip><a:grayscl/>` (Color > Grayscale), and
  `getShapeImageBiLevelThreshold(shape)` returns the threshold percent
  for `<a:blip><a:biLevel thresh="…"/>` (Color > Black and White).
  Renderers can project these onto CSS / SVG filters.
- 263bf52: feat: `getShapeImageLinkUrl(shape)` returns the external URL of a
  picture whose `<a:blip>` carries an `r:link` (Insert > "Link to file")
  instead of `r:embed`. Bytes for these aren't in the package; the
  playground now shows the linked URL in the placeholder rather than a
  generic "no bytes" label.
- 508627a: feat(site/playground): grayscale + biLevel image filters in the
  playground. The filter pipeline now composes:

  1. brightness + contrast (linear feComponentTransfer)
  2. grayscale (luminance-preserving feColorMatrix) when
     `<a:blip><a:grayscl/>` is set
  3. biLevel two-tone (discrete tableValues snapped at the authored
     threshold) when `<a:blip><a:biLevel thresh="…"/>` is set

  Pictures with Color > Grayscale or Color > Black and White now
  render with the same visual treatment PowerPoint shows.

- 66edcbc: feat: add `isShapeTextBox(shape)` — `true` when a shape is a text box
  (`<p:cNvSpPr txBox="1">`) rather than an autoshape. The two have different
  default text formatting (text boxes left/top, autoshapes center/middle), so
  renderers and layout code need to tell them apart.
- 263bf52: feat: `getSlideLayoutBackground(layout)` mirrors `getSlideBackground`
  for slide layouts. Playground falls back to it when the slide's own
  background reports `'inherit'`, so brand-color or template backgrounds
  authored on the layout actually paint behind slides that don't override
  the bg themselves.
- 2a1d712: feat: `getSlideLayoutBackgroundGradientFill(layout)` returns the
  gradient definition when a layout's background is
  `<p:bgPr><a:gradFill>`. Same shape as the slide-level variant —
  renderers can reuse the same projection logic for layout gradient
  backgrounds via the shared `gradientDef` helper.
- a1229d5: feat: `getSlideLayoutBackgroundShapes(pres, layout)` returns the
  non-placeholder shapes on a layout as a render-ready view
  (`SlideLayoutBackgroundShape` — bounds, preset, fillHex, strokeHex,
  strokeWidthEmu, rotation, flip). Playground paints them behind the
  slide's own shapes so brand-template decoration (corner bars, divider
  lines, background rectangles) shows through on slides that don't
  redefine the layout themselves.

  Adds the `SlideLayoutBackgroundShape` type to the public surface.

- ba056db: feat: `getSlideLayoutBackground` now handles `<p:bgRef>` the same way
  `getSlideBackground` does. Layouts in real brand templates almost
  always reference the theme via `<p:bgRef>` rather than authoring an
  explicit `<p:bgPr>` — picking up the inner color element as a solid
  fill closes the cascade so the playground paints the right brand
  color even when the slide's own background reports `inherit`.
- 3ea2ed5: feat(site/playground): line / area chart legend swatches use the
  series marker glyph. The legend previously rendered every series as
  a 9×9 square color rect. For `line` / `area` charts the renderer now
  passes the per-series `markerSymbol` (circle / square / diamond /
  triangle / star / x / plus / dash / dot) so legend entries match
  the data points. Bar / column / pie keep the square swatch.
- b1073ff: feat(site/playground): right / left chart legend stack centers
  vertically. Previously the `r` and `l` legend positions both
  stacked from a fixed `f.y + 12` top, the same as `tr`. PowerPoint
  vertically-centers right / left legends inside the chart area; the
  renderer now matches by computing `yStart` from the legend's total
  height. `tr` keeps the top-anchored stack.
- ee27024: feat: chart legend honors authored `<c:txPr>` font / color.
  `ChartSpec.legend.textStyle` carries the same `ChartTextStyle` shape
  used for the chart title and axis titles. The playground renderer
  projects font-size, bold, italic, and fill color onto every legend
  label across all four position layouts (right / left / top / bottom /
  top-right).
- fca13ca: feat(site/playground): line and area charts paint per-point value labels
  when `<c:dLbls><c:showVal val="1"/>` is set. Labels sit just above each
  marker and route through the chart number-format projector (so
  `<c:numFmt formatCode="0%"/>` etc. apply the same as on bar / pie).
  Honors the per-series → chart-level cascade.
- 263bf52: feat: `getParagraphLineSpacing(shape, p)` returns the paragraph's
  `<a:lnSpc>` as `{ kind: 'pct' | 'pts', value }`. Percent values come
  through as a unit fraction (1.5 = 150%); point values are pt.

  The playground projects both forms to CSS `line-height` per paragraph,
  and uses the existing `getParagraphSpacing` to project spcBef / spcAft
  to `margin-top` / `margin-bottom`. Text blocks now keep the vertical
  rhythm the deck authored instead of falling back to a fixed line
  height for everything.

- 5381e9d: feat(chart): line / area charts now overlay the per-series
  `<c:trendline>` when authored. Same regression types as the
  column-chart variant (linear / log / exp / movingAvg / poly+power
  fallback). Only emitted on the clustered layout — stacked plots
  already convey the cumulative shape.
- ef4d410: feat: `getSlideMasterBackground(pres, layout)` returns the master's
  `<p:bg>` (both `<p:bgPr>` and `<p:bgRef>` forms). Playground extends
  its background fallback chain to slide → layout → master so brand
  backgrounds authored on the master alone finally render on inheriting
  slides instead of falling through to the theme's `light1`.
- 9c7a852: feat: `getSlideMasterBackgroundGradientFill(pres, layout)` returns
  the master's gradient background when `<p:bg><p:bgPr><a:gradFill>`
  is authored. Completes the three-level bg cascade for gradient
  backgrounds — slides can fall through slide → layout → master.
- 60df186: feat: more name-based finders now accept `RegExp` —
  `findSlideLayout(pres, name)`,
  `findCommentAuthorByName(pres, authorName)`, and
  `findSlidesByLayoutName(pres, layoutName)`. Pairs the RegExp support
  recently added to `findShapeByName` / `findShapesByName` /
  `findCommentsByAuthor` / `findSlideByTitle`. String callers unchanged.
- 10a9d05: feat: `getParagraphPropertiesEffective(pres, shape, p)` — paragraph-property
  cascade mirroring the rPr one. Resolves alignment, left / right / first-line
  indents, line spacing, paragraph spacing (before / after), and rtl through
  the paragraph → text-body lstStyle → layout placeholder lstStyle →
  master placeholder lstStyle → master txStyles chain.

  The playground uses it as the primary source of paragraph properties so
  placeholders inherit their default alignment / line-spacing / indent from
  the layout / master, with any per-slide override winning on top.

  Adds the `ParagraphProperties` type to the public surface.

- 263bf52: feat: `getShapeParagraphElements(shape, paragraphIndex)` returns the
  inline children of a paragraph (runs, field placeholders, and line
  breaks) in document order. Renderers can walk this list to reproduce
  the full visible content — footer / date / slide-number `<a:fld>`
  text was previously dropped by the strict `<a:r>`-only run accessors.

  The playground now uses it: footer text + slide numbers + datetime
  fields show up in the preview, and `<a:br>` line breaks render as
  real `<br/>` inside the foreignObject body.

  Adds the `ShapeParagraphElement` discriminated union to the public
  type surface.

- 263bf52: feat: `getParagraphIndent(shape, p)` returns the paragraph's
  `<a:pPr marL marR indent>` values in EMU (`null` for sides the
  paragraph doesn't author). Playground projects each side to CSS
  `padding-left` / `padding-right` / `text-indent` and skips the
  level-based default when the paragraph carries an explicit `marL`.
- 263bf52: feat: `getShapePatternFill(pres, shape)` returns the pattern preset
  token plus the foreground / background colors resolved against the
  deck's theme. Pairs with the existing `setShapePatternFill`. The
  playground renderer now paints real SVG `<pattern>` tiles for the
  common `ST_PresetPatternVal` tokens (pct5..pct90, light/dark diagonal
  and horizontal/vertical stripes, grids, weave, wave, sphere, diamonds)
  instead of substituting a flat tint.
- 1b08908: feat(chart): per-series `<c:ser><c:dLbls>` overrides. `ChartSeries.dataLabels`
  mirrors the chart-level `ChartSpec.dataLabels`; the series-level
  override wins when present. Playground's bar / column renderers
  check the per-series flag first so one series can show labels while
  others stay clean — common in financial decks.
- 263bf52: feat: playground now applies the picture corrections that already
  shipped on the API: source-rectangle crop (`<a:srcRect>`), brightness
  (`<a:lumOff>`), contrast (`<a:lumMod>`), and opacity (`<a:alphaModFix>`).

  Crops project to an enlarged `<image>` element clipped to the shape's
  bounds (matching PowerPoint's "Crop" tool). Brightness + contrast
  compose into an SVG `<feComponentTransfer>` filter. Opacity drives
  the `opacity` attribute directly.

- 7303fa8: feat(chart): `ChartSpec.firstSliceAngleDeg` reads `<c:firstSliceAng>`
  and `ChartSpec.holeSizePct` reads `<c:holeSize>` for doughnut charts.
  Playground rotates the first slice's starting position clockwise from
  12 o'clock per the authored angle, and sizes the doughnut hole at the
  authored percent (10..90) of the outer radius instead of the
  hard-coded 55%.
- 0ca34e1: feat: pie / doughnut slice explosion. `ChartSeries.pointExplosions`
  exposes the per-data-point pull-out percentage from `<c:dPt><c:explosion val="N"/>`,
  and the playground renderer offsets exploded slices (and their labels)
  outward along the slice mid-angle. Matches the "pulled-out" pie look
  authors get from Excel's "Vary colors by point" toggle.
- 2e9776d: feat(site/playground): chart / media count badges on each slide.
  `getSlideCharts(slide)` and `getSlideMediaPartNames(pres, slide)`
  power two new badges (`N chart`, `N media`) showing how many chart
  shapes and how many media parts (images / audio / video) the slide
  references — useful for quick deck audits.
- 997d507: feat(site/playground): comment badge tooltip carries the comment
  texts. The `N cmt` badge's `title=` attribute now joins each
  comment's body text so hovering surfaces the review remarks
  without opening PowerPoint.
- 3ede588: feat(site/playground): additional slide badges — `hidden` (when
  `show="0"`) and `N cmt` (count of authored review comments). Threads
  `isSlideHidden` and `getSlideComments` through the slide-snapshot
  and surfaces both next to the existing `trans` / `anim` badges, so
  audit views see the full set of slide-level flags at a glance.
- 7a489fa: feat(site/playground): layout-type badge tooltip carries the
  user-visible layout name. Hovering the small `obj` / `title` / etc.
  badge now reveals `layout: <Name> (type: <token>)` from
  `getSlideLayoutName(layout)`. Helps identify which authored layout
  each slide is bound to without leaving the playground.
- d7d8571: feat(site/playground): show the slide's layout type (`title`, `obj`,
  `twoObj`, `blank`, …) as a badge next to the slide title. Reads
  `<p:sldLayout type="…">` via `getSlideLayout` + `getSlideLayoutType`
  so deck audits can spot which layout each slide is bound to without
  opening PowerPoint.
- 5861d1e: feat(site/playground): include slide-master count in the
  "masters · layouts · sections" meta cell. `getPresentationSummary`
  already returned layout / section counts; the playground now also
  calls `getSlideMasterCount` so multi-master decks surface that fact
  in the audit panel.
- 9a64bb4: feat(site/playground): expose `getPresentationSummary` data in the
  meta panel — theme name, layout / section counts, total shape count,
  and deck-wide flags (hidden slides, charts, comments, animations).
  Gives deck audits a one-glance overview without scrolling through
  every slide.
- d9ba44d: feat(site/playground): render section dividers in the slide list.
  Reads `getSlideSections(pres)`, maps each section's first slide to
  the section's name, and renders a dashed divider above that slide
  in the slide list. Deck audits can now see the section grouping at
  a glance.
- 62069f7: feat(site/playground): make the per-slide number an anchor link.
  Each slide's two-digit index in the head row is now an `<a
href="#slide-N">` link, so users can right-click → "Copy link
  address" to share a deep link to a specific slide. Paired with the
  `id="slide-N"` already on each `<li>`, the link also scrolls the
  slide into view when clicked.
- ffec23d: feat(site/playground): show speaker notes under each slide. The
  playground now calls `getSlideNotes` for every slide and renders a
  collapsible `<details>` block when notes exist, so users can
  inspect the deck author's notes without opening PowerPoint.
- b90f1bc: feat(site/playground): show `validatePresentation` results. The
  playground now runs the validator after parsing and surfaces any
  issues in a dedicated panel (with severity tint and the offending
  part name when available). Lets users spot missing rels, dangling
  slide IDs, etc. without dropping into the test harness.
- e078498: feat: `getShapeRunClickAction(shape, p, r)` returns the per-run
  hlinkClick action with the same `ShapeClickAction` discriminated
  union the shape-level `getShapeClickAction` uses. Recognises external
  URLs, slide-jump (`ppaction://hlinksldjump`), and the four
  nav-preset actions (next / prev / first / last slide). Lets callers
  treat per-run hyperlinks symmetrically with shape-level ones.
- 5015413: feat(site/playground): per-run hyperlinks. Runs carrying `<a:hlinkClick>`
  now render in the theme's hyperlink color with an underline, and the
  span is wrapped in an `<a href>` so the preview is clickable. Per-run
  font / color / formatting overrides still apply on top — the link
  styling fills the gaps the run didn't author.
- cc592d2: feat(site/playground): per-run slide-jump click actions render as
  in-page anchors. Mirrors the shape-level slide-jump support shipped
  in the prior batch — `getShapeRunClickAction` resolves to either a
  URL or `#slide-N` anchor, and the run-level `<a href>` wrapper
  respects whether the href is in-page (no `target=_blank`) or
  external.
- 2207ed1: feat: scatter, radar, and bubble charts are now modeled as their own
  `ChartKind`s instead of being folded into `line`. `ChartSeries` gains
  `xValues` (`<c:xVal>`) and `bubbleSizes` (`<c:bubbleSize>`); `ChartSpec`
  gains `scatterStyle`, `radarStyle`, `bubbleScale`, and
  `bubbleSizeRepresents`. Read + render only: the preview draws real
  scatter (two value axes + markers), radar (polar spokes/rings), and
  bubble (area-proportional circles) plots, and the write path now rejects
  these kinds loudly — previously a read-modify-write silently corrupted a
  scatter chart into a line chart.
- 08dc68b: feat(chart): `ChartSeries.lineWidthEmu` and `lineDash` read
  `<c:ser><c:spPr><a:ln>` per-series stroke width and preset dash.
  Playground line / area renderer uses the authored stroke width
  (scaled to px) and projects the preset dash to the same
  `stroke-dasharray` cadence shape strokes use.
- e09e734: feat(chart): per-series marker symbol + size.
  `ChartSeries.markerSymbol` / `markerSizePt` read `<c:ser><c:marker>`
  (`<c:symbol val="…"/>` + `<c:size val="N"/>`). Playground line / area
  renderer emits the matching SVG glyph at each data point — circle /
  square / diamond / triangle / star / x / plus / dash / dot — sized
  per the authored point value. `none` hides the markers.
- 995825f: feat: `setShapeHyperlink` and `setShapeRunHyperlink` now accept an
  optional `tooltip` argument that writes a `tooltip="…"` attribute on the
  emitted `<a:hlinkClick>`. Backwards compatible — calls that omit the new
  arg behave exactly as before.

  fix: `getShapeHyperlinkTooltip` previously only looked at the shape's
  `<p:cNvPr><a:hlinkClick>`, missing the run-level tooltip that
  `setShapeHyperlink` writes. It now scans run-level `<a:rPr>` first
  (mirroring `getShapeHyperlink`'s read path) and falls back to the
  shape-click hyperlink — so the reader / writer pair is consistent.

- 051f4bb: feat: writers for the three stroke attributes that had readers but no
  setters — `setShapeStrokeCap(shape, 'rnd' | 'sq' | 'flat' | null)`,
  `setShapeStrokeJoin(shape, 'round' | 'bevel' | 'miter' | null)`, and
  `setShapeStrokeCompound(shape, 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri' | null)`.

  Cap and compound map to `<a:ln cap=…/>` and `<a:ln cmpd=…/>` attributes;
  join writes one of the `<a:round/>` / `<a:bevel/>` / `<a:miter/>` child
  variants. Passing `null` clears the attribute / removes the child so the
  shape inherits the default. Creates `<a:ln>` if absent.

- 56da3ee: feat: `setShapeTextBodyRotationDeg(shape, rotationDeg | null)` — companion
  writer for the existing `getShapeTextBodyRotationDeg` reader. Sets
  `<a:bodyPr rot="N"/>` (in 60000ths of a degree, per OOXML) so the text
  body can rotate independently of the shape's own `<p:xfrm rot>`. Passing
  `null` or `0` clears the attribute so the shape inherits the default.
- a65c05c: feat: `setShapeTextColumns(shape, { count, gapEmu? } | null)` — multi-
  column writer pairing the existing `getShapeTextColumns` reader. Writes
  `<a:bodyPr numCol="N" [spcCol="EMU"]/>`. Passing `null` clears both
  attributes so the text body falls back to PowerPoint's default single
  column. `count` must be `>= 2` (single column is the default — pass
  `null` instead); the function throws otherwise.
- fea7725: feat: `setShapeTextDirection(shape, direction | null)` — companion
  writer for the existing `getShapeTextDirection` reader. Sets
  `<a:bodyPr vert="…"/>` with any of the six `ST_TextVerticalType`
  values (`vert`, `vert270`, `wordArtVert`, `eaVert`, `mongolianVert`,
  `wordArtVertRtl`); passing `null` or `'horz'` clears the attribute so
  the shape uses the default horizontal direction.
- 4006813: feat: `setTableCellAnchor(cell, 'top' | 'center' | 'bottom' | null)` and
  `setTableCellMargins(cell, {left?, right?, top?, bottom?} | null)` —
  writers for two `<a:tcPr>` properties that already had readers
  (`getTableCellAnchor`, `getTableCellMargins`). The anchor setter maps
  `top`/`center`/`bottom` to the schema's `t`/`ctr`/`b` values and clears
  the attribute on `null`. The margins setter writes per-side EMU on
  `marL`/`marR`/`marT`/`marB`; sides set to `null`/`undefined` are
  stripped (PowerPoint falls back to its defaults); passing the whole
  arg as `null` clears every side. Both create `<a:tcPr>` if absent.
- 3921802: feat: `setTableCellBorders(cell, sides | null)` — partial-update writer
  for all 6 cell-border slots (`left`, `right`, `top`, `bottom` + the
  `tlToBr` / `blToTr` diagonals). Pairs the existing
  `getTableCellBorders` reader. Sides listed with `null` are removed from
  `<a:tcPr>`; sides omitted are left untouched. Passing `null` as the
  whole `sides` arg clears every side. Creates `<a:tcPr>` if absent.

  The diagonals are independent of the four cardinal sides — a
  strikethrough cell can have only `tlToBr`.

- c3e01b3: feat: `setTableCellTextDirection(cell, direction | null)` — vertical-
  text writer for table cells, paired with the existing
  `getTableCellTextDirection` reader. Same six `ST_TextVerticalType`
  values as `setShapeTextDirection`. Passing `null` (or `'horz'`) clears
  the `<a:tcPr vert="…"/>` attribute so the cell uses the default
  horizontal direction. Creates `<a:tcPr>` if absent.
- 328f207: feat: `setTableStyleFlags(table, flags)` — partial-update writer for
  the six `<a:tblPr>` boolean style flags (`firstRow`, `lastRow`,
  `firstCol`, `lastCol`, `bandRow`, `bandCol`). Pairs the existing
  `getTableStyleFlags` reader. Only the keys present in `flags` are
  touched — omitted keys keep their current state. A flag set to `false`
  strips the attribute (matching how PowerPoint round-trips defaults).
  Creates `<a:tblPr>` if absent. Throws when the shape isn't a table
  graphic frame.
- 1ea509b: feat: `setTableStyleId(table, styleId | null)` — writer for
  `<a:tbl><a:tblPr><a:tableStyleId>`. Pairs the existing `getTableStyleId`
  reader. Pass the curly-braced GUID (e.g.
  `'{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'` for PowerPoint's "Medium
  Style 2 - Accent 1") or `null` to remove the reference so the table
  uses the slide's default style. Creates `<a:tblPr>` if absent. Throws
  when the shape isn't a table graphic frame.
- 2438696: feat(site/playground): shape `aria-label` from authored alt text.
  Each rendered shape with a non-empty alt title (or, as fallback,
  alt description) now exposes `role="img" aria-label="…"` on the
  root `<g>`. Screen readers announce decks the same way PowerPoint's
  Accessibility Inspector reports them, without affecting visuals.
- b8e24d6: feat(site/playground): each shape's authored name surfaces as a
  `data-pptx-shape-name` attribute on its root `<g>` element. Lets
  DevTools, a11y inspectors, or test selectors target shapes by their
  PowerPoint name without parsing SVG geometry. Cheap to emit and has
  no visual impact.
- fdd4770: feat: `getShapeTextBodyRotationDeg(shape)` returns the shape's text-body
  rotation from `<a:bodyPr rot="N"/>` (where N is in 60000ths of a
  degree). Distinct from the shape's geometry rotation (`<p:xfrm rot>`):
  this rotates the text body _inside_ the shape without rotating the
  geometry. The playground renderer pivots the text body around the
  inset midpoint when the angle is non-zero, matching PowerPoint's
  behaviour for vertical-label callouts and rotated text frames.
- 263bf52: feat: `getSlideBackgroundGradientFill(slide)` returns the gradient
  stops + path for slides with a `<p:bgPr><a:gradFill>` background.
  Playground paints gradient slide backgrounds via the same projector
  that handles shape gradients (linear / radial / rect / shape).
- 263bf52: feat: `getSlideBackgroundPatternFill(pres, slide)` returns the pattern
  preset + theme-resolved foreground / background for slides whose
  `<p:bgPr>` carries a `<a:pattFill>`. Playground now paints pattern
  slide backgrounds via the same SVG `<pattern>` tile generator that
  handles shape pattern fills.
- c2bcc39: feat: `getSlideBackground` now handles `<p:bgRef>` (the theme-fill-
  reference variant of slide background, e.g. `<p:bgRef idx="1003">
<a:schemeClr val="bg1"/></p:bgRef>`). Returns the inner color element
  as a solid fill so renderers paint the slide background even when
  the deck uses the theme-reference form instead of explicit `<p:bgPr>
<a:solidFill>`.
- 8b7cab6: feat: `slideHasAnimations(slide)` — per-slide animation predicate.
  Returns `true` when the slide carries a `<p:timing>` block (at least
  one authored animation effect). Complements the deck-wide
  `getPresentationSummary().hasAnimations`. The site playground uses
  it (plus `getSlideTransition`) to show small `anim` / `trans`
  badges next to each slide title so deck audits don't need to open
  PowerPoint.
- c0e0dc2: feat(site/playground): shapes with slide-jump click actions
  (`<a:hlinkClick action="ppaction://hlinksldjump"/>`) render as
  in-page hash anchors. The renderer resolves the target via
  `getShapeClickAction` and emits `<a href="#slide-N">`; each slide's
  `<li>` carries `id="slide-N"` so clicks scroll to the target slide.
  Plain URL click actions render the same way as shape-level
  hyperlinks (with `target="_blank"`).
- 7410df9: feat: `getSlideMasterPartName(slide)` returns the part-name of the
  slide master the slide inherits from. Useful for multi-master decks
  where different slides live under different brand templates and the
  caller needs to scope theme / fontScheme / clrMap lookups to the
  correct master.
- 1f536ab: feat: `getShapeStrokeEffective(pres, shape)` walks the layout → master
  placeholder cascade when the shape's own stroke is `'inherit'`. Same
  discriminant types (solid / none / inherit) as `getShapeStroke`; first
  non-inherit layer wins. Playground uses it so placeholder outlines
  authored on the master / layout finally render.
- 263bf52: feat: full stroke read-back surface — `getShapeStrokeCap`,
  `getShapeStrokeJoin`, `getShapeStrokeCompound` plus the existing
  `getShapeStrokeDash` / `getShapeStrokeArrow`. Renderers now have
  enough information to reproduce dashed outlines, rounded vs square
  caps, miter vs bevel joins, and per-end arrow heads.

  The playground composes `stroke-dasharray` from the preset dash
  patterns (cadence multiplied by stroke width as PowerPoint does),
  emits SVG `<marker>` defs for triangle / stealth / diamond / oval
  arrowheads on connectors and shapes, and maps cap / join through.

- f68bd96: feat(site/playground): table cell borders honor `<a:prstDash>`. The
  reader already surfaced the dash token; the renderer now projects it
  to an SVG `stroke-dasharray` (scaled by the border's authored width).
  Applies to every side, the top-left → bottom-right diagonal, and the
  bottom-left → top-right diagonal.
- a037fa7: feat: `getTableCellAnchor(cell)` returns the cell's vertical text
  anchor (`<a:tcPr anchor="t|ctr|b"/>`) as `'top' | 'center' |
'bottom' | null`. Playground projects each onto a CSS
  `justify-content` so cell text sits at the authored vertical
  position instead of always centering.
- e50f1d6: feat(site/playground): table cell text honors authored `<a:tcPr
marL/marR/marT/marB>` insets. The renderer previously hard-coded a
  4-pixel pad on every side; it now converts each EMU-valued margin to
  px (falling back to 4px only when the side isn't authored) so cells
  with custom inner padding line up the way PowerPoint shows them.
- 42cf575: feat: `getTableCellMargins(cell)` returns the cell's `<a:tcPr marL
marR marT marB>` inset margins in EMU. Each side is `null` when the
  cell doesn't author it, so renderers know to fall back to
  PowerPoint's defaults (91440 EMU / 0.1 in horizontal, 45720 EMU /
  0.05 in vertical).
- ba94f5e: feat: `getTableCellParagraphs(cell)` returns a table cell's text as structured
  paragraphs — each carrying its alignment and per-run format (`size`, `bold`,
  `italic`, `color`, `font`, …) — the rich counterpart to `getTableCellText`,
  which only returns the flat visible string.
- f05aa62: feat: `getTableCellTextDirection(cell)` reads `<a:tcPr vert="…"/>` —
  the same token set as `getShapeTextDirection` (`vert`, `vert270`,
  `eaVert`, `mongolianVert`, `wordArtVert`, `wordArtVertRtl`).
  Vertical column headers in tables commonly use `vert270` / `eaVert`
  so the header label reads bottom-to-top alongside its column.
- 263bf52: feat: table span + border read-back.

  - `getTableCellSpan(cell)` returns `{ gridSpan, rowSpan, hMerge, vMerge }`
    so renderers know which cells own a merged region and which are
    absorbed into one.
  - `getTableCellBorders(pres, cell)` returns per-side borders (left,
    right, top, bottom, plus the two diagonals tlToBr / blToTr) with
    theme-resolved colors, widths, and dash style.

  Playground table rendering now honours both: merged cells are skipped
  on their absorbed positions, and per-cell borders render at the
  authored color / width on top of the default thin grid.

- 263bf52: feat: `getTableStyleFlags(table)` returns the `<a:tblPr>` boolean
  toggles — `firstRow` / `lastRow` / `firstCol` / `lastCol` / `bandRow`
  / `bandCol`. Playground projects each onto a theme-derived tint
  (accent1 for header / footer rows, 92%-white-mixed accent for bands)
  when the cell doesn't supply an explicit fill of its own. Header text
  rendered on the accent gets white text instead of the default body
  color, matching PowerPoint's built-in table styles.
- 243e731: feat: `getTableStyleId(table)` returns the GUID string inside
  `<a:tbl><a:tblPr><a:tableStyleId>`. PowerPoint references built-in
  table styles (`{5C22544A-…}` = Medium Style 2 - Accent 1, etc.) and
  theme-local styles by GUID. Returns `null` when the table doesn't
  author one.
- dfae64a: feat: add `getSlideLayoutShapes(pres, layout)` and `getSlideMasterShapes(pres,
layout)` — the non-placeholder decorative shapes (corner bars, divider lines,
  logos, watermark text) on a slide layout and its master, as render-ready
  `SlideShapeData`. Unlike the older flat `getSlideLayoutBackgroundShapes`, these
  include pictures and groups and work with every `getShape*` reader, so a
  picture logo's bytes resolve (against the layout/master's own relationship
  table). For reading/rendering — the handles are bound to the layout/master
  part, not a slide.
- 263bf52: feat: `getShapeTextColumns(shape)` returns `{ count, gapEmu? }` for
  text bodies that author `<a:bodyPr numCol="N" spcCol="EMU"/>`.
  Playground emits `column-count` / `column-gap` on the foreignObject,
  so newspaper-style multi-column placeholders flow correctly.
- 263bf52: feat: extend `TextFormat` with the remaining commonly-authored
  `CT_TextCharacterProperties` (ECMA-376 §17.18.83) attributes:

  - `strike` — `true` / `false` / `'sngStrike'` / `'dblStrike'`
  - `spc` — character spacing in 1/100 pt
  - `kern` — kerning threshold in half-points
  - `baseline` — superscript / subscript offset as a unit fraction
  - `cap` — `'none'` / `'small'` / `'all'`
  - `highlight` — per-run background color

  All round-trip through `setShapeRunFormat` / `getShapeRunFormat` and
  flow through `getShapeRunFormatEffective`'s inheritance cascade. The
  playground renderer honours each of them in the rendered HTML.

- dc98eb1: feat(site/playground): default text body to the theme's font scheme.
  `<a:fontScheme><a:majorFont>` becomes the default face for title /
  ctrTitle placeholders; `<a:minorFont>` covers everything else. The
  existing per-run `<a:rPr typeface>` override still wins. Templates
  that brand-themselves to Aptos / Inter / etc. now render with their
  authored fonts instead of always falling back to Calibri.
- 263bf52: feat: Tier B fidelity batch.

  - `getShapeTextDirection(shape)` returns the `<a:bodyPr vert="…"/>`
    token (`vert`, `vert270`, `wordArtVert`, `eaVert`, `mongolianVert`,
    `wordArtVertRtl`). Playground projects each onto a CSS
    `writing-mode` / `text-orientation` declaration so Asian and
    Mongolian-style vertical text renders without manual transforms.
  - Playground wraps shapes carrying a `<a:hlinkClick>` in an SVG `<a>`
    element so the preview is clickable — matches PowerPoint's
    slide-show behaviour for shape-level hyperlinks.
  - Group shape rendering now applies the group's own `<a:xfrm rot
flipH flipV>` to the whole subtree before the scale + translate
    that maps internal coords onto slide coords.

- 4688af3: feat: chart trendline `<c:forward>` / `<c:backward>` extensions.
  `ChartTrendline.forward` and `backward` carry the N-period
  extrapolation past the last / before the first data point. The
  playground renderer projects the linear fit further along the x-axis
  by `N * step` so projected-future trendlines render the way
  PowerPoint shows them. Moving-average / log / poly trendlines keep
  their data-range output since extrapolation isn't meaningful for
  them.
- 57117a7: feat: chart value-axis tick labels honor `<c:valAx><c:txPr><a:bodyPr
rot="N"/>`. `ChartSpec.valueAxisLabelRotationDeg` returns the rotation
  in degrees (converted from OOXML's 60000ths-of-a-degree). The
  playground renders each value-axis tick label with a
  `transform=rotate()` around its anchor, symmetric to the
  `categoryAxisLabelRotationDeg` we already projected.
- 3e1c8a1: feat: chart value-axis exposes `<c:scaling><c:logBase val="N"/>`.
  `ChartAxisScaling.logBase` carries the authored log base (commonly
  `2`, `10`, or `Math.E`). The reader clamps to PowerPoint's `[2, 1000]`
  range. Callers that round-trip charts now preserve the log-scale
  flag; the playground renderer still draws linear (log-scale
  projection is a follow-up — exposing the field unblocks it).

### Patch Changes

- 499c590: fix: presentation handles now interoperate across the `pptx-kit` and
  `pptx-kit/node` entry points. The two entries ship as separate bundles,
  and the opaque handles (`PresentationData`, `SlideData`, …) were keyed by
  plain `Symbol`s minted per bundle. Loading a deck with
  `loadPresentationFile` (from `pptx-kit/node`) and then reading it with,
  say, `getSlides` (from `pptx-kit`) crashed with
  `Cannot read properties of undefined`. The handle keys now use the global
  symbol registry (`Symbol.for`), so a handle from either entry is readable
  by the other — and by companion packages that bundle their own reader copy.
- 8fc8f12: fix(site): playground stopped rendering after `SlideCommentData` was made
  opaque and `getSlideMediaPartNames` lost its `(pres, slide)` two-arg
  form. The playground was still doing `comment.text` and
  `getSlideMediaPartNames(pres, slide)`, both of which threw at runtime.
  Switched to the public `getCommentText(comment)` accessor and the
  single-arg `getSlideMediaPartNames(slide)` signature.
- 3fb5101: fix: `getShapeRunFormatEffective` / `getParagraphPropertiesEffective` no longer
  inherit the slide master's `bodyStyle` for plain text boxes. A shape without a
  `<p:ph>` is not a placeholder, so its unsized runs now resolve to no inherited
  size (consumers apply the ~18pt text-box default) instead of wrongly picking up
  the master body size (often much larger). Real placeholders — including ones
  whose `<p:ph>` omits a `type` — still inherit as before. This makes effective
  text formatting match what PowerPoint and LibreOffice render for text boxes.
- cfe8b69: fix: placeholder inheritance now applies the OOXML `ctrTitle`↔`title` and
  `subTitle`→`body` type equivalence. A `ctrTitle` (centered title) now inherits
  its layout/master `title` placeholder's `bodyPr` (e.g. `anchor="ctr"`),
  `lstStyle`, and geometry instead of dropping them — fixing
  `getShapeBodyPrEffective`, `getShapeBoundsResolved`,
  `getShapeRunFormatEffective`, and `getParagraphPropertiesEffective` for
  centered titles and subtitles.
- 610ecac: fix(validator): `validatePresentation` now flags duplicate
  `<p:cNvPr id="N">` values inside a single slide's `<p:spTree>` as
  errors. PowerPoint requires every shape's non-visual ID to be unique
  within its slide; duplicates often appear after pasting shapes from
  another slide without re-allocating IDs. The walk recurses into
  `<p:grpSp>` so duplicates nested in groups are also caught.

## 1.0.0

### Major Changes

- f47b78b: **1.0.0** — first stable release. The public API is now frozen under SemVer.

  **What works at 1.0:**

  - **Read** any `.pptx` produced by PowerPoint, Keynote, Google Slides, or
    LibreOffice Impress, and save it back without corruption. Unknown
    extensions are preserved verbatim on round-trip.
  - **Template editing**: token / text replace across slides and speaker
    notes, image swap with geometry preserved, slide CRUD with placeholder
    inheritance from layout / master.
  - **Authoring on top of an existing master**: 180+ preset shapes, custom
    text formatting, tables, embedded charts (column / line / bar / pie /
    doughnut / area) with auto-generated xlsx, solid / gradient / pattern /
    image fills, shadows and glows, rotation / flip / z-order, hyperlinks
    and click actions, notes and comments, slide transitions, simple
    entrance / exit animations.
  - **Diagnostics**: `validatePresentation` returns invariant violations;
    every XML part is validated against the ECMA-376 XSDs in CI.
  - **Bundling**: one ESM build runs in both Node ≥ 20 and modern browsers.
    Tree-shaking is enforced by a CI test — minimal `load → save` bundle
    is < 75 KB unminified, full fn-API bundle is ~120 KB.

  **Deferred to post-1.0** (read pass-through preserved on round-trip):

  - Constructing new themes / masters / layouts from scratch.
  - SmartArt authoring.
  - Complex animation timing-tree authoring.
  - OLE / ActiveX authoring.
  - Document encryption (read + write).

  **Performance (M-series Node 20):** 100-slide synthetic deck saves in
  ~25 ms, loads in ~20 ms. 100 MB templates fit comfortably under the 2 s
  load/save targets.

  **Migration:** if you were on the pre-1.0 class API
  (`Presentation` / `Slide` / `SlideShape` / `SlideLayout`), see the
  preceding changeset for the rename table. There is no class API at 1.0.

- 665c979: **BREAKING**: the class-based API (`Presentation`, `Slide`, `SlideShape`,
  `SlideLayout`) has been removed. Use the free-function API for every
  capability — one canonical path per operation.

  | Was                              | Now                                      |
  | -------------------------------- | ---------------------------------------- |
  | `Presentation.load(bytes)`       | `loadPresentation(bytes)`                |
  | `Presentation.create()`          | `createPresentation()`                   |
  | `pres.save()`                    | `savePresentation(pres)`                 |
  | `pres.slides`                    | `getSlides(pres)`                        |
  | `pres.slideLayouts`              | `getSlideLayouts(pres)`                  |
  | `pres.addSlide({ layout })`      | `addSlide(pres, { layout })`             |
  | `pres.removeSlide(slide)`        | `removeSlide(pres, slide)`               |
  | `pres.moveSlide(slide, i)`       | `moveSlide(pres, slide, i)`              |
  | `pres.duplicateSlide(slide)`     | `duplicateSlide(pres, slide)`            |
  | `pres.replaceTokens(map)`        | `replaceTokensInPresentation(pres, map)` |
  | `slide.shapes`                   | `getSlideShapes(slide)`                  |
  | `slide.findPlaceholder('title')` | `findSlidePlaceholder(slide, 'title')`   |
  | `slide.addTextBox(opts)`         | `addSlideTextBox(slide, opts)`           |
  | `slide.addShape(opts)`           | `addSlideShape(slide, opts)`             |
  | `slide.addImage(bytes, opts)`    | `addSlideImage(slide, bytes, opts)`      |
  | `slide.addTable(opts)`           | `addSlideTable(slide, opts)`             |
  | `slide.addLine(opts)`            | `addSlideLine(slide, opts)`              |
  | `slide.setBackground(color)`     | `setSlideBackground(slide, color)`       |
  | `slide.setTransition(opts)`      | `setSlideTransition(slide, opts)`        |
  | `slide.setNotes(text)`           | `setSlideNotes(slide, text)`             |
  | `slide.layout`                   | `getSlideLayout(slide)`                  |
  | `slide.notes`                    | `getSlideNotes(slide)`                   |
  | `slide.text`                     | `getSlideText(slide)`                    |
  | `shape.text`                     | `getShapeText(shape)`                    |
  | `shape.setText(value)`           | `setShapeText(shape, value)`             |
  | `shape.position`                 | `getShapePosition(shape)`                |
  | `shape.setPosition(x, y)`        | `setShapePosition(shape, x, y)`          |
  | `shape.setFill(color)`           | `setShapeFill(shape, color)`             |
  | `shape.setStroke(opts)`          | `setShapeStroke(shape, opts)`            |
  | `shape.setRotation(deg)`         | `setShapeRotation(shape, deg)`           |
  | `shape.setHyperlink(url)`        | `setShapeHyperlink(shape, url)`          |
  | `layout.name`                    | `getSlideLayoutName(layout)`             |

  Node entry (`pptx-kit/node`) drops the `Presentation` subclass; use
  `loadPresentationFile` / `savePresentationToFile` instead.

  **Why**: every capability used to have two paths through the public API
  — a class method and a free function. The duplication hurt
  discoverability (which one should you use?), made the bundle larger
  (class consumers dragged the whole prototype in), and forced every
  breaking change to land in two places. The free-function API is the
  canonical surface from now on.

### Minor Changes

- b41c502: Comprehensive feature surface for PPTX authoring + editing. This is the
  first release that covers every L1–L4 capability in the foundation
  plan. Highlights:

  **Round-trip + template editing (L1 / L2)**

  - `loadPresentation` / `savePresentation` (`Uint8Array` / `ArrayBuffer` / `Blob`).
  - Node convenience: `loadPresentationFile`, `savePresentationToFile`.
  - Token replace: `replaceTokensInPresentation`, `replaceTokensInSlide`.
  - Free-text replace: `replaceTextInPresentation`, `replaceTextInSlide`.
  - Slide CRUD: `addSlide`, `removeSlide`, `moveSlide`, `duplicateSlide`,
    `getSlideAt`, `getSlideIndex`, `clearSlideShapes`, `sortSlides`.
  - Cross-deck: `importSlide` (with image-media propagation).
  - Cross-slide: `copyShape`.
  - Diagnostics: `validatePresentation`, `getPresentationSummary`,
    `listPackageParts`, `readPackagePart`, `getMediaParts`,
    `setMediaPartBytes`, `compactPackage`.

  **Authoring (L3)**

  - Shapes: `addSlideTextBox`, `addSlideShape` (180+ presets),
    `addSlideLine`, `addSlideTable`, `addSlideImage`, `addSlideChart`.
  - Charts: `bar` / `column` / `line` / `pie` / `doughnut` / `area` with
    embedded xlsx; read/update via `getSlideCharts` / `setChartSpec`.
  - Tables: per-cell access (`getTableCells`, `setTableCellText`,
    `setTableCellFill`, `setTableCellTextFormat`,
    `setTableCellAlignment`); row + column insert/remove.
  - Slide layout swap: `setSlideLayout`, `findSlideLayout`.

  **Text**

  - Per-shape: `setShapeText`, `setShapeBullets`, `setShapeAlignment`,
    `setShapeTextFormat`, `setShapeHyperlink`, `setShapeTextAnchor`,
    `setShapeTextMargins`, `setShapeTextWrap`, `setShapeTextAutoFit`.
  - Per-paragraph: `setParagraphAlignment`, `setParagraphBullet`,
    `setParagraphLevel`, `setParagraphSpacing` + read-back pairs.
  - Per-run: `setShapeRunFormat`, `setShapeRunText`,
    `getShapeRunFormat`, `getShapeParagraphCount`, `getShapeRunCount`,
    `getShapeRunText`.

  **Geometry**

  - Position / size / rotation / flip + combined `setShapeBounds` /
    `getShapeBounds`. Z-order: `bringShapeToFront`, `sendShapeToBack`,
    `bringShapeForward`, `sendShapeBackward`.

  **Fill / stroke / effects**

  - Fill kinds: solid, gradient, pattern, image, none + `getShapeFill`
    read-back.
  - Stroke: color + width + dash + arrowheads + `getShapeStroke` /
    `getShapeStrokeDash` / `getShapeStrokeArrow` read-back.
  - Effects: `setShapeShadow`, `setShapeGlow`, `clearShapeEffects` +
    `getShapeEffect` read-back.

  **Pictures**

  - Crop, opacity, brightness (`lumOff`), contrast (`lumMod`),
    image replacement, image-as-fill. Read-back pairs for every setter.

  **Slide-level (L4)**

  - Notes (`getSlideNotes` / `setSlideNotes`).
  - Transitions (every effect + read-back).
  - Animations (`fadeIn` / `fadeOut` / `appear` / `disappear`) +
    read-back.
  - Comments (legacy schema, author dedup, optional position + date).
  - Backgrounds: solid color or embedded picture; read-back.
  - Visibility: `setSlideHidden` / `isSlideHidden`.
  - Slide sections (p14:sectionLst).
  - Slide size + presets (`SLIDE_SIZE_4_3` / `16_9` / `16_10`).
  - Slide title shortcut (`getSlideTitle` / `setSlideTitle`).
  - Click actions: URL / slide jump / preset nav + read-back.

  **Theme + package**

  - `getPresentationTheme` — color scheme (`accent1`–`accent6`, `dark1`,
    `light1`, `hyperlink`, ...).
  - `getMediaParts`, `listPackageParts`, `readPackagePart` for audit /
    export workflows.

  **Tree-shake**

  - The minimal `load`+`save` import is ~60 KB; the full fn-API
    bundle ~123 KB. CI guard via `test/tree-shake.test.ts`.

  All emitted XML validates against the ECMA-376 strict schemas
  (pml.xsd, dml-chart.xsd, opc-relationships.xsd, opc-contentTypes.xsd)
  via Layer-1 tests.

  **Additional helpers** (all tree-shakeable free functions)

  - Properties: `getCoreProperties` / `setCoreProperties`,
    `getExtendedProperties` / `setExtendedProperties`, plus convenience
    `getPresentationCreated`, `getPresentationModified`,
    `incrementRevision`, `touchModified`.
  - Thumbnail: `getThumbnail` / `setThumbnail` / `removeThumbnail`.
  - Theme: `getPresentationTheme`, `getPresentationFonts`.
  - Slide queries: `getSlideCount`, `getSlideLayoutCount`,
    `getVisibleSlides`, `getHiddenSlides`, `getSlidesWithNotes`,
    `getSlidesWithComments`, `getSlidesWithImages`,
    `getSlidesWithCharts`, `getSlidesWithTables`,
    `getSlidesByLayout`, `findSlideByTitle`, `findSlideByText`,
    `findSlidesByText`, `findSlideByPartName`,
    `findSlideLayoutByType`, `findSlideLayoutByPartName`.
  - Bulk inventories: `getAllNotes`, `getAllComments`, `getAllCharts`,
    `getAllTables`, `getAllImages`, `getPresentationText`,
    `getSlideOutline`.
  - Shape introspection: `getShapeAt`, `getShapeIndex`,
    `getShapeSlide`, `getShapeXmlString`, `getShapeChartKind`,
    `getShapeChartSpec`, `getShapeImageFillBytes`,
    `getShapeImageFormat`, `getShapeImagePartName`,
    `getShapeAltTitle` / `setShapeAltTitle`,
    `getShapeDescription` / `setShapeDescription`.
  - Shape predicates: `isChartShape`, `isTableShape`,
    `isShapeHidden` / `setShapeHidden`, `isShapePlaceholder`,
    `hasShapeImage`, `hasShapeText`.
  - Shape search: `findShapeByText`, `findShapesByText`,
    `findShapesByKind`, `findChartByKind`,
    `findChartsBySeriesName`, `findCommentsByAuthor`,
    `findSlidePlaceholders`, `findSlidePlaceholderByIdx`.
  - Mutation: `setShapeRunHyperlink`, `getShapeRunHyperlink`,
    `getSlideBody`, `appendShapeText`,
    `appendSlideNotes`, `removeSlideNotes`,
    `swapSlides`, `mergePresentations`, `slidesUsingMediaPart`,
    `setTableColumnWidth`, `setTableRowHeight`, `getTableColumnWidths`,
    `getTableRowHeights`, `getTableCellAlignment`, `getTableCellFill`.
  - Diagnostics: `getSlideXmlString`, `getSlidePartName`,
    `getSlideLayoutPartName`, `getSlidesByLayout`.
