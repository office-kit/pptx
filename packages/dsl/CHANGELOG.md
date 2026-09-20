# @office-kit/pptx-dsl

## 0.4.0

### Minor Changes

- 983eb10: feat: `Line` and `Group` elements, bullets and paragraph spacing on `Text`, text alignment on `Shape`, per-cell table styles

  These are the pieces a typical consulting-style slide needed `Raw` for:

  - `<Line x1 y1 x2 y2 color width />` draws a straight line (rules, dividers). It takes end points instead of bounds.
  - `<Group name>` groups the shapes its children create, so a card made of a panel, an accent bar and two text boxes moves and resizes as one object. Groups nest.
  - `Text` gains `bullets` (`'bullet'`, `'number'`, a custom character, …) and `paragraphSpacing={{ before, after }}` in points; both apply to every paragraph.
  - `Shape` gains `align` and `anchor` for its `text`, so a label can sit centred in a circle or an arrow.
  - `Table` cell styles gain `align` and `borders` (per side, width in points), and the new `styleCell={({ row, column, value }) => style}` callback styles one cell at a time: a coloured status column, a bold total row. Styles merge in the order `cellStyle`, `headerStyle`, `stripeFill`, `styleCell`.

  Existing decks compile to the same output.

## 0.3.0

### Minor Changes

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

## 0.2.0

### Minor Changes

- 4fe8fde: Add typed TSX authoring without a React or Vue runtime, including native text,
  shapes, images, tables and charts, template slide reuse and explicit editing,
  and a Raw callback escape hatch. Add local project initialization, PPTX export,
  template inspection and watch preview with error recovery, plus VSCode tasks
  and Claude Code authoring guidance.
