# @office-kit/pptx-dsl

## 0.6.0

### Minor Changes

- 4235418: fix: `Text` with `bullets` crashed on `@office-kit/pptx` 0.20.0 with `api.setShapeBullets is not a function`

  0.20.0 renamed `setShapeBullets` to `setShapeBulletStyle`. The DSL was updated in
  the same change but never republished, because its `>=0.17.0` peer range kept the
  new core in range and the release tooling only bumps peer dependents that fall out
  of range. npm and pnpm therefore installed the two together and every deck using
  `bullets` failed at render time.

  The peer range is now `^0.20.0`, so `@office-kit/pptx` 0.17–0.19 are no longer
  accepted. Upgrade the core package alongside this one.

## 0.5.0

### Minor Changes

- 36c4987: feat: `Fill` accepts `autoFit`

  `<Fill target={...} autoFit="normal">` sets text auto-fit on the placeholder it writes into, with the same values as `Text` (`'none'`, `'normal'`, `'shape'`). Text longer than a template placeholder can now be told to shrink without a `Raw` callback that selects the same shape a second time. Omitting `autoFit` keeps whatever the template's `<a:bodyPr>` says, as before.

- 36c4987: feat: rich runs and merges in `Table` cells

  `styleCell` and the cell styles (0.4.0) cover a cell's fill, format, alignment, anchor and borders. This adds the two things a style cannot express.

  - A cell in `rows` is a string or a cell object, `{ text }` or `{ paragraphs }` with optional `colSpan` / `rowSpan`. Strings behave as before and can be mixed with objects; an empty cell is the string `''`. The new exported types are `TableCell` and `TableCellSpec`.
  - `paragraphs` takes core `ParagraphSpec[]`, for bold or colored runs inside one cell. The merged cell style (`cellStyle`, `headerStyle`, `styleCell`) is the base of every run, so a run states only what differs; a `Raw` callback used to lose the table-wide size and color. The style's `align` applies unless a paragraph has its own. `styleCell` receives a rich cell's `value` as its run texts joined, one line per paragraph.
  - A merge is declared on its top-left cell with `colSpan` / `rowSpan`. `rows` stays a full rectangular grid and every covered position is written as `''`; any other content there is an error. Covered positions get no style and no `styleCell` call, because the merged block takes its fill and borders from its top-left cell.

- 36c4987: feat: `bullet` and `level` on single `Text` paragraphs

  `bullets` and `paragraphSpacing` (0.4.0) apply to every paragraph. This adds the per-paragraph half, for nested lists and for a heading line that stays out of the list.

  - Each entry of `paragraphs` accepts `bullet` (`'bullet'`, `'number'`, `'none'`, `{ char }`, `{ autoNum }`) and `level` next to the core `ParagraphSpec` fields. The new exported types are `TextParagraph` and `TextLevel`.
  - A paragraph's `bullet` wins over `bullets`, so `bullet: 'none'` exempts one line.
  - `level` is typed `0` to `8`, so `tsc` rejects a level outside that range. A nested bullet is indented deeper than its parent.
  - A nested list no longer needs a `Raw` callback that addresses paragraphs by index.

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
