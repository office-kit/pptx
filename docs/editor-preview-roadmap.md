# Editor and development preview

The editor on `feat/pptx-editor` must become the editing surface of the local
`office-pptx dev` preview, with English and Japanese support. A callable API in a
command registry is evidence of dispatch coverage only; it does not demonstrate
that a user can complete the corresponding editing workflow.

## Completion requirements

| Area                    | Required workflow and verification                                                                                                                                                                                          |
| ----------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Development preview     | Open the current deck, edit directly, persist edits, rebuild from source, restart the server, and export the same result without losing edits. Verify conflicts with external source edits.                                 |
| Slides                  | Insert, duplicate, delete, reorder, copy/paste slides; edit background, layout, dimensions, notes and transitions. Verify saved slide order and relationships.                                                              |
| Text                    | Inline editing and Japanese IME; selections with mixed formatting; font, size, color, bold, italic, underline, paragraph alignment, lists, indentation, spacing, links, and find/replace. Preserve unaffected runs on save. |
| Objects                 | Insert text boxes, shapes, lines, images and media; select, move, resize, rotate, multi-select, align, distribute, snap, group, ungroup, reorder, copy/paste and delete. Verify each gesture and exported geometry.         |
| Images                  | Upload, replace, crop, adjust appearance, and author alternative text. Preserve embedded assets through copy/paste and save/load.                                                                                           |
| Tables                  | Edit cells, formatting, borders, row/column insertion and deletion, dimensions, and merge/unmerge. Verify merged text and exported table data.                                                                              |
| Charts                  | Create common chart types, edit categories/series and formatting, and verify both preview and exported chart data.                                                                                                          |
| Presentation            | Theme/layout selection, transitions, animation, notes, comments, hyperlinks, presentation mode, import and export. Verify real document content, not only command availability.                                             |
| History and recovery    | Undo/redo each discrete edit and gesture, rapid edits, edits during asynchronous serialization, new/open, failed saves, and reload recovery.                                                                                |
| Localization and access | Switch Japanese/English live; persist locale; translate menus, dialogs, labels, status and errors; verify keyboard and accessible names in both languages.                                                                  |
| Delivery                | Build the site and published development-tool package, run relevant library and browser tests, inspect rendered UI, update PR description and publish the verified branch changes.                                          |

## Verified so far

- A bilingual paragraph panel targets one paragraph or all paragraphs in a
  selected text shape. It reads inherited formatting, indicates mixed values,
  and edits alignment, bullets/numbering, nine outline levels, proportional or
  fixed line spacing, and before/after spacing. Browser verification covers
  individual and bulk edits, rich-text preservation, saved PPTX properties,
  undo/redo, Japanese labels and reload. Empty text boxes can be formatted before
  typing; first input and added paragraphs retain the chosen settings, including
  through undo/redo. The same controls support individual paragraphs or all paragraphs in a selected
  table-cell range, with bullets and spacing reflected in both preview renderers.
  The inline toolbar also targets the caret paragraph or selected paragraphs for
  alignment, list style/level, line spacing and before/after spacing.

- Slide text replacement now matches across adjacent formatting runs, preserving
  surrounding run formats and assigning replacements the first matched run’s
  formatting. Tests cover Japanese/English, emoji, regular-expression captures,
  replacement tokens, deletion, no-op edits, paragraph/line-break boundaries and
  save/load preservation. A bilingual find/replace dialog navigates matches in
  shapes and table cells, replaces individual matches or all matches, and offers
  case-sensitive and starting-slide-only search with Ctrl/Cmd+F and Ctrl/Cmd+H.
  Exact UTF-16 replacement ranges preserve formatting when adjacent characters
  are identical. Browser verification covers navigation, scope, literal dollar
  signs, Japanese text, cell edits, undo/redo and saved document contents.

- Bilingual speaker-note dialogs load and edit multiline notes for the selected
  slide. Notes readers/writers target the body placeholder, preserve unrelated
  footer placeholders, and read soft breaks and field text. Browser tests verify
  separate notes on two slides, cancellation, undo/redo and reload.
  Transition dialogs configure effects, effect-specific directions, speed,
  through-black fades, click/automatic advance and optional application to all
  slides in one undo step. Browser tests verify push, fade, split, no-effect
  timing, reset, isolation of other slides and persisted settings. Visual
  playback coverage and remaining effects are listed below.

- A bilingual slide options pane selects layouts already in the document, sets
  background colors or uploaded images, and resets to the inherited background.
  Browser tests verify layout relationships, untouched text and neighboring
  slides, embedded image bytes, reset, undo and reload. Bilingual page setup
  supports 16:9, 4:3, 16:10 and custom dimensions in inches or centimeters.
  Browser tests verify live canvas ratios, saved dimensions, unit conversion,
  invalid input, cancellation, undo/redo and unchanged object geometry.
  Page setup also offers proportional content fitting with centered placement.
  A dedicated new-slide layout picker now creates editable placeholders (see below); resetting placeholder geometry is verified below; restoring deleted slots and resetting formatting/remapping remain outstanding.

- Bilingual chart dialogs create column, bar, line, area, pie, doughnut and radar
  charts, edit titles, category labels, series names, colors and numeric data,
  and add or remove categories and series. Browser tests verify all seven kinds,
  blanks/zero/negative values, geometry preservation, cancel, undo/redo, reload
  and actual embedded workbook contents. Legend positions and data label toggles
  are editable in both languages. Chart commands require exactly one selected
  chart and resolve it from the current slide. Other chart kinds and advanced chart
  formatting still need dedicated controls and workflow verification.

- The editor branch incorporates main's development-tool source and current API.
- The generated manifest lists 151 authoring operations. Registry coverage and
  smoke tests verify function discovery and selected API mutations.
- All 151 command labels have English and Japanese versions, enforced by a test.
- Shape text creation and main's bullet-preservation behavior coexist; targeted
  text and registry tests pass.

- The development tool bundles the shared editor, persists committed edits to a
  project sidecar, and uses them for CLI exports. API tests cover restart, stale
  writes and explicit resolution of source conflicts.
- Browser tests cover Japanese/English inline text, movement, autosave, undo,
  reload, locale persistence and source conflict resolution.
- Inline text changes retain unchanged run and paragraph XML, and browser tests
  verify mixed bold/italic formatting in saved PPTX files. Selection formatting
  now has an inline formatting bar for selected characters (bold, italic,
  underline, font, size and color). Selection formatting preserves surrounding
  runs, paragraph properties and complete fields; partial fields become literal
  text. Caret formatting for newly typed text and rich styling within the editing
  overlay still need refinement.
- History tests cover rapid edits, rapid undo/redo, edits during restore,
  new/open invalidation, gesture boundaries and version-aware save completion.

- Clipboard snapshots survive source edits, deletion, undo and opening a new
  document. A browser test verifies cut, paste, undo, redo and persisted output.
- Cross-presentation shape copies preserve images, charts, workbooks and unknown
  dependencies, including cycles. Nested group IDs remain unique after repeated
  copies and subsequent shape creation.

- Slide import now preserves relationship-linked charts, workbooks, notes, media
  and extension parts. Tests cover save/reload, independent repeated imports,
  notes back-references and no destination changes on missing dependencies.
- Slide commands bind the active slide rather than asking for internal objects.
  Navigator buttons, context menus and keyboard navigation cover insertion,
  duplication, deletion and reordering. Browser tests verify Japanese labels,
  undo selection, persisted order and reload. Slide copy/cut/paste now use
  independent snapshots, retain source layouts/themes, charts and notes, and
  insert after the active slide. Browser coverage checks menus, shortcuts,
  Japanese/English labels, undo/redo and pasting after replacing the document.
  The navigator supports Shift ranges, Ctrl/Command toggles, Select All, and
  batch copy/cut/paste, duplicate, delete and reorder (keys, buttons and drag).
  Selection is restored with undo/redo; browser tests cover persisted order and
  both languages. Background color/image/reset, existing layout selection, skipped
  presentation state, and transitions apply to every selected slide in one undo
  step. Mixed values are identified in the property controls; speaker notes
  explicitly target the current slide. Browser coverage verifies mixed state,
  unaffected slides, undo/redo and persisted background images/transitions.

- The Arrange pane lets multiple objects align to either the selection or the
  slide, and always uses the slide for a single object. It distributes equal
  horizontal/vertical gaps. Six-direction slide alignment and one-step undo are
  unit-tested; bilingual browser coverage checks centered/bottom alignment and
  saved geometry. Group/ungroup bind selected
  top-level objects and preserve stacking order and undo selection. Canvas and
  Select All treat nested groups as one object. Browser tests cover bilingual
  controls, moving a group, ungrouping, undo, saved geometry and reload.
  A child inside a group is also resized and rotated from its own handles, with
  the group turned obliquely, scaled differently on each axis, reflected, and
  nested inside a second such group, and with a child that carries a turn of its
  own. Browser coverage drags the handles and checks — from the screen and from
  the saved file, both worked out independently of the editor's own transform
  code — that the dragged corner follows the pointer, the opposite corner and
  the centre of rotation stay put on the slide, the stored angle matches the
  angle swept in the child's own space, and every frame around the child keeps
  its outer _and_ inner transform and child order. Undo/redo, a reload and
  Japanese controls are included; a corner driven past the one opposite it stops
  at the editor's minimum size instead of inverting.

- Image replacement detaches shared media and relationships so copied pictures
  remain independent. Tests cover same-slide and cross-slide copies, same-format
  and cross-format replacement, repeated edits and saved PPTX reloads. Rejected
  crop edits retain the previous crop through a later successful edit and save.
  Bilingual upload and replacement dialogs and image options now cover numeric
  crop, opacity, brightness, contrast and alternative text. Browser tests verify
  insertion, replacement, undo, crop reset and persisted output after reload.
  Preview contrast treats zero as neutral and scales colors around mid-gray.
  The visual crop dialog opens from a picture double-click or image controls,
  supports eight drag handles and selection movement, keyboard fine/coarse
  adjustment, a result preview, reset and cancel, and commits one undo step.
  Browser coverage verifies drag bounds, English/Japanese controls, unchanged
  original image bytes/geometry, and save/load/history. Crop aspect-ratio
  presets and shape masks remain to be assessed.

- A bilingual table pane exposes a cell grid, text, fill, alignment, row heights,
  column widths and row/column insertion and deletion. Browser tests verify mixed
  formatting preservation, saved dimensions/content, undo and Japanese reload.
  Merged-table structural edits are disabled until span-aware mutation is supported.
  Shift-click selects a cell range for merging; the merged cell can be split.
  Merging keeps formatted paragraphs in the anchor cell. Browser tests verify
  merge, split, saved content, undo and reload in Japanese.
  Cell ranges support fill, horizontal/vertical alignment, bold and italic in a
  single undo step, including covered cells in merged regions. Browser tests
  verify saved formatting, unaffected cells, mixed-format undo and both languages.
  The shared text toolbar also applies font families, fractional point sizes,
  color and underline to cell ranges. Browser tests verify mixed values, saved
  formats after reload, and restoring mixed sizes with undo in Japanese.
  Border controls apply color, width and line style to all cell edges or just
  the outside of a selected range, including merged cells. Browser tests verify
  preserved interior borders, reset, undo and reload in both languages.
  Double-clicking cells on the canvas opens an inline editor at the cell's
  bounds, including merged regions and unequal column widths. Cell selection
  remains visible on the canvas and synchronizes with the properties pane.
  A bilingual insertion dialog creates tables with chosen dimensions, header
  rows and alternating row colors, then selects the first cell. Browser tests
  start from an empty slide and verify cancellation, dimensions/style flags,
  undo/redo, editing the inserted table and reload.
  Text edits preserve existing runs; the shared toolbar formats selected text
  through the optional UTF-16 range in setTableCellTextFormat. Browser tests
  cover English/Japanese editing, merged-cell hit areas, cancellation and undo.

- A bilingual animation pane lists a slide's object animations in click order,
  naming the object the way a person does rather than by the handle the file
  uses. Effects can be added, retimed, reordered and removed, with start
  condition, duration, delay and paragraph builds; one this library reads but
  cannot play keeps its place and says why. Play opens the slide over the editor
  from the document in hand, adding no history entry. One player serves the
  preview, the presenter view and the pane, compositing two effects over one
  object by the order they begin, with the document breaking ties
  (SMIL 3.0 §12.4.3, which PresentationML timing is built on, ECMA-376 Part 1
  §19.5).

- Beyond fading and appearing, the authoring API, the reader, the pane and the
  player cover flying in and out through a named edge of the slide, zooming in
  and out about the shape's centre, and one clockwise spin. A preset is matched
  on all three of `presetClass`, `presetID` and `presetSubtype`, and a spin is
  only named when the rotation really is a single full clockwise turn, so an
  imported effect that means something else is listed and saved as it arrived.
  Each motion animates a CSS property of its own — `translate`, `scale`,
  `rotate` — so a fly and a spin over one shape both take effect and the
  renderer's own `transform` is untouched. Browser tests check the edge a fly
  arrives from (including under a rotated group, where the slide's edge and the
  marker's own axes disagree), paragraph builds, overlapping effects playing
  forward and on resume, a spin never revealing a shape no entrance has shown,
  and reduced motion; the saved-file path is played end to end in the preview.
  Real PowerPoint playback is unverified.

## Outstanding work

Complete workflow coverage, remaining UI translations, accessibility and draft
recovery are still being audited. Mixed-format text editing, table and chart
editing, image operations, slide operations and presentation tools require
workflow-level verification before claiming coverage of common Google Slides
editing tasks. Persisted saves survive reload; uncommitted inline drafts do not
currently have recovery storage.

Presentation view now reads the saved per-slide transition timing: automatic
advance runs only during presentation, stage clicks respect `advanceOnClick`, and
explicit keyboard/control navigation remains available. Timers stop on exit and
at the final slide, survive unchanged live refreshes, and support the full OOXML
unsigned millisecond range. The browser playback test verifies these behaviors
against a serialized deck. Visual playback coverage is listed below.

A separate presenter window now displays the current and next slides, saved
speaker notes, an elapsed timer with reset, and previous/next/exit controls.
It follows the editor's English/Japanese locale and synchronizes navigation with
the audience window. Notes render as plain text outside the audience surface.
Browser checks cover multiline Japanese/English notes, literal HTML in notes,
empty notes, saved note updates, last-slide navigation, popup reload, locale
switching, and a closed opener. Visual playback coverage is listed below.

Skipped slides have a bilingual checkbox in slide options and a marked thumbnail
in both editing and preview views. Playback omits skipped slides when starting,
advancing automatically, using buttons/keyboard/stage clicks, and computing the
presenter window's next slide. Editing still includes the entire deck. The
browser test covers leading, middle, trailing, and all-skipped slides; undo/redo;
Japanese controls; and reload of saved visibility. All-skipped decks cannot start
presentation.

- Inline table editing supports Tab / Shift+Tab across visible cells and appends
  a row at the end of an unmerged table. Spreadsheet TSV paste preserves quoted
  tabs, newlines and quotes, expands the table, and commits as one undo step.
  Pasting across merged cells or expanding a merged table requires splitting
  those cells first. Browser coverage checks navigation, Japanese multiline
  paste, saved round trips, undo/redo and rejected merged-cell paste.

- Cell selection uses arrows to navigate merged anchors without moving the table.
  In the table pane, Shift+Arrow extends a range and Delete/Backspace clears the
  selected cells in one undo step. Canvas cell selection clears only cell text;
  Escape selects the whole table before subsequent object operations. Browser
  checks cover table geometry, range history, Japanese controls and reload.

- Table cell margins are editable on each side in points, including mixed ranges,
  per-side defaults and a full reset. The controls preserve other authored sides
  and reject negative input. Browser checks cover exact saved EMU values, range
  updates, untouched cells, undo/redo, reset and Japanese reload.

Table cell ranges now live in document selection, including their active end.
Panel and global Delete share the same merged-cell-aware range resolution;
Shift+Arrow outside the grid extends the range, and undo restores it.

Cell selection supports native copy/cut/paste of plain-text ranges, including
spreadsheet quoting for tabs, newlines and quotes. Cut clears cell text rather
than deleting the table; paste retains destination formatting and expands
unmerged tables when needed. Source formatting is not transferred by this path.

- Table cell context menus keep deletion scoped to cell text, preserve ranges on
  right-click, and offer Select All Cells and Select Table in English and Japanese.
  Ctrl/Cmd+A selects all cells while a cell is selected, including merged tables.

- Context menus support Shift+F10, arrow/Home/End navigation that skips disabled
  items, Enter activation and Escape focus return. Menu keyboard events do not
  move or delete underlying objects or cells. Browser checks cover EN/JA.

Image cropping now offers original, square, landscape, and portrait aspect presets. Resizing keeps the chosen ratio, and Apply fits the image frame around its existing center in the same undo transaction. Browser coverage checks all presets, constrained dragging, cancel, undo/redo, and saved crop/frame persistence in the bilingual editor.

Canvas right-click targets the clicked object or table cell and preserves existing multi-shape/cell ranges when clicked within them. Empty-space menus clear shape selection; right-button gestures cannot move, resize, or rotate content. Native text context menus remain available during inline editing. Browser regression coverage verifies targeting, geometry preservation, bilingual cell actions, undo, and persisted edits.

Shape appearance commands apply fills, outlines, shadows, glow and text formatting to all selected objects in one undo step. Text formatting preserves each object's content and checks that every target supports text before changing any target. Unit coverage verifies selection retention, untouched objects and mixed-selection rejection; browser coverage verifies fill/outline edits, English/Japanese undo controls and saved reloads. Scaling font sizes and outline widths during multi-object resizing remains pending.

Quick fill and outline controls read authored solid colors from the current selection, resolving theme color transforms. They refresh after selection changes, history navigation and saved reloads, and label mixed colors, non-solid paint and inherited states in English/Japanese. Inherited paint is identified as inherited rather than presented as an authored solid color.

The quick appearance panel also exposes No fill and No outline in English/Japanese. Browser checks verify explicit no-paint states on every selected shape, unchanged unselected shapes, undo/redo and persistence after reload.

Quick outline controls support point widths and all eleven preset dash styles, with English/Japanese labels and inherited/mixed readback. Width edits preserve colors and reject invalid input. Browser coverage checks selected-only changes, all preset values in saved PPTX, undo/redo and reload.

Position/size inputs preserve the original precision of untouched fields. Empty, negative-size and out-of-range inputs restore the current value without changing the document. Browser checks cover fractional positions, negative slide offsets, fractional rotation, English/Japanese controls, undo/redo and saved reloads.

Numeric size editing supports a bilingual aspect-ratio lock. Width and height updates scale the paired dimension from the current full-precision bounds, preserve position/rotation, and undo together. The lock is unavailable for zero-size shapes; proportional overflow is rejected before mutation. Browser checks cover both dimensions, unlocking, undo/redo, overflow and reload persistence of the resulting geometry.

While a table cell is selected, canvas clicks select cells, Shift-click extends the range, and dragging selects a rectangle without moving the table. All selected visible cells are highlighted using the same merged-cell-aware range as formatting and clipboard actions. Browser coverage checks nine-cell selection, four-cell dragging, geometry preservation, undo, and merged-cell range highlights.

Canvas moves, resizes, and rotations can be canceled with Escape, pointer interruption, or window blur. Cancellation restores committed geometry and retains redo history. Escape also restores the selection before a marquee or cell-range gesture; small click jitter does not mutate geometry. Browser coverage verifies geometry restoration, interrupted gestures, marquee cancellation, redo, and saved reloads.

Image options include eight non-rectangular preset masks and rectangle reset in EN/JA. Picture presets use native PPTX geometry, and the preview clips the cropped image using shared shape paths. Source bytes, bounds, and crop survive shape changes, undo/redo, and reload. Browser tests cover every offered mask and the existing crop workflow; unit tests verify geometry round trips and reset of preset adjustment guides. Custom image paths remain separate fidelity work.

Image outlines follow the cropped picture’s preset geometry, rotation, and flips without clipping the line at the image boundary. EN/JA controls set color, width in points, solid/dashed/dotted lines, or no outline. Browser coverage verifies saved borders, undo/redo, reload, and preserved source bytes; renderer coverage checks rotated elliptical outlines and removal.

Outline width-only updates preserve existing color XML, including theme references and opacity. Image width and dash controls retain these properties. Round-trip tests cover RGB and theme colors with alpha; browser coverage checks width edits preserve the chosen image border color.

Page setup and the size API enforce the ST_SlideSizeCoordinate limits (1–56 inches / 2.54–142.24 cm). Invalid API requests leave the saved presentation unchanged. Unit tests exercise both axes, non-finite inputs, out-of-range values, and boundary round trips; browser coverage checks the corrected Japanese centimeter validation. Page fitting now scales physical object and text properties while retaining the original aspect ratio. Table previews now use the OOXML default cell margins in both text modes, with proportional text-box coverage during page fitting. Inherited placeholder bounds and margins also have save/reload coverage. Unsized table runs and empty-cell end marks now scale from the editor’s 18pt fallback and survive repeated fits and save/reload; explicit local size inheritance remains intact. Other imported implicit font defaults, shared theme effects on notes, and advanced diagram/3D formatting still need fidelity coverage.

The link command now opens a bilingual dialog with existing URL and tooltip values, mixed-selection feedback, batch application and removal, and a stale-selection guard. Browser coverage verifies cancellation, preserved bold formatting, undo/redo, and reload. The dialog also edits image/object links and offers an internal slide destination picker in Japanese and English. Browser coverage verifies image URL-to-slide conversion, removal, undo/redo/reload, and text slide-to-URL conversion with bold formatting intact. Selected text ranges now support Web links through the inline formatting bar or Ctrl/Cmd+K, with bilingual prefill, cancellation, removal, undo and saved reload coverage. The range API preserves surrounding links and run formatting across paragraph boundaries and rejects invalid UTF-16 ranges before mutation. Selected ranges also support internal slide targets and next/previous/first/last navigation, with JA/EN destination prefill, removal, preserved surrounding formatting, and saved reload coverage. Run-level navigation is exercised in audience and presenter previews, including skipped slides. Link descriptions are now editable for images, other objects, internal slide destinations, navigation presets, and selected text in both languages. Browser coverage verifies image description prefill, editing, clearing, undo/reload, hover text, and navigation from the image to its target slide. Table-cell text selections now share the bilingual link dialog, supporting URLs, internal slides, navigation presets, descriptions, cancellation, removal, undo and saved reload. Browser coverage checks the neighboring cell stays unchanged and an internal cell link navigates in the development preview. The cell API preserves surrounding run formats and rejects invalid ranges or foreign slide targets before mutation. Both text render modes emit cell anchors; pure SVG now also emits clickable anchors and descriptions for text runs.

Internal slide-link targets now resolve against the current slide order even when link readers return reconstructed handles. Regression coverage checks preview SVG destinations, reorder/save/reload, deleted targets, and isolation between presentations. The development viewer now follows internal links in preview and presentation, supports Enter activation, and avoids ordinary click advancement when either internal or external links are activated. Browser coverage also verifies external links open separately and existing timers, skipped slides, and bilingual presenter controls still work. Internal links also work inside the separate presenter window: browser coverage verifies mouse/Enter activation, synchronized audience slides and notes, and rejection of invalid destination indices. Explicit links to skipped slides still need coverage.

Slide comments now have a bilingual list/editor dialog reachable from add, edit, and remove commands. Changes apply together in one undo step; cancellation leaves the slide intact. The text API retains author, creation date, position, and surrounding XML. Browser coverage checks multiple comments, Japanese multiline content, deletion, undo/redo, and saved reloads. Imported legacy comments and author registries now resolve through their relationships, including nonstandard paths; new parts avoid collisions with existing comments on other slides. Round-trip coverage checks editing, author reuse, isolated deletion, and colliding default names. Threaded replies and resolution remain pending.

Adding and deleting legacy comments now retains extension XML on surviving comments, author records, and their list roots. Author counters update in place, and appended entries carry the correct namespace even when imported XML uses another prefix. Round-trip tests cover add/delete/text edits and preservation of both comment and author extensions.

The bilingual link dialog now offers Next, Previous, First, and Last slide navigation presets. Preview SVG resolves shape and text-run presets to slide anchors, skips hidden slides, and stays on the current slide at a boundary. Renderer tests cover all four presets and hidden slides; browser tests cover saved preset editing, preserved bold formatting, undo, reload, keyboard activation, and presenter current/next preview links. Deck changes now retain surviving slide and shape handles. A regression test edits retained references after add, duplicate, move, reverse, remove, and import, then saves and reloads to verify Japanese text and hidden-slide settings; duplicate content remains independent.

Link mutations preserve the DrawingML child order when imported run or object properties contain hover actions, RTL settings, or extension lists. Regression coverage verifies replacement and removal through saved reloads for whole-text, selected-range, per-run, and object links, retaining the surrounding XML.

Preview links now wrap the rendered object across drawing kinds, including picture render paths. Image link descriptions are escaped in SVG and retained after saved reload. External links authored on a text range remain on that range, rather than extending the first linked run to the surrounding object; renderer regression coverage checks that only the selected text has an anchor.

Table-cell selection links now use the bilingual link dialog from both the cell formatting panel and Insert ribbon. Whole-cell and rectangular cell selections share URL, slide and preset destinations plus descriptions; mixed links are shown before replacement, and the change is one undoable transaction. Browser coverage verifies panel entry without an explicit cell selection, mixed-link batch replacement, untouched neighboring cells, undo/redo, removal and persisted reload.

Deleting a slide now clears incoming object, text and table-cell links and their relationships, preventing links from silently targeting a later slide that reuses the deleted part name. Other URLs and formatting are preserved, including through retained handles. The editor restores the deleted destinations with Undo and removes them again with Redo/save/reload; foreign presentation handles are rejected before mutation.

The bilingual chart dialog now exposes category/value axis titles, value-axis minimum/maximum, major/minor tick intervals and number formats for column, bar, line and area charts. Blank numeric fields restore automatic scaling; invalid bounds, nonpositive intervals and nonpositive bounds on imported logarithmic axes prevent submission. Existing axis settings are preserved unless edited. Browser coverage checks Japanese authoring, rendered titles, English reload/reset, validation and undo/redo.

Chart axis settings also expose category/value axis visibility and value-axis major gridlines in Japanese and English. These controls preserve scaling and other imported axis settings, support cancel and a single undoable update, and persist through save/reload. Browser checks cover visible SVG gridlines and their removal alongside the saved chart specification.

Chart preview tick generation is bounded for extremely small authored intervals and large floating-point values. Intervals requiring more than 1,000 ticks use automatic preview ticks; exported axis settings remain unchanged. Regression tests cover dense, subnormal, fractional, and large-offset intervals.

Chart axis controls now include bilingual value-axis minor-gridline visibility. Cartesian previews (including combo charts) render authored minor spacing, color and width, or automatic spacing when unspecified or too dense. Major tick positions are excluded from minor gridlines. Browser checks cover insertion, editing, save/reload and history.

Chart editing now exposes normal, stacked and 100% stacked series modes in Japanese and English for column, bar, line and area charts. Explicit mode changes reset bar overlap to the appropriate default while preserving chart data and other formatting. Browser coverage exercises all 12 type/mode combinations, percentage rendering, cancellation, undo/redo and reload.

New chart series show the same default palette color in the dialog that the ChartML writer saves. Browser coverage checks the second-series swatch in Japanese, the saved color, and the English editing dialog.

Line chart series now expose marker shapes, marker sizes (2–72 whole points or automatic), and smoothing in Japanese and English. Explicit markers also render on stacked and percentage-stacked lines at their cumulative positions. Browser coverage checks saved settings, rendered markers/curves, validation, cancellation, automatic-size reset, undo/redo and reload.

Smoothed line charts preserve gaps at blank values: curves are calculated separately for each uninterrupted segment. Rendering tests cover gap/span/zero handling, consecutive and edge blanks, and isolated or two-point segments.

Line and area charts expose blank-value handling (gaps, zero, or connected points) in Japanese and English. Area fills close each uninterrupted segment separately, preserving gaps. Browser checks cover all six type/mode combinations, unchanged source values, cancellation, save/reload and undo/redo.

Pie and doughnut dialogs expose first-slice rotation (0–360°); doughnut charts also expose hole size (10–90%). Both settings support whole-number validation and automatic reset in Japanese and English. Browser checks cover saved geometry, cancellation, reset, undo/redo, reload and switching to pie. Imported 3-D and pie-of-pie charts retain their existing angle behavior.

Pie and doughnut charts support per-slice colors, a one-click color palette, and resetting individual slices to the series color in Japanese and English. Color overrides follow their categories when rows are removed and survive undo, redo and reload.

Data label formatting exposes chart-appropriate placement and numeric format codes in Japanese and English. Blank settings restore automatic defaults; edits update existing series and point overrides without changing unrelated label properties. Browser coverage verifies formatted text, placement, cancellation, undo/redo and reload.

Pie and doughnut previews render series-name labels and honor series/point label overrides for visibility, literal text, placement, numeric formatting and text style. Regression tests cover inheritance and XML escaping; the browser verifies series-name-only labels.

Column, bar, line and area previews now honor category/series label toggles and per-point label visibility, literal text, number format and text style, including stacked charts. Regression coverage checks all four chart types across standard, stacked and percent-stacked grouping; browser checks cover series-name-only labels while switching chart types.

Stacked column and bar previews honor center, inside-end, outside-end and inside-base label placement relative to each segment. Regression tests cover positive and negative segments and percentage stacks; browser checks verify that changing placement moves labels in both orientations and stacking modes.

Percentage-stacked columns and bars display original data values when value labels are enabled, consistently with lines and areas. Chart, series and point numeric formats now apply to those labels; percentages on the value axis remain independent. Unit coverage checks format inheritance and unformatted values, and browser coverage checks formatted values while switching among all four chart types.

Image options include a bilingual reset for opacity, brightness and contrast in one undo step. The action is disabled when adjustments are already neutral. Browser coverage checks reset, undo/redo and reload in both languages while preserving image bytes, crop, shape and bounds.

The inline text toolbar exposes paragraph alignment and list style for the caret paragraph or all paragraphs touched by a text selection, including table-cell text. Pending text edits are committed with formatting in a single undo step, preserving rich runs. Selection boundaries exclude a paragraph whose start equals the selection end. Browser coverage verifies English/Japanese editing, added paragraphs, range boundaries, table-cell isolation, history and reload.

The inline paragraph toolbar also exposes all nine list levels in Japanese and English. It reports mixed levels across a selection and changes only touched paragraphs, including within table cells. Browser coverage checks boundary levels, selection isolation, undo/redo and persisted levels after reload.

Inline paragraph spacing supports inherited, proportional and fixed line spacing plus before/after points. Mixed selections remain explicit; blank paragraph spacing restores inheritance and zero remains an explicit value. Browser coverage checks caret/range boundaries, table-cell isolation, rich-run preservation, undo/redo, Japanese controls and reload.

Selected inline text supports Ctrl/Cmd+B, I and U to toggle bold, italic and underline in shapes and table cells. Toggle decisions read the selected runs after replaying pending text edits, preserving selection offsets after inserted paragraphs. Browser coverage verifies both modifier keys, Japanese/English labels, surrounding formatting, undo/redo and reload. Caret-only typing-format controls are implemented; browser coverage in `caret-formatting.test.mjs` verifies surrounding-text isolation, Japanese input, clear formatting, history and reload.

Inline toolbar bold/italic/underline toggles share the keyboard formatting path, so pending text edits are applied before deciding whether to enable or disable a style. A browser regression reproduces the previously incorrect bold toggle after inserting a leading paragraph and checks toolbar/keyboard parity, history and Japanese reload.

Inline character-format controls read a disposable projection of pending edits using the same formatting-preserving mutation path as commit. Inserting paragraphs no longer shifts bold/italic/underline, font or color indicators onto unrelated saved text. Shape/table unit coverage verifies source isolation and sequential inheritance; browser coverage checks the bold indicator before applying or saving an edit. Pending paragraph-property indicators still need equivalent inheritance-aware projection.

The inline and selected-cell character-format bars now include strikethrough, superscript and subscript in English/Japanese. Imported strike styles and positive/negative baseline offsets are recognized; repeated clicks restore normal text, and switching scripts replaces the previous baseline. The shared toggle path respects pending edits and mixed selections. Browser coverage checks shape/cell isolation, switching and reset, Japanese controls, history and reload.

Inline/selected-cell text formatting includes highlight color and removal in English/Japanese. SVG text rendering now paints per-run backgrounds using measured font metrics and the same horizontal anchors as text, including wrapped lines and raised/lowered text. Renderer tests cover alignment, wrapping, plain text, shape/table colors and removal in both rendering modes; browser coverage verifies color persistence and Japanese removal.

Visual QA also exposed that SVG script runs were measured at their original font size while drawn smaller. Width measurement now uses the rendered superscript/subscript size, keeping centered/right-aligned backgrounds and subsequent runs aligned; deterministic tests cover both script directions.

Pending paragraph controls now follow the projected text after insertions and deletions, including caret-only selections. The disposable text model materializes inherited alignment, bullets, line spacing and paragraph spacing before replaying edits; this preserves placeholder defaults without changing the source presentation. Browser coverage checks English insertion and Japanese deletion before saving, alongside existing paragraph formatting, Undo/Redo and reload checks.

List-level option values now use the same string type as the current paragraph value, so the selected level is displayed instead of becoming blank. Table-cell coverage also checks pending insertion before a formatted paragraph, including list style, level and line spacing.

Clear text formatting is available in the inline toolbar and selected-cell toolbar in English and Japanese, with Ctrl/Command+Backslash for inline selections. It resets direct font, color, highlight, decorations and other run appearance while preserving text, hyperlinks, language metadata and paragraph formatting. The existing shape/cell text-format setters accept `reset: true`; whole-body resets also remove run-format defaults, and range resets leave surrounding characters intact. Tests cover round-trip preservation, table ranges and whole cells, pending input, keyboard/button paths and Undo/Redo.

Canvas resize handles now operate in the selected object's rotated axes, keeping
the opposite corner or edge midpoint fixed, including when shrinking to the
minimum size. Shift preserves the starting aspect ratio for corner and edge
handles. Unit coverage checks all eight handles across five rotations, aspect
constraints, minimum size and zero-height lines. Browser coverage verifies
English/Japanese labels, pointer movement on a rotated shape, saved geometry,
Undo/Redo, reload and cancellation without losing redo history.

Multiple selected objects now expose a shared rotation handle. Its centre comes
from the visible bounds, including each object's existing rotation. Dragging
rotates object centres and angles together, preserving sizes, distances and
relative angles; Shift snaps the selection's turn to 15-degree increments. The
handle and hints are bilingual. Browser coverage verifies saved geometry,
unselected objects, Undo/Redo, Escape cancellation and reload.

Multiple selections now also have four bilingual corner resize handles. They scale
positions and dimensions uniformly about the opposite corner, preserve rotations,
and handle zero-width/height lines without introducing thickness. Unit coverage
checks every corner with rotated objects and minimum-size limits; browser coverage
checks expansion, contraction, fixed anchors, unselected objects, English/Japanese,
Undo/Redo, Escape and saved reloads. Scaling font sizes, borders and other appearance
attributes alongside geometry remains pending.

Ungrouping now transfers a group's rotation and horizontal/vertical reflections to
its children, preserving their transformed centres and composing their existing
angles and flips. Unit coverage checks uniform resizing, all flip combinations,
nested groups and saved reloads. Browser coverage checks rotating a group before
ungrouping in English/Japanese, Undo/Redo and reload. Nonuniformly scaled groups
with rotated children still need affine-transform fidelity work.

Imported shape/group flip flags now accept both XML boolean spellings (`true`/`false`
and `1`/`0`), including surrounding whitespace. Round-trip tests compare 36 flag
combinations against numeric equivalents, including preview SVG output and the
positions, rotations and reflections produced by ungrouping.

### Group text reflection follow-up

- Preview shape text now cancels inherited reflection in its local axes, preserving the transformed center and baseline through rotated, nested groups.
- Regression coverage compares text-box corners before and after ungrouping for all 16 combinations of two ancestor groups' horizontal/vertical flips, after a PPTX save/reload, using Japanese and English text.
- This covers shape text overlays. Table/chart internals and anisotropic group scaling with rotated descendants still need separate fidelity verification.

### Table text reflection follow-up

- Table cell text now cancels reflections from the table and enclosing groups around each cell center; cell fills, borders, and transformed cell placement remain governed by the table transform.
- Both foreignObject and pure-SVG text paths are covered by 32 rotation/flip cases with Japanese and English text, after PPTX save/reload. Tests require positive glyph orientation and matching text coordinates after ungrouping.
- Chart label reflections and anisotropic group scaling remain separate fidelity checks.

### Chart label reflection follow-up

- Chart labels now share reflection compensation, including existing label rotations and start/end anchoring. The plot geometry continues to follow the chart and ancestor group transforms.
- Regression coverage exercises all chart/group flip combinations across column, bar, line, area, pie, doughnut, scatter, bubble, and radar charts (144 cases), with Japanese/English labels and save/reload. Every emitted label must retain positive glyph orientation and match its coordinates after ungrouping; reflected legends retain their side of the anchor.
- Nonuniform scaling of rotated descendants still requires fidelity work; these cases use unscaled groups.

### Empty paragraph caret formatting

- Inline character-format indicators and toggle decisions now read the end mark of an empty paragraph in text shapes and table cells. This retains authored bold and size settings before typing, rather than reporting an unformatted caret.
- Regression coverage verifies empty first/last paragraphs, neighboring nonempty text, save/reload and pending insertions that move the caret's paragraph.

### Inherited character formatting in the inline toolbar

- Caret and range controls resolve text-box character formats through the same layout/master/theme cascade as the rendered text, including detached pending edits. Toggle decisions therefore honor inherited bold rather than treating an absent run property as false.
- Unit coverage checks master font size, paragraph-end bold, pending Japanese paragraph insertion and unchanged source XML. Browser coverage checks inherited title size in the toolbar before typing, during pending input and after saving/reopening.

### Imported layout relationship targets

- Switching layouts now retains the selected layout's actual package path, including nested and root-level imported parts and matching filenames in different directories.
- Regression coverage verifies adding/replacing the relationship, stable relationship IDs, unchanged slide content and unrelated relationships, save/reload, and rejection of missing layout parts without mutation. This repairs the existing layout picker; resetting placeholder geometry is verified below; restoring deleted slots and resetting formatting/remapping remain outstanding.

### New slides from document layouts

- The navigator and Home ribbon open a dedicated English/Japanese layout picker. It inserts a slide immediately after the active slide and creates the chosen layout's editable placeholders. Cancel leaves the deck unchanged.
- Browser verification covers both entry points, slide order/selection, placeholder editing with Japanese and English, undo/redo, saved layout relationships and reload persistence. Existing slide content remains intact.

### Restore placeholder geometry

- A bilingual action in the Home ribbon and slide options restores top-level placeholder position, size, rotation and flips from the current layout, with master fallback. It applies to all selected slides in one undoable transaction.
- Unit tests cover same-type slots with different indices, master inheritance with unrelated indices, rotation/flips, retained text/formatting/hyperlinks, unaffected ordinary shapes and save/load. Unmatched and grouped placeholders are preserved.
- Browser verification covers multiple-slide selection, undo/redo, English/Japanese entry points, saved geometry and reload persistence. This action does not recreate deleted slots or reset text formatting; those wider layout workflows remain outstanding.

### New-slide placeholder metadata

- New slides preserve the chosen layout's placeholder orientation and size category alongside type/index. Layout prompt text remains excluded, and geometry/formatting continue to inherit.
- Regression tests cover all three size categories, vertical orientation, omitted defaults, source isolation and Japanese/English text through public add/edit/save/load operations. This verifies placeholder metadata preservation, not complete vertical-text rendering.

### Visual transition playback

- Presentation mode plays cut through black, fade (including through black), push, wipe, cover, uncover and zoom, plus split, circle, diamond, plus, blinds, comb, checkerboard, strips, random bars, dissolve, wedge, newsflash and wheel. Split supports horizontal/vertical and inward/outward variants. Normal preview remains immediate.
- Browser tests load saved decks and verify intermediate animation frames, interrupted-layer cleanup, completion, reduced-motion behavior and automatic advancement after animation. The shape-transition test includes Japanese/English text and visual inspection of the midpoint plus mask.
- Cut through black holds a black frame for the selected transition duration, then switches without fading. Plain cut remains immediate. Imported `true`/`false` and `1`/`0` XML booleans (including surrounding whitespace) preserve through-black and click-advance settings across save/reload.
- Random chooses a supported visual effect on every visit while preserving the saved random setting and transition speed. Saved-deck browser tests verify different selections, repeated visits, interruption, completion and reduced motion with Japanese/English content. Object-animation playback remains a separate requirement.

- Blinds and comb honor horizontal/vertical direction (horizontal by default), transition speed and reduced motion. Saved-deck browser tests check visible/hidden regions in every band at the midpoint, interrupt consecutive effects and verify final cleanup with Japanese/English text. One clipped incoming layer avoids duplicating a full slide per band.

- Checkerboard playback reveals alternating cells in two stages and honors horizontal/vertical direction. Saved-deck browser tests sample all 48 cells at five animation times, including empty/full endpoints, and check interruption, reduced motion and completion. Midpoint screenshots verify the checker pattern with Japanese/English content.

- Strips playback staggers eight bands toward the selected corner, including the default left-up direction. Saved-deck tests verify all four directions plus the omitted default, start/midpoint/end visibility, interrupted cleanup, reduced motion and Japanese/English content.

- Random-bar playback reveals shuffled horizontal/vertical bands using a single incoming layer. Saved-deck browser tests cover default and explicit directions, endpoints, midpoint geometry, interruption, completion and reduced motion with Japanese/English text.

- Dissolve playback adds randomly ordered square cells with a mask sized to the slide aspect ratio. Saved-deck browser tests verify all three speeds, increasing coverage at five points, endpoints, interruption, completion and reduced motion with Japanese/English content.

- Wedge playback opens two radial boundaries from top to bottom; newsflash spins the incoming slide counter-clockwise while enlarging it. Saved-deck tests verify endpoints, midpoint geometry, interruption, completion and reduced motion with Japanese/English content.

- Wheel spokes can be edited in Japanese/English, undone, and saved/reloaded without losing the count. Presentation mode reveals clockwise sectors (default four), using bounded-size gradient masks even for large unsigned counts; zero spokes uses a fade. Saved-deck pixel tests cover one, two, three, four and eight spokes, plus interruption, completion and reduced motion.

Numeric position and size controls resolve placeholder geometry inherited from layouts or masters. Editing one coordinate preserves the remaining resolved bounds; aspect-ratio locking also works before local geometry exists. Browser regression coverage checks English/Japanese edits, undo back to inheritance and saved reloads.

Shape properties expose horizontal and vertical flip checkboxes in English and Japanese, including indeterminate values for mixed selections. Flip commands apply the selected axis to each selected object without changing the other axis or unselected objects. Model and browser tests cover mixed-state history, bulk flags and saved reloads.

The numeric rotation field displays mixed values for multi-selection and applies an entered angle to each selected object about its own center. Model and bilingual browser tests verify negative/fractional angle normalization, empty input, unchanged bounds and unselected objects, undo to mixed angles, redo and saved reloads. Selected shape lookup builds one ID map rather than repeatedly scanning the slide.

### Remaining animation presets

- Only nine presets are authored and played: appear, fade, fly, zoom in each
  direction, and one clockwise spin. The rest of PowerPoint's gallery — the
  other emphasis effects, the diagonal flies, the zoom sub-variants, and
  `presetClass="path"` motion paths — is read and saved unchanged but reported
  as something this library does not play, and a slide carrying one keeps its
  place in the click order without being approximated.
- Interactive sequences, `<p:iterate>` letter/word staggering and repeat counts
  are likewise read rather than played.
- Playback is verified in Chromium against the preview's own renderer. What
  PowerPoint itself does with the trees this library writes is unverified.

### Multiple-selection numeric position and size

The property panel shows common position and size values or localized mixed-value placeholders. Entered values apply to each selected object, as explained beside the controls. Aspect-ratio locking preserves each object's own ratio; paired dimensions are validated for the entire selection before mutation. Browser coverage verifies English and Japanese edits, unchanged unselected objects, proportional-overflow rejection, one-step undo/redo, and unlocked width changes preserved through saved reloads.

### Direct object text formatting

Selecting objects containing text now shows a property-panel formatting bar with font, size, emphasis, text/highlight colors and clear-format controls. It reads effective formats across the selected text and applies changes to every selected object in one undoable transaction. Bilingual scope labels distinguish whole-object edits from character-range editing. Browser coverage verifies mixed bold state, bulk size and Japanese italic edits, clear-format undo, preservation of text and unselected objects, saved reloads, and existing caret-formatting behavior. Mixed text/non-text selections do not show this shortcut. Empty-shape support is covered below.

### Text formatting before typing

Blank autoshapes now expose the same formatting controls, including shapes with no text body. Setting a format creates the body and stores paragraph-end character properties for future typing; reading the paragraph count returns zero before creation. Invalid format and empty range requests leave a missing body untouched. Core tests check round trips and subsequent bilingual text insertion; the browser checks blank-shape formatting, saved reload, toolbar readback and later text entry. Mixed selection preflight still rejects actual non-text objects before applying any change.

### Selected-object text alignment

The quick property panel includes paragraph and vertical alignment controls in English and Japanese. Horizontal state is read across every selected paragraph, with mixed or inherited values shown explicitly. Changes apply to every selected text shape in one undo step, including blank autoshapes without a text body. Non-text targets are checked before any mutation. Core, model and browser coverage checks invalid input, paragraph-wide application, blank-shape persistence, mixed readback, undo/redo, bilingual controls, saved reload and preservation of unselected objects.

### Text content selection scope

The quick text content editor is available only when one text-capable shape is selected. Multiple text shapes keep bulk formatting controls and show a localized instruction to select one shape for content editing. Non-text objects no longer expose an unusable text area. Browser regression coverage checks selection changes, bilingual guidance, multiline content replacement on the intended shape, undo/redo, saved reload and line selection.

### Comments across slides

The comment dialog offers a bilingual slide selector with slide titles and draft comment counts. Each slide retains its own drafts as reviewers navigate. Apply saves all slides in one undoable transaction; Cancel discards the entire draft. Validation includes unfinished comments on other slides, while untouched new rows are ignored and unchanged dialogs cannot create no-op edits. Browser coverage checks bilingual cross-slide add/edit/delete, draft retention, validation outside the visible slide, blank rows, cancel, one-step undo/redo and saved reload. Reply threads and resolution remain separate outstanding work.

### Comment replies

- The comments dialog can reply to existing comments and unsaved drafts, including replies to replies. Parent author and text are visible in English and Japanese.
- Apply commits all drafts in one undo step; Cancel discards them. Deleting a parent removes its descendants and is labeled “Delete thread”.
- Replies persist using the PowerPoint p15 threading extension in legacy comment parts. Core and browser tests cover save/reload, editing, alternate XML prefixes, unknown extension preservation, sibling retention, and undo/redo.
- Modern p188 comment parts and resolved status are covered below.

### Modern comments and resolved threads

- Comments written by PowerPoint 2021 and Microsoft 365 are read and edited rather than only carried through. The format is Microsoft's own, not ECMA-376: [MS-PPTX] §2.16.1 for the Author and Comment parts and §5.14 for the schema, both read from the published specification rather than inferred from a file.
- A thread owns its replies, its text is a DrawingML body, authors are GUIDs in `/ppt/authors.xml`, and a thread carries `status` — `active`, `resolved` or `closed`. `getSlideComments` returns both formats; `getCommentFormat`, `getCommentStatus` and `setCommentStatus` are public. A legacy comment reports no status, and setting one throws instead of silently doing nothing.
- Editing goes through the file's own tree, so the slide or shape anchor, the pin, extension lists and reactions all survive a text edit, a resolve and a save. Core tests cover reading threads/replies/authors, resolve and reopen through save and reload, preserved unknown XML, replying inside a thread, starting a thread with the anchor the slide's own threads use, removing a reply or a whole thread, dropping the part and its relationship with the last thread, and a deck that carries both formats at once.
- The comments dialog hides resolved threads with a count, brings them back on request, and resolves or reopens them in English and Japanese. A status change on its own enables Apply — the dialog used to compare only counts and text. Browser coverage drives resolve, Cancel discarding it, Apply, one-step undo/redo, reopening a resolved thread in Japanese, and one Apply carrying edits made on two slides.
- Fixed on the way: `<p:pos>` is required by `CT_Comment` and nothing made a caller pass one, so every comment the editor added was schema-invalid. New comments are pinned to the slide origin, with a schema test that would have caught it.
- Not covered: how real PowerPoint renders any of this is unverified, as is every other visual claim here. Assignment (`assignedTo`, `dueDate`, `complete`), reactions and shape-anchored authoring are read and preserved but not editable.

### Comment conversation navigation

- Replies appear immediately after their parent thread even when the PPTX stores them after unrelated comments. Display ordering does not change save ordering or reply identity.
- Adding a comment or reply focuses its author input. Deletion moves focus to a surviving comment or the Add comment button.
- Thread ordering is iterative and linear; tests cover orphaned parents, cycles, and 20,000 nested replies without recursive stack growth. Browser coverage verifies interleaved threads across save/reload and keyboard focus after additions and deletions.

### Restore deleted placeholders

- Restore missing title and content slots from the current layout through the slide panel or Home ribbon in English and Japanese. Selected slides are restored in one undoable transaction.
- Existing text, geometry and formatting stay intact. Restored slots start empty and inherit layout styles and geometry; grouped placeholders are recognized and repeated restoration does not duplicate slots.
- Core tests cover unique IDs, empty and blank slides, grouped placeholders and editable saved content. Browser tests verify multiple selected slides, undo/redo, both languages and saved reload. The combined layout reset below also clears direct appearance overrides.

### Placeholder text formatting reset

- The slide panel and Home ribbon can reset placeholder text formatting to inherited layout defaults in English and Japanese, for every selected slide in one undo step.
- Direct run styling, paragraph spacing/alignment/bullets and text-body alignment/margins/autofit are cleared. Text, fields, links, language, paragraph outline levels and unknown extensions survive. Shape geometry and non-placeholder content are preserved; grouped placeholders are left unchanged.
- Core tests cover preserved content and metadata, idempotence, save/reload and grouped objects. Browser tests check selected-slide scope, bilingual controls, undo/redo, saved reload and retained geometry. The combined layout reset below includes shape appearance; grouped placeholders are covered below.

### Readable layout actions

- Home separates layout restoration/reset actions into their own group with distinct compact English and Japanese labels, full accessible names and tooltips.
- Ribbon groups retain their width and scroll horizontally instead of squeezing command labels as more actions are added. The ribbon can grow vertically to keep group titles below two-line labels visible. Existing bilingual layout-picker and reset browser workflows verify dispatch, history and persistence after the presentation change.

### Combined layout reset

- Reset layout restores missing slots, layout geometry and inherited shape/text appearance in one command, available in English and Japanese in the slide panel and Home ribbon. All selected slides share one undo step.
- Text, hyperlinks, image content and crops, relationship targets, shape IDs and unknown extension metadata survive. Non-placeholder objects remain unchanged.

### Resetting grouped placeholders

- Text formatting and shape appearance are restored for a placeholder inside a group, from both `resetSlidePlaceholderTextFormatting` and `resetSlideLayout`, and both count it. A group scales and turns what is inside it; it does not decide what font the text is in.
- Geometry is deliberately left alone, from `resetSlidePlaceholderGeometry` and from the geometry half of `resetSlideLayout`. The layout states a rectangle on the slide while a grouped shape's geometry is written in its group's coordinate space, so restoring it would either tear the shape out of the arrangement it was grouped into or invent a rectangle the layout never described. The reason is recorded on both functions.
- Core tests check that a grouped title and body lose direct formatting, fills and strokes while a plain text box beside them in the same group keeps its own, that every shape's resolved bounds are unchanged, and that a second reset changes nothing further. Browser coverage drives the ribbon on a rotated group in English and Japanese, checks the untouched second slide, one-step undo/redo and a reload.
- Core tests verify idempotence, imported picture placeholders and saved round trips. Browser tests cover selected-slide scope, formatting and deleted slots, undo/redo, both languages and saved reload. Grouped layout reset and the other outstanding workflows above remain unfinished.

### Editing objects inside groups

- Double-click a group to select its children in place; the bilingual group navigation button or Escape returns to the parent. Select All stays within the current group. Nested coordinate transforms include rotation, reflection and nonuniform scaling for selection overlays, pointer movement and screen-direction keyboard nudges.
- Delete and z-order commands now operate on the owning group, retaining unrelated siblings and trailing extension metadata. Copy and duplicate use `copyShape` with `preserveGroupTransform` to retain ancestor transforms without copying siblings, including transforms that cannot be flattened to a standalone shape.
- Core tests verify nested ordering, deletion, copied ancestor geometry and save/reload. Browser coverage verifies rotated-child drag, keyboard movement, duplicate/delete, Undo/Redo, English/Japanese navigation and saved reload. Matrix tests cover nested reflection and nonuniform scaling.
- Remaining: broader resize/rotation gesture coverage and paragraph/body layout parity. These are not claimed complete by the movement tests.

### Visible geometry for snapping and arrangement

- Smart guides snap the visible rotated envelope in slide coordinates, including all ancestor group transforms. Snap tolerance remains tied to screen pixels, and guides stay in slide space. The resulting translation is converted back into group coordinates without changing size, rotation or reflection.
- Align and distribute use visible edges rather than unrotated local rectangles. Slide alignment works for grouped children; equal-gap distribution keeps its visible endpoints fixed.
- Regression tests cover six alignment directions and both distribution axes under combined rotation, reflection and nonuniform scaling, with Undo restoring original geometry. The previous controller fails the same test. Snap tests cover rotated children, reflected scaling and screen-distance thresholds; browser coverage checks the guide at the slide edge, saved snapped geometry, alignment and Undo.

### Shape text input under group transforms

- The shape text input now follows the preview's text-specific rotation, including a vertically flipped shape's half-turn, ancestor reflection compensation and group scale cancellation. Its expanded layout box stays centered on the selected shape while glyph transforms match the saved preview.
- Browser regression coverage compares actual CSS/SVG glyph matrices and input/selection centers for all four ancestor flip combinations, with rotated text inside one or two rotated, nonuniformly scaled groups. The nested cases add a reflected outer group, covering both reflection cancellation and accumulation across ancestors. Each case edits Japanese/English content and checks persisted text and Japanese UI after reload. Existing rich-text input and inherited-format workflows remain covered.
- Nested-group unit coverage verifies accumulated text scale. Paragraph/body layout parity and the renderer's anisotropic-scaling fidelity remain separate checks.

### Grouping within existing groups

- Group and Ungroup operate within the selected objects' immediate parent, preserving all outer group transforms and unrelated siblings. Grouping keeps member stacking order even when selection arrives in reverse order, and rejects selections spanning different parents instead of acting on a subset.
- The core round-trip test verifies nested grouping and release inside a rotated, reflected, nonuniformly scaled ancestor. Editor model coverage verifies selection restoration through Undo/Redo and mixed-parent rejection. Browser coverage exercises both keyboard commands, history and saved reload in English and Japanese.
- Ungrouping a resized group still has the documented affine limitation: nonuniform scaling combined with rotated children can introduce shear that a standalone child transform cannot represent. Supporting nested ownership does not resolve that fidelity limitation.

### Table cell input under transforms

- Cell hit testing reverses table rotation and both flips after mapping the pointer through ancestor groups. Selected-cell highlights and input centers use the same reflected cell positions.
- Cell text input follows the preview's table text rotation and reflection cancellation while retaining ancestor scaling. This differs from shape text, whose renderer cancels group scaling for glyphs.
- Browser regression coverage clicks the rendered cell independently of editor geometry, verifies its text, selection/input centers and CSS/SVG glyph matrices, and checks saved cell contents and Japanese input after reload. It covers all four table flip combinations, with and without a rotated, reflected, nonuniformly scaled ancestor. Paragraph/body layout parity remains separate from these transform checks.

### Text body layout while editing

- Inline editing uses effective shape body margins and vertical alignment, including inherited shape settings and default autoshape centering. Table cells use their own margins and anchor. Insets follow canvas zoom, and the selection outline no longer consumes text layout space.
- Browser regressions verify asymmetric margins and top/center/bottom placement in text boxes and table cells, with bilingual text editing and Japanese saved reload. Existing paragraph/selection and transformed-group editing checks guard the unchanged contenteditable DOM structure.
- Complete renderer paragraph fidelity remains open. Preset region coverage is described below; autofit and vertical writing have their own section.

### Vertical writing and autofit while editing

- `<a:bodyPr vert=…>` and `<a:bodyPr numCol=… spcCol=…>` become CSS through `verticalTextStyle` / `textColumnsStyle` in `@office-kit/pptx-preview`, which the renderer's `<foreignObject>` path and the inline editor both use. One mapping, so the caret reads in the same direction as the painted glyphs. The half turn `vert270` needs is reported apart from its writing mode, because the editor already rotates the editing box for shape rotation and has to compose the two.
- `shapeAutoFitScale` reports the factor the preview shrinks a `<a:normAutofit/>` body by, for the box the caller lays it out in. Inline editing scales its text by it, so a shrunk title keeps its size when the caret appears. `<a:noAutofit>` and `<a:spAutoFit>` report `1`, since PowerPoint shrinks neither.
- The factor is computed from the committed model, which is what the preview painted. Text typed into an autofit box therefore keeps the current factor until the edit commits, when both sides pick up the new one together. Live re-shrinking per keystroke is not implemented.
- Browser coverage compares the editor against the rendered SVG for `vert`, `wordArtVert` and `vert270`, and compares the shrunk box against an unshrunk reference beside it so the two px scales are commensurable. Both halves check bilingual editing and Japanese saved reload. Table cells still edit horizontally; cell-level `vert` is not wired up.

### Preset text regions while editing

- The preview renderer and inline editor share the preset text rectangle calculation. Triangles, diamonds, pentagons, five-pointed stars and double arrows retain their constrained text regions and asymmetric body margins while editing. Autoshape paragraphs without authored or inherited alignment now use the preview's centered default.
- Browser coverage compares the editing content rectangle against the rendered SVG foreignObject for six presets, a nonuniformly scaled group, and margins that collapse a star's region. It checks bilingual edits and Japanese saved reload. Renderer tests cover degenerate inset fallback and existing SVG/audit behavior.
- This matches the preview's existing preset approximations. Adjustment-dependent geometry still requires renderer work; custom-geometry text rectangles are described below.

### Custom-geometry text rectangles

- `getShapeCustomGeometry` reports the optional `<a:rect>` as `textRect`, evaluated against the same guides as the path commands, so a custom shape's stated text rectangle is readable through the public API.
- The preview renderer and the inline editor both lay text in that rectangle, through `shapeCustomTextRect` in `@office-kit/pptx-preview`. A custGeom shape that states no rectangle keeps its whole box: the preset table is an approximation of shapes we do not have the geometry for, and substituting it for a shape that describes itself would move text the file placed.
- The rectangle is written in the shape's own `<a:ext>` space, which is what turns it into fractions. A group's scale reaches the text bounds but not that extent, so it is applied once, never twice.
- Renderer tests pin the stated rectangle, the whole-box fallback and a group-squashed shape. Browser coverage compares the editing content rectangle against the rendered SVG foreignObject for a custGeom shape with and without a nonuniformly scaled group, and checks bilingual editing and Japanese saved reload.

### Default paragraph alignment readback

- The inline toolbar and paragraph panel now report centered alignment for autoshapes with no authored or inherited alignment, matching their rendered text. Text boxes, placeholders and table cells retain the left-aligned fallback. Editing layout and both controls share the editor's shape defaults.
- Browser coverage checks English and Japanese readback for preset shapes, text-box/table defaults, explicit right alignment and Undo/Redo. Mixed or inherited object-level controls retain their existing distinction between local settings and effective paragraph values.

## Google Slides feature inventory

What a Google Slides user reaches for, and how the editor answers it today.
Every mutating library export is already reachable through the command palette
and the properties panel (see the [editor README](../site/src/lib/editor/README.md));
this table is about the everyday paths, and about what is not there at all.

| What a user does                               | Today                                                                                                                       |
| ---------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| Add / duplicate / delete / reorder slides      | Navigator, ribbon, keyboard, context menu                                                                                   |
| Change a slide's layout, background, size      | Ribbon and slide panel; new-slide dialog picks a layout                                                                     |
| Type and format text                           | Inline editing with IME, inline toolbar, paragraph panel; bold/italic/underline/strike, super/subscript, font, size, colour |
| Lists, indentation, line and paragraph spacing | Paragraph panel and inline toolbar, Tab / Shift+Tab for levels                                                              |
| Vertical writing, autofit                      | Inline editing follows the renderer (see above)                                                                             |
| Insert / arrange objects                       | Ribbon insert, marquee and Shift multi-select, snapping, align, distribute, group, order, rotate, flip                      |
| Images                                         | Insert, replace, crop, mask with a shape, brightness / contrast / opacity, alt text                                         |
| Tables                                         | Cell editing and selection ranges, borders, fills, margins, row/column edits, merge / split                                 |
| Charts                                         | Chart dialog creates and edits categories, series and formatting                                                            |
| Links                                          | Link dialog, including links to another slide                                                                               |
| Comments                                       | Thread dialog with replies, resolve / reopen, both comment formats                                                          |
| Transitions and animations                     | Transition dialog, animation pane with playback                                                                             |
| Speaker notes, presenter view                  | Notes dialog, presenter view                                                                                                |
| Find and replace                               | Quick find and the find/replace dialog                                                                                      |
| Undo / redo, zoom, copy / paste                | Throughout, one undo step per gesture                                                                                       |
| Japanese / English                             | Live switch, persisted, in every dialog and label                                                                           |

### Gaps, in the order they are worth closing

1. ~~**Format painter** (書式のコピー/貼り付け).~~ Done — see below.
2. **Character-level effects.** `setShapeGlow` / `setShapeShadow` apply to a
   shape; a text run has no outline, shadow or glow, so WordArt-style text cannot
   be authored or round-tripped as such. Library work first.
3. **Media playback.** `addSlideMedia` embeds a clip and its poster, but nothing
   states autoplay, loop, volume or a trimmed range, so a deck with a video plays
   it the way PowerPoint defaults to. Library work first.
4. **Slide number, date and footer.** The text is reachable through
   `setSlidePlaceholders`, but there is no "insert slide number" path and no
   deck-wide toggle, which is how Google Slides presents it.
5. **Layout and master editing.** The library can apply a layout and reset a
   slide to it, but not author one. "Edit theme" is therefore out of reach
   entirely. The largest of these by far, and the one to design before building.

Out of scope on purpose: real-time collaboration, version history, sharing and
publishing, spell check, and Explore-style suggestions — none of them are
properties of a `.pptx` file.

### Format painter

- Copying formatting reads a whole object — fill (solid, gradient, pattern or an explicit none), outline with its dash, cap, join, compound and arrowheads, shadow, glow, and the character and paragraph formatting its text starts with — or, while editing text, the character format at the selection plus the properties of the paragraph it starts in. Pasting puts an object pickup on every selected object, and a text pickup on the selected range and the paragraphs it touches. The text itself never travels.
- Everything goes through the library's own readers and writers, never a lift of the source's XML, so a pasted format is data an author could have set by hand. That is also the boundary: a picture fill, an inner shadow, and reflection / soft edge / blur have writers the library does not yet pair with a copyable reader, so they are named in the toast instead of being dropped quietly. A source with no text says nothing about text, and pasting it leaves the target's own text formatting alone rather than clearing it.
- Alignment and bullets that the source inherits rather than authors are left alone on the target: there is no writer for "inherit", and the target's own inheritance is the closest thing to what the source shows.
- Reachable from the object right-click menu, the inline text toolbar, and Ctrl/Cmd+Alt+C / Ctrl/Cmd+Alt+V in both places (`code`, not `key`, because Alt rewrites the character on macOS). Site tests cover the clipboard's own rules including the save/load round trip; browser tests copy between objects in English and Japanese, undo in one step, and repaint one text selection from another, all verified against the saved `.pptx`.
- This work also lifted a library limitation it ran into: `setShapeShadow` and `setShapeGlow` used to replace the whole `<a:effectLst>`, so a shape could never carry both. They now compose in the order `CT_EffectList` states.
