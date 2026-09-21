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

- Bilingual speaker-note dialogs load and edit multiline notes for the selected
  slide. Notes readers/writers target the body placeholder, preserve unrelated
  footer placeholders, and read soft breaks and field text. Browser tests verify
  separate notes on two slides, cancellation, undo/redo and reload.
  Transition dialogs configure effects, effect-specific directions, speed,
  through-black fades, click/automatic advance and optional application to all
  slides in one undo step. Browser tests verify push, fade, split, no-effect
  timing, reset, isolation of other slides and persisted settings. Animated
  playback of the configured effects remains to be verified.

- A bilingual slide options pane selects layouts already in the document, sets
  background colors or uploaded images, and resets to the inherited background.
  Browser tests verify layout relationships, untouched text and neighboring
  slides, embedded image bytes, reset, undo and reload. Bilingual page setup
  supports 16:9, 4:3, 16:10 and custom dimensions in inches or centimeters.
  Browser tests verify live canvas ratios, saved dimensions, unit conversion,
  invalid input, cancellation, undo/redo and unchanged object geometry.
  Content scaling and fuller layout/placeholder workflows remain outstanding.

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

- Slide commands bind the active slide rather than asking for internal objects.
  Navigator buttons, context menus and keyboard navigation cover insertion,
  duplication, deletion and reordering. Browser tests verify Japanese labels,
  undo selection, persisted order and reload. Slide clipboard, layout selection
  and multi-slide selection still need complete UI workflows.

- The Arrange pane aligns objects to the selection (or slide for a single object)
  and distributes equal horizontal/vertical gaps. Group/ungroup bind selected
  top-level objects and preserve stacking order and undo selection. Canvas and
  Select All treat nested groups as one object. Browser tests cover bilingual
  controls, moving a group, ungrouping, undo, saved geometry and reload.
  Rotated-group transforms and group-child editing still need workflow verification.

- Image replacement detaches shared media and relationships so copied pictures
  remain independent. Tests cover same-slide and cross-slide copies, same-format
  and cross-format replacement, repeated edits and saved PPTX reloads. Rejected
  crop edits retain the previous crop through a later successful edit and save.
  Bilingual upload and replacement dialogs and image options now cover numeric
  crop, opacity, brightness, contrast and alternative text. Browser tests verify
  insertion, replacement, undo, crop reset and persisted output after reload.
  Preview contrast treats zero as neutral and scales colors around mid-gray.
  Interactive crop handles still need a dedicated workflow.

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
against a serialized deck. Visual transition effects remain outstanding.

A separate presenter window now displays the current and next slides, saved
speaker notes, an elapsed timer with reset, and previous/next/exit controls.
It follows the editor's English/Japanese locale and synchronizes navigation with
the audience window. Notes render as plain text outside the audience surface.
Browser checks cover multiline Japanese/English notes, literal HTML in notes,
empty notes, saved note updates, last-slide navigation, popup reload, locale
switching, and a closed opener. Animated transitions remain outstanding.
