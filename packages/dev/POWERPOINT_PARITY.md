# PowerPoint UI and operation parity

The requested target includes operations that were absent from the original
preview. Matching only the existing feature set is not sufficient. The current
reference, confirmed by the user, is the Mac desktop PowerPoint interface; a platform/version change
requires a new visual comparison. This implementation is not yet full parity.

The overview below reflects the current implementation, rather than the original
preview baseline. “Available” means implemented; it does not certify a match to
native Mac behavior or appearance. The dated-by-sequence notes below retain the
history of implementation and verification; their historical “remaining” lists
may have been addressed by later entries.

| Area | Available | Remaining for parity |
| --- | --- | --- |
| Workspace | Ribbon tabs, thumbnail and selection panes, pane resizing, status bar, zoom, normal/sorter views, editable persistent drawing guides and smart alignment/spacing guides | Exact native chrome, spacing and complete ribbon groups, rulers, native guide details |
| Objects | Selection/marquee/multiple selection, move/resize/rotate, duplicate/delete, group/ungroup, stacking, alignment/distribution, connection attachment, fill/stroke controls | Selection within groups, complete native snapping and multi-selection geometry, freeform editing |
| Text | Rich text and caret formatting, font/paragraph dialogs, case changes, search/replace, columns/direction/anchoring, autofit and persistent history | Complete bullet schemes, native text layout/caret/selection, remaining dialog and gallery fidelity |
| Insertion | Categorized shape and action-button galleries, lines/arrows/connectors, pictures, tables and direct cell editing | Charts/media, freeform drawing, automatic connector routing/obstacle avoidance, photo browser/stock/online pictures |
| Pictures and tables | Picture replacement/crop/fit/fill/corrections/transparency; table formatting, cell selection/merge/split, row/column operations and resizing | Complete native styles/effects and contextual UI, all imported-content edge cases |
| Slides | Layout insertion/application/reset, duplicate/delete/reorder/hide, notes, sections, page setup with content scaling | Complex placeholder mapping/reset edge cases, master/layout editing, themes, exact native section interactions |
| Clipboard | Snapshot object copy/cut/paste across slides, text clipboard and context menus | OS object clipboard, full rich/inter-application paste and native paste-format semantics |
| History and files | Persistent undo/redo journal, conflict detection, replay and PPTX export | Native open/save workflow, source reconciliation UI, shared edit history |
| Slide show | Navigation/shortcuts, custom shows/show settings, action links/sounds, presenter/audience views, transitions/timing, pointer/ink and retained annotations | Full effect/media coverage, physical-display verification and exact native presenter/annotation behavior |
| Animations | Appear/disappear/fade playback and authoring, start/duration/delay, pane selection/context menu, batch editing/removal, drag reorder and selected preview | Complete effect gallery/options, paragraph/trigger/repeat controls, complex native timing and exact native pane behavior |
| Review | Agent panes preserved | Comments, spelling, accessibility review and collaboration |

Edits persist in `<entry>.edits.json`. Keep the journal with the TSX source.
A conflict preserves both files and stops replay; undo history remains available
for recovery. Direct text input records range replacements, preserving unaffected runs and links across insertions, deletions and paragraph splits/joins. Applying character formatting to an existing selection preserves the surrounding runs and hyperlinks, including after journal replay and PPTX export. Insertion-point font commands apply to subsequently typed text, and moving the caret resets that pending style. The contenteditable editing overlay displays character-level font, color, highlight and decoration differences, and reads effective body margins, vertical anchoring, wrapping and column settings from the deck. Input has local undo/redo, with composition updates grouped into one undo step. Paragraph alignment, bullets, spacing and outline level apply to the selected paragraphs or the paragraph at the caret during text editing; selecting a shape applies them to all its paragraphs. Paragraph commands restore the text selection after saving. The editing surface also renders paragraph alignment, line/before/after spacing, indents, text direction within horizontal paragraphs, standard bullets and decimal numbering. Paragraph splits and joins inherit the same paragraph properties as saved text. Custom numbering schemes, bullet fonts/colors/images, vertical text, non-rectangular shape text regions and rich clipboard paste still need parity work.

The Shapes gallery follows the native Mac category order for lines, rectangles, basic shapes, block arrows, equations, flowcharts, stars/banners and callouts. Icons reuse the preview geometry; some curved and callout geometries remain approximations. Straight lines, arrows and elbow/curved connectors preserve drag direction and arrowheads through saving and replay. Shift constrains straight lines to 45-degree increments. Selected lines expose start/end handles with drag previews, direction reversal and Shift constraints; rotated lines retain the fixed endpoint. Line endpoints snap to native shape connection sites, highlight available sites and can be detached by dragging away (Alt bypasses snapping). Connections follow target movement, resizing, flipping and rotation, including transformed groups; root-level targets also show a connector preview during dragging. Native OOXML connection IDs/site indices persist through PPTX export and journal replay. Copying a selection remaps connections among copied shapes; copying a connector alone detaches it. Automatic orthogonal rerouting, obstacle avoidance, connection-site selection within groups, full-styled drag previews and native connected-handle appearance still need parity work. Action buttons and several missing presets are also not yet included.

Verification covers browser editing/persistence, shape insertion/move/resize,
repeated paste, marquee selection, distribution, notes, undo/redo, serialized replay, conflicts, group selection,
PPTX output and CLI replay, picture insertion and text/paragraph formatting, visible mixed character styles, newline undo/redo and simulated composition undo. Native macOS IME behavior still requires manual verification. Pictures are embedded in the journal (20 MB per picture limit). Existing browser tests also exercise live source
updates, presentation controls and agent pane behavior. This is functional
verification, not a claim of pixel or complete behavior equivalence.

The New Slide split button offers the deck’s layouts in a keyboard-accessible thumbnail gallery. New slides retain native layout relationships and empty placeholders through saving, undo/redo and journal replay. Empty placeholders show editing-only prompts and accept a single click to begin text entry. Gallery thumbnails currently show schematic placeholder bounds; complete theme previews, inherited prompt typography, content insertion buttons, complex picture/table/chart placeholder mapping and native reset edge cases remain unfinished. Existing-slide Layout and Reset reconcile text slots, preserve content and ordinary shapes, add missing placeholders, and reset placeholder geometry and formatting. Unmatched content retains its position. Tests cover persistence, replay, UI undo, and OOXML schema validity. The user authorized temporary mutations and undo in `/tmp/pptx-macro-audit/support.pptx`. Native mutation-based comparison is pending because subsequent Computer Use observations failed with a ScreenCaptureKit capture error; no layout mutation was verified.

Layout/Reset behavior is also referenced against Microsoft’s [Apply a slide layout](https://support.microsoft.com/en-us/powerpoint/training/apply-a-slide-layout) documentation. This does not establish complete Mac behavior equivalence.

Sections are stored in the native presentation section list and participate in journal replay and undo/redo. Adding/duplicating/deleting/moving slides updates section membership. The Home Section menu, thumbnail context menu, Rename Section dialog, section counts and collapse/expand controls are implemented. Automated checks cover persistence, replay conflicts, undo, exported PPTX, and browser interaction. Section context menus also support moving whole sections up/down and removing a section with its slides; headers support drag reordering with an insertion indicator. Tests verify retained slide identities/order, empty sections, destructive-edit undo, replay and PPTX serialization. Trailing empty sections appear after the last thumbnail, and empty decks retain editable section headers. Header selection is separate from the disclosure control. Command+Up/Down moves the focused slide or section, and adding Shift moves it to the beginning/end; section movement retains keyboard focus. Arrow navigation skips collapsed slides, and Shift+F10 opens the focused section menu. Command+Shift+N and Command+Shift+D add and duplicate slides. These shortcuts follow Microsoft’s [Mac keyboard reference](https://support.microsoft.com/en-us/accessibility/powerpoint/use-keyboard-shortcuts-to-create-powerpoint-presentations). Browser checks cover selection without collapse, section keyboard movement, focus retention and keyboard context menus. Full native selection semantics and exact Mac appearance still require direct comparison. The baseline workflow is documented in Microsoft’s [Organize slides into sections](https://support.microsoft.com/en-au/powerpoint/training/organize-your-powerpoint-slides-into-sections); direct Mac comparison remains unavailable due to the capture error.

Mac object shortcuts now route Command+Shift+B/F to back/front, Option+Command+Shift+B/F to one layer backward/forward, and Option+Command+G / Option+Command+Shift+G to group/ungroup. Physical letter codes also handle Option-modified character values. Object commands take precedence over character formatting outside text editing. Browser coverage exercises all four stacking shortcuts and group persistence across reload before ungrouping, retaining child identities and geometry. The mappings follow the Microsoft Mac keyboard reference linked above; regrouping and full text-editing shortcut parity remain unfinished. Option+Left/Right rotates selected objects by -/+15 degrees relative to each existing angle, with normalized angles across full turns. Relative rotation is one undoable journal operation and retains independent bounds/angles through export and replay. Tests cover multiple different starting angles, wrapping, inverse rotation, undo/replay and browser key routing.

Native keyboard resizing was observed in `/tmp/pptx-macro-audit/support.pptx` after AX reads and keyboard actions recovered (coordinate clicks still failed). With aspect lock off, Shift+Right/Up multiplies width/height by 1.1 about the object's center; Shift+Left/Down divides by 1.1. Observed width 10.99 → 12.09 / 9.99 cm and height 1.52 → 1.68 / 1.39 cm, with corresponding center-preserving position changes. Temporary dimension changes were undone and baseline width/height/position values verified. The editor implements these four gestures with EMU rounding and minimum dimensions, replacing its previous Shift+Arrow large-nudge behavior. Browser checks verify dimensions, fixed centers and undo in all four directions.

With native aspect lock enabled, Shift+Right changed both dimensions (10.99 × 1.52 → 12.09 × 1.68 cm), retaining the center; undo and restoring the original unlocked setting were verified. The editor reads DrawingML `noChangeAspect` locks and scales both dimensions for locked objects. Tests cover picture-lock persistence/export/replay and centered keyboard resizing with undo. Command/Control/Option+Shift+Right produced no native geometry change, and those combinations no longer fall through to movement in the editor. Right and Command+Right both moved the reference object by one point. Rotated/grouped/connector-specific resizing and pointer resizing with aspect locks remain unfinished.

The object context menu now opens a Size and Position pane with native-ordered Height, Width, Rotation, Lock aspect ratio, Horizontal position and Vertical position controls. Dimensions use centimeters. Native numeric Width changes were confirmed to retain the top-left position both with and without aspect lock (locked 10.99 × 1.52 → 12 × 1.66 cm); mutations were undone and lock restored. Numeric size edits preserve each selected object’s ratio when locked. Mixed selection values are blank/indeterminate. Lock changes persist in DrawingML, export, replay and undo, with lock-specific source conflict detection that retains legacy journal compatibility. Browser checks exercise the context menu, lock toggle, numeric resizing, fixed position and undo; the pane screenshot was inspected. Complete category tabs, Scale Height/Width, original-picture size, position reference selectors and exact native spacing remain unfinished.

The Size & Properties pane now includes a Text Box disclosure with Top/Middle/Bottom vertical alignment, four centimeter margin fields (0–55.88 cm), and Wrap text in shape. Native labels, ordering and limits were observed in the reference deck’s AX tree. Margin and wrap changes serialize to DrawingML and support journal replay, source-conflict detection and undo without changing legacy fingerprints. Mixed text selections display blank/indeterminate values; selections containing non-text objects disable these controls. Tests cover exported values, replay, undo, input validation and browser margin/wrap edits. Centered anchor variants remain unfinished; exact native spacing and temporary native mutation comparison for these settings remain to be verified.

The Text Box Columns button and Home > Columns > More Columns open a modal with Number of columns (1–16), Spacing between columns (0–40.64 cm), Cancel and OK. Native temporary edits confirmed both two columns at 0.5 cm and one column retaining 0.5 cm; both were undone and original one-column/zero-gap values verified. The core API now supports an explicitly authored single column with a gap. The editor preserves unchanged fields independently for mixed selections, commits both fields as one history entry, and detects changed source column settings. Backend checks cover export/replay/undo, mixed values and bounds. Browser checks cover apply/cancel/undo. Exact visual fidelity and multi-column text flow still require broader comparison.

The Text Box pane now exposes Horizontal, Rotate all text 90°, Rotate all text 270° and Stacked directions using the same persisted command as the ribbon. The direct editing overlay reads the effective body direction, retains vertical writing and upright/sideways orientation, and reverses physical margins for the 270° transform. Paragraph indents and spacing use inline/block axes. Backend checks cover all four values through PPTX export, replay and undo; browser checks exercise the three vertical choices and their editing styles. Native direction-menu comparison was attempted but remains unverified because ScreenCaptureKit reads failed again. Complete vertical text layout, mixed-script orientation, overflow and caret parity remain unfinished.

The Text Box pane now includes the three native Autofit radio choices, reading inherited layout/master modes and persisting explicit none/normal/shape choices in DrawingML. These settings support export, replay, undo and mode-specific conflict checks without changing legacy history fingerprints. A regression test also found that selecting horizontal text removed the override and reintroduced inherited vertical text; horizontal now writes an explicit `vert="horz"`. Core tests cover inheritance and explicit overrides. Automatic shape geometry growth and shrink behavior during direct editing remain unfinished; the existing preview shrink path handles authored normal autofit, including inherited font and line-spacing scale parameters. Native mutation comparison was unavailable because the ScreenCaptureKit error persisted.

The effective body-property resolver now carries autofit mode and scale parameters from the same nearest shape/layout/master declaration. An explicit normal autofit without scale starts from 1/0 instead of retaining an ancestor’s baked scale; none/shape overrides discard inherited shrink parameters. Both SVG and foreignObject preview paths consume this resolved value. Regression coverage verifies a 0.65 inherited font scale in both renderers, overrides, and existing bare-normal shrink behavior. Direct-edit shrink/grow behavior remains separate outstanding work.

The direct text editor now receives the effective baked autofit parameters as well. It retains the font scale on entry, input, composition completion and local undo/redo, and applies the line-spacing reduction to proportional and fixed-point spacing. Text-frame margins and column gaps retain their physical dimensions. Backend tests cover inherited parameters and mode overrides; a browser regression checks mixed font sizes, both spacing types, text and selection preservation, and restoration without autofit. Recomputing shrink as content changes (including bare normal autofit) and automatic shape growth are still outstanding; this change only preserves the stored scale.

Normal autofit now recalculates on direct text input and composition completion using the live editing surface's overflow, restoring up to the authored font size when space permits. Bare normal autofit is measured on entry; baked scale remains unchanged until text changes. The computed scale is saved with the text command and participates in conflict detection, replay and undo; local undo restores the corresponding scale. The current search uses the preview's 25% floor. Browser tests cover fit/expansion and selection preservation, and backend tests cover combined text/scale persistence. This still uses browser editing metrics rather than PowerPoint's text engine; exact native break/scale parity, resizing outside text editing, font-format changes, and automatic shape growth need further work.

Shape autofit now measures natural content size during direct text input and composition completion. Wrapped horizontal text keeps its width and resizes height; vertical text keeps its height and resizes width; unwrapped text measures both dimensions. Measurement uses a hidden non-editable clone to preserve the active selection. The text anchor remains fixed through resizing, including rotation. Bounds are stored with the text command, and local/global undo restores text and geometry together. Tests cover natural horizontal/vertical measurements, selection preservation, anchor geometry, and export/history round trips. Native comparison is still blocked by ScreenCaptureKit capture failure. Exact PowerPoint typography, preset-shape text rectangles, setting-triggered resizing before text input, and formatting-triggered resizing remain unverified or incomplete.

Selecting shape autofit in the properties pane now measures the existing text immediately and submits the selected shapes' new bounds with the mode in one history command. Measurement shares the direct editor's typography, margins, direction and column setup. Mode undo restores the previous dimensions; the browser regression checks this before and after a subsequent text edit. Rendering an inactive measurement field no longer moves the active text selection. Formatting-triggered and geometry-triggered autofit updates still need integration; native typography/preset-geometry fidelity remains outstanding.


Text Box alignment now includes Top/Middle/Bottom Centered, using DrawingML `anchorCtr` independently from paragraph alignment. The flag inherits from layout/master placeholders, supports explicit false, and is persisted with conflict detection/replay/undo without changing older journal fingerprints. SVG layout translates the complete measured text bounds; browser preview uses that same SVG path for centered anchors. Direct editing translates paragraph content as a unit using browser range measurements and preserves text selection. Browser/native font metrics, vertical/column combinations, and exact native bounding-box semantics still require comparison; the Mac capture attempt again failed with ScreenCaptureKit -3811, without changing the reference file. Spec: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bodyproperties.anchorcenter?view=openxml-3.0.1

Formatting, text-frame changes, and dimension changes now recompute normal/shape autofit before saving. A revision-checked preview command obtains the effective post-edit text model without changing the presentation or history; the client measures that model and saves per-shape fit results in the original command. Pure movement/rotation skips measurement, and direct text entry retains its live measurement path. Fit targets and modes are validated, and history fingerprints include effective fit/frame settings while preserving older records. Backend coverage checks combined format/fit export, replay, undo and invalid targets/modes. Native manual-resize semantics, browser/native typography differences, and preset text rectangles remain unverified.

Validation for this increment: all 29 editor backend tests and all 3 browser tests pass. The browser checks a preview response without changing the live model/revision, font-size-triggered shape growth, and restoration of the full shape with one undo. Dev build, TypeScript, lint and whitespace checks pass.

Corner-handle resizing now honors the authored aspect-ratio lock as well as Shift. Proportional resizing uses the dominant relative width/height displacement, so a vertical-only corner drag also changes the size; both dimensions retain a one-point minimum without distorting the ratio. Existing rotated opposite-corner anchoring remains in use. Edge-handle behavior is unchanged pending native comparison. Microsoft documents Shift+corner proportional resizing: https://support.microsoft.com/en-US/Office/graphics-visuals/change-the-size-of-a-picture-shape-text-box-or-wordart . The Mac AX observation again failed with ScreenCaptureKit -3811; no native file mutation was performed. Exact native pointer projection and lock/Shift interaction remain unverified.

Validation: the canvas browser integration passes with a vertically dragged locked corner, proportional dimensions, stationary opposite corner, and full-shape restoration by undo. Dev build, TypeScript, lint, and whitespace checks pass.

Drag gestures now request a transient render of the affected slide using the same server-side SVG renderer as saved edits. This covers resizing, rotation, moves and connector endpoints without scaling text glyphs as an image. Requests are coalesced, serialized, revision/slide checked, and discarded after cancellation or newer pointer input. Temporary SVG paint-server/clip IDs are namespaced to avoid collisions with the saved slide. The saved SVG is restored on release/cancel; saving waits for an outstanding preview request and writes one history entry. Escape cancels the active drag while preserving the selection. Native frame timing and rendering fidelity remain unverified; this implementation has an 80 ms coalescing delay plus server rendering latency, and autofit recalculation still occurs at commit rather than in transient rendering.

Validation: the canvas browser integration passes after the single-slide rendering optimization. It verifies a rendered drag preview with unchanged revision/model, matching preview/committed geometry, cleanup after release, undo, and Escape cancellation restoring the saved SVG without a revision change. Dev build, TypeScript, lint and whitespace checks pass.

Move gestures now snap visible selection bounds to the nearest edge/center of another object or the slide within five screen pixels. Red dashed smart guides span the aligned selection and target. Multiple selected shapes share one translation, rotated bounding boxes are included, and Shift-constrained moves retain their constrained axis. Holding Command while dragging suppresses guides/snapping (the existing Option bypass is also retained). Guides disappear on commit or cancellation. Pure geometry tests cover nearest target choice, misses, multiple selection, rotation and axis constraints. Native threshold/tie-breaking/guide extents, equal-spacing guides, guide preferences/menu controls, and exact rotated-shape semantics remain unverified or incomplete. Mac interaction reference: https://support.microsoft.com/en-us/office/graphics-visuals/align-or-arrange-objects and https://support.microsoft.com/en-us/powerpoint/work-with-gridlines-and-use-snap-to-grid-in-powerpoint .

Validation: the canvas browser integration passes, covering visible snap alignment, Command bypass and Escape cleanup without model/revision changes. Guide spans now use the final position after both axis corrections; the geometry regression verifies their endpoints during simultaneous horizontal/vertical snapping. All three geometry tests, TypeScript, lint and whitespace checks pass after this correction. The dev build and browser integration passed before the final span-only correction.

The View ribbon now exposes Guides and a Grid and Guides menu with independent Guides / Smart Guides switches; the canvas background context menu exposes the same controls. Guides draws the two center lines and enables center-guide snapping even when smart alignment is disabled. Disabling Smart Guides suppresses object/slide smart snapping and red dynamic guides. Preferences survive browser reload through local storage without changing the presentation, revision or undo stack; drawing guides live only in the editing overlay and are hidden during presentation. Native guide movement/addition/deletion/colors, grid spacing/snap controls, viewProps import/export, the native application menu placement and exact styling remain outstanding. These browser preferences do not yet round-trip through PPTX. The native observation attempt again failed with ScreenCaptureKit -3811 without modifying the reference deck. Reference: https://support.microsoft.com/en-us/powerpoint/work-with-gridlines-and-use-snap-to-grid-in-powerpoint .

Validation: all four guide geometry tests and the canvas browser integration pass. The browser verifies disabled snapping, visible center guides, menu/checkbox state, reload persistence, and unchanged document/revision. Dev build, TypeScript, lint and whitespace checks pass.

Drawing guides now have individual center-relative EMU offsets and colors. They can be dragged with a center-distance tooltip in centimeters, cancelled with Escape/pointer cancellation, added horizontally/vertically, deleted through their context menu (or dragged off the slide), and recolored through the Color submenu. Object alignment uses the actual guide positions and respects an empty guide list. Position/color/list state survives browser reload through local storage; hiding guides keeps their state. Guide pointer handling is isolated from shape movement and marquee selection. Browser coverage verifies movement, Escape restoration, tooltip cleanup, blue color selection, add/delete, reload restoration and unchanged presentation/revision. All five geometry tests, canvas browser integration, dev build, TypeScript, lint and whitespace checks pass. The browser test timeout increased to 90 seconds to accommodate the expanded interaction sequence.

Remaining guide fidelity: actual native palette/order/appearance, add-guide initial placement (currently one centimeter beyond the previous same-axis guide), movement quantization/modifiers, keyboard editing, undo/redo for guide edits, layout/master guide inheritance, and PPTX viewProps persistence. Browser local storage remains an interim implementation and does not establish native document parity. Mac documented behavior (moving guides shows distance from slide center, add/delete and context Color): https://support.microsoft.com/en-us/powerpoint/work-with-gridlines-and-use-snap-to-grid-in-powerpoint .

Presentation-level extended drawing guides now load from and save to the native `p15:sldGuideLst` extension in `ppt/presentation.xml`. The public API uses EMU from the slide edge and converts to PowerPoint master units (1/576 inch) at serialization. The editor converts these to center-relative positions. Move, color, add and delete now create document history entries, support undo/redo, survive reload and PPTX export, and detect changed source guides during replay. Guide coordinates/list/colors no longer use browser local storage; visibility and smart-guide preferences still do. Existing names, userDrawn attributes and extension metadata are retained for surviving IDs. Empty explicit lists stay empty. References: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/9ca7eff3-ea78-4547-9aaf-c18ca5e8319a and https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/cda224b4-e720-42fd-80a6-b3dc8efca272 .

Validation: two core guide tests (native coordinates, color round-trip, explicit deletion, quantization, metadata retention, invalid input atomicity), all 30 editor tests, and the canvas browser integration pass. The browser verifies guide changes through reload and undo/redo, while shape content stays unchanged. Remaining: legacy viewProps guides, layout/master inheritance, theme/system color resolution (currently retained in XML but displayed gray), visibility preferences in PPTX, and native UI/palette/movement fidelity. No native PowerPoint round-trip has been observed for this change because the last native capture failed with ScreenCaptureKit -3811.

Legacy `viewProps` slide drawing guides now provide the fallback when the presentation has no extended guide list. The loader follows the internal presentation relationship, including nonstandard part paths, and reads native master-unit positions with the XML orientation/position defaults. Existing legacy lists are synchronized on guide edits so earlier consumers do not retain stale positions or deleted guides. Notes-view guides, zoom, grid spacing and visibility settings are retained. Explicit extended empty lists take precedence over legacy data. Both XML documents are prepared before package mutation. Validation: three core guide tests pass, including custom relationship paths, native defaults, negative positions, save/reload, extended precedence, notes/settings preservation, schema validation and malformed XML atomicity. Remaining: creating legacy view settings in documents without them, layout/master inheritance, theme/system color resolution, visibility persistence and native UI fidelity. References: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.guide?view=openxml-3.0.1 and https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/20fd0c14-3ea5-4e45-8a72-bf495a8fcb48 .

Drawing-guide visibility now reads/writes the slide-view `cSldViewPr/@showGuides` attribute in PPTX instead of browser local storage. The Guides checkbox/context menu uses the document state; commands save and replay visibility and restore it through history. Adding a guide while guides are hidden records the new guide and visible state together. The core API creates schema-valid slide view properties and internal relationships when missing, avoids unrelated parts at the conventional path, and preserves existing zoom, snapping preferences and guide lists. Smart Guides remains a browser preference. Reference for the view attribute: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commonslideviewproperties?view=openxml-3.0.1 . Validation: four core tests (including schema validation and missing-part relationship creation), 31 editor tests, root/dev builds, TypeScript and lint pass. Native Mac behavior for whether visibility changes belong in Undo still requires direct comparison; this change does not establish complete native UI parity.

Browser validation also passes: toggling Guides waits for persisted document state, reload restores visible guides and their positions/colors, and the expanded editing sequence retains correct history cursor accounting with two visibility commands (37.7 seconds).

Smart alignment now also considers equal edge-to-edge gaps along a row or column. Moving one object or a multiple-selection bounding box can match the space between two stationary objects or extend their existing spacing on either end, including objects with different sizes and rotated visible bounds. Candidates require overlap on the perpendicular axis and exclude obstructed placements. The nearest correction wins against edge/center alignment; Shift limits the axis and the existing Command/Option bypass and Smart Guides setting also control spacing. Two red interval lines with end marks show the matching gaps while dragging. Geometry coverage now has eight passing tests. The exact Mac marker styling, candidate priority, same-size resize guides and spacing relative to slide edges still require work; this does not establish full native parity. Feature reference: https://support.microsoft.com/en-us/powerpoint/training/guides-for-arranging-things-on-a-slide .

Browser validation passes for equal spacing: two equal-length intervals with end marks appear, Escape cancels without a revision, Command suppresses snapping, and a committed move saves at the exact equal-gap position and survives reload. The existing full canvas editing browser test also passes. Dev build, TypeScript, lint and whitespace checks pass.

Slide-size implementation audit: Mac Page Setup/Slide Size offers Scale and Don't Scale when changing dimensions; content scaling must be implemented separately from the existing canvas-only API. The existing `setSlideSize` now validates both dimensions against the OOXML 1–56 inch range before mutation and rounds to whole EMU using the shared numeric-boundary helpers. Creating a missing `sldSz` now respects presentation schema ordering even without a slide list, preserving notes and other subsequent properties. Six slide-size tests and 38 enum-validation tests pass, including valid boundary sizes, rejected nonfinite/out-of-range values without mutation, no-slide schema ordering, and portrait save/reload. These are prerequisites; the Design ribbon, Page Setup dialog and Scale/Don't Scale content behavior are still pending. Reference: https://support.microsoft.com/en-us/powerpoint/change-the-size-of-your-powerpoint-slides . Native reference capture was retried and again failed with ScreenCaptureKit -3811; no reference document modification was performed.

Canvas-only slide resizing is now an editor history command. It validates numeric dimensions and the format hint, persists the presentation-wide size, and supports replay, undo and redo without changing slide contents. Its source fingerprint includes the original size (including the type hint), so a changed source canvas raises a replay conflict while older command fingerprints remain unchanged. Regression coverage checks portrait export/reload, content preservation, replay/undo/redo, source conflicts, invalid input without mutation, and presentations with no slides. The Design/Page Setup UI and Scale content behavior are still pending; this backend command alone does not claim native Don't Scale parity. Native capture was retried but still failed with ScreenCaptureKit -3811, without modifying the reference deck.

Validation for the canvas-size command: all 32 editor backend tests, TypeScript, lint and whitespace checks pass. No UI behavior was added in this increment.

Design now exposes Slide Size → Standard (4:3), Widescreen (16:9), and Page Setup. Page Setup edits custom width/height in centimeters and slide orientation, validates the supported range, and asks Scale / Don't Scale before applying a changed size. Escape/Cancel leaves the document unchanged, and a captured revision prevents applying a dialog against a newer build. The scaling option proportionally fits the previous canvas into the new one and centers it. It scales explicit slide/layout/master geometry, nested group coordinate spaces, font sizes, paragraph distances, insets, line widths, table row/column measures, and supported effect distances; percentage/angle/crop values, media bytes, and notes parts stay unchanged. Default body/table insets are materialized where there is no placeholder cascade. Existing top-level slide/shape handles refresh. All affected documents are parsed/serialized before writes; malformed layout input leaves package bytes unchanged. Source fingerprints include all slides and presentation/layout/master bytes for scaling commands.

Validation: eight core slide-size tests pass, covering grouped geometry, live handles, explicit text scaling, portrait and schema round-trips, validation, malformed-part atomicity and preservation of unrelated parts. All 33 editor tests pass; the scaling regression additionally checks a modified master causes a replay conflict. The Page Setup browser test passes for orientation cancellation, 50% geometry/text scaling, reload, undo/redo, and Don't Scale content preservation. The dialog screenshot was visually inspected. Root/dev builds, TypeScript and lint pass.

Remaining size fidelity: direct native Mac comparison of Scale/Don't Scale positioning and prompt conditions; full paper/on-screen presets, units, notes/handout orientation, start slide number and File menu entry; exact native dialog/ribbon layout; theme-referenced strokes and table/chart/SmartArt typography, implicit paragraph defaults, guide transformations, and extreme scaling bounds. The current proportional fit behavior is an implementation awaiting native verification, not proof of complete Mac parity. Microsoft workflow reference: https://support.microsoft.com/en-us/powerpoint/change-the-size-of-your-powerpoint-slides .

Final browser regression: the existing full canvas editor integration and the new Page Setup integration both pass after the Design tab and source-fingerprint changes (50.5 seconds combined). Final lint and whitespace checks also pass.

Page Setup now offers the 13 paper/screen presets in Microsoft's Mac size table, including Letter, Ledger, A3/A4, B4/B5 ISO, 35mm, Overhead, Banner, and both on-screen aspect ratios. The ribbon keeps its two quick presets. Preset orientation is retained when changing the selection; switching orientation keeps the preset identity. The editor model carries the saved slide-size type, distinguishing equal-dimension presets such as Letter and Overhead on reopen. A type-only change saves without scaling and participates in undo/reload. Same-type, same-size submissions remain no-ops. Save failure leaves the type-only dialog available for retry.

Preset dimensions follow the published inch table; rounded paper values (especially Ledger/A4/B4/B5) still require comparison with actual native EMU values. Widescreen uses the existing exact 12192000 × 6858000 EMU canvas. These additions do not establish exact native dimensions or dialog appearance. Unit suffix input, notes/handout orientation, slide starting number, File entry, and the previously listed scaling gaps remain pending. Source: https://support.microsoft.com/en-us/powerpoint/change-the-size-of-your-powerpoint-slides .

Validation for preset persistence: the Page Setup browser test and all 33 editor backend tests pass (34 total). Browser coverage includes Letter → Overhead without dimension changes or a scaling prompt, content preservation, reload, portrait retaining the preset, and undo to Letter. Dev build, TypeScript, lint, and whitespace checks pass.

Page Setup dimensions now accept explicit `cm` and `in` suffixes (case-insensitive) and unitless centimeters. Valid values normalize to centimeters on leaving the field; orientation compares converted dimensions, and save converts to integer EMU. Invalid syntax, empty values, nonfinite values, and dimensions outside 1–56 inches block submission without modifying the deck. Selecting a preset clears previous validation errors. The browser regression passes mixed-unit editing, normalization, invalid-input rejection, scaled persistence, reload, undo/redo, and equal-size preset identity. Dev build, TypeScript, lint, and whitespace checks pass.

The native capture retry again returned ScreenCaptureKit -3811 without changing the reference deck. Pixel-suffix conversion remains pending: the Microsoft support page documents acceptance of `px` but does not establish the current Mac installation's conversion basis. Do not substitute browser CSS pixel density as proof of native behavior. Current display units remain centimeters; OS measurement preference integration and exact native validation UI are also pending. Documentation: https://support.microsoft.com/en-us/powerpoint/change-the-size-of-your-powerpoint-slides .

Page Setup now includes Number slides from (0–9999). The value is stored in presentation.xml's `firstSlideNum` alongside size changes in one edit; changing only the number does not prompt for content scaling. An omitted value defaults to 1 and size-only API calls preserve it. Validation occurs before mutation; history fingerprints include the original starting number when that setting is edited. The editor model carries the value so reopening Page Setup and journal replay restore it.

Preview text now resolves slide-owned `slidenum` fields from the current deck order plus the first number, rather than displaying cached field text. It matches slides by part identity, including handles retained across reordering; other field types retain their cached text. Both SVG and foreignObject rendering are covered. This does not add a number placeholder to slides without one. Master/layout-owned field instances, number insertion controls, native thumbnail/status numbering behavior, and field editing still require additional parity work.

Validation: nine slide-size core tests, the Page Setup browser test, and all 33 editor backend tests pass. The browser covers number-only save, reload and undo with unchanged slide content. The renderer regression covers both rendering paths, numbering after reordering, and preserving unrelated cached fields. Root/dev builds and dev/preview TypeScript checks pass. Reference for the 0–9999 restriction: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/5cc61456-aa4e-4c48-a115-3efed3e57d95 . Number formula: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.slidenumber .

Page Setup now exposes a separate Notes, handouts & outline orientation group. Its setting reads `p:notesSz` and swaps the existing notes canvas dimensions independently of the slide canvas; absent notes size initializes the existing blank-deck default (7.5 × 10 inches). The core validates orientation and existing dimensions before committing, retains unrelated notes attributes, and creates notesSz immediately after sldSz when absent. The editor saves the setting in the same page-setup history entry and fingerprints the original notes size for replay conflicts. Notes-only changes do not trigger slide content scaling.

This is canvas-setting persistence, not verified native notes/handout layout parity: notes shapes and master geometry are not repositioned, print output/views are not yet implemented, and actual native behavior still needs comparison. The group naming and independent orientation follow Microsoft's Mac Page Setup documentation: https://support.microsoft.com/en-us/powerpoint/page-setup-options-in-powerpoint-for-mac . The scope of NotesOrientation is documented at https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pagesetup.notesorientation .

Validation for notes orientation: all ten slide-size core tests and 34 editor/browser tests pass. Coverage includes invalid orientation with no mutation, missing notesSz insertion with schema validation, round-trip dimensions, independent radio groups, unchanged slides, reload, and undo. Root/dev builds, dev TypeScript, lint, and whitespace checks pass.

File → Page Setup now opens the shared settings dialog from an application menu above the title bar. Both entry points finish pending text edits and capture the current revision before saving. The menu supports keyboard opening, Escape/focus restoration, pointer toggle, and expanded-state reporting; the header and collapsed/mobile grid heights include the new row. Browser coverage checks the File path and cancellation alongside the existing scaling/persistence/history regression. Dev build, TypeScript, lint, and whitespace checks pass; the updated dialog screenshot was inspected.

This is an in-page application menu with the Page Setup entry only. It is not the native macOS system menu, and the rest of File plus the other application menus remain incomplete. Exact Mac menu/dialog appearance and behavior still require native comparison. The File → Page Setup route follows the already reviewed Microsoft Mac Page Setup reference above.

The in-page application bar now includes Edit with Undo, Redo, Duplicate, Delete and Select All, sharing existing editor actions. Opening Edit preserves an active text editing session so its undo history remains available; Select All selects text in that session. Object Duplicate/Delete are disabled during text editing or without selected shapes. These are the first Edit entries, not a complete native Edit menu: text deletion, clipboard commands, thumbnail context, Find/Replace and further native entries remain pending.

Application menus now use a roving tab stop, Left/Right navigation between File and Edit while open or closed, Home/End navigation when closed, pointer rollover between open menus, and menu-origin expanded state. Shortcut hints are displayed separately from accessible command labels. The browser regression passes File → Edit keyboard navigation, disabled Duplicate with no selection, Select All, duplicate, Undo and Redo, followed by existing Page Setup persistence/scaling/history checks. Dev build, TypeScript, lint and whitespace checks pass. Native reference capture retried this turn and still failed with ScreenCaptureKit -3811; no reference document was changed. Native ordering, appearance, and the complete menu inventory are not yet verified.

Edit → Delete now operates on the selected text during direct text editing, using the rich field's saved selection after the menu takes focus. It emits the same beforeinput/input sequence used by existing edit tracking, so local Undo restores text and selection and Redo repeats the deletion. A collapsed caret disables the menu command and produces no change through the field method. Object deletion keeps its existing behavior when no text session is active.

Validation: the application browser regression covers partial text deletion, Undo with restored selection, Redo, disabled caret-only Delete, and text Select All through Edit. A separate rich-text browser case covers saved selection across formatted paragraphs, correct input events and no-op caret deletion. These and both existing rich-text cases pass (four browser tests across the two suites). Dev build, TypeScript, lint, and whitespace checks pass. Native Mac menu naming/order and full clipboard/Edit behavior remain unverified or incomplete as previously recorded.

Presentation text search is now available in the upper-right title bar and through Edit → Find. Command/Ctrl+F focuses the search field; Enter/Shift+Enter, the previous/next buttons, and Command/Ctrl+G with optional Shift navigate results. Search treats queries literally, ignores case with Unicode-aware regex matching, and retains UTF-16 offsets for direct text selection. It starts on the current slide, traverses top-level editable shapes in model order, wraps in either direction, and reports the match index or no matches. Before navigation it commits active text edits; failed saves stop navigation. Search results open the exact text range for continued editing.

Browser validation covers mixed-case repeated matches, navigation across two slides in both directions, wraparound, direct editing and save before further searching, literal bracket queries, no matches, Escape, Edit submenu access, and nonoverlapping title/search layout. The existing Page Setup/menu regression also passes. Dev build, TypeScript, lint and whitespace checks pass. Screenshot inspection revealed and corrected title overlap.

This does not complete Find/Replace parity: replacement commands/options, grouped child text, notes/master/table/chart text, native search ordering, selection restoration details and exact Mac layout remain outstanding. The search-box entry point follows the macOS section of Microsoft support: https://support.microsoft.com/en-us/powerpoint/find-and-replace-text . Native reference capture remains unavailable as previously recorded.

Edit → Find → Replace now opens a modeless Find and Replace dialog with Find what, Replace with, Find Next, Replace and Close. The dialog preserves active text selection when taking focus. Replace only mutates a currently selected matching range; otherwise it locates the next result. It uses the existing rich-text input/edit journal path, commits before advancing, stops on save failure, and disables repeated actions while pending. After a successful replacement, navigation starts after the inserted text, avoiding an immediate repeat when the replacement contains the search term; wraparound remains available.

Validation: the search browser case covers ordinary replacement, an inserted string containing the search term, movement to the later original occurrence, empty replacement, and undoing each saved replacement. All three rich-text browser cases also pass. Final dev build, TypeScript, lint and whitespace checks pass; the modeless dialog screenshot was inspected. Replace All, match options, mixed-format replacement-specific coverage, additional text containers, and exact native dialog layout/interaction remain pending. This increment is not full Mac Find/Replace parity.

Find and Replace now includes Replace All for the current searchable scope (top-level editable slide text). It captures original literal, case-insensitive matches and applies them backwards, preserving rich formatting and preventing inserted search terms from being replaced recursively. All slides are recorded in one history entry; the replay fingerprint covers every searched slide and rejects conflicting source edits. The dialog reports the replacement count and guards repeated submissions while saving.

Validation: 34 editor tests and the search browser regression pass, including mixed-format replacement, cross-slide replay/source conflicts, replacement containing the query, and one-step Undo restoring both slides. Dev build, typecheck, lint and whitespace checks pass. The updated dialog screenshot was inspected. Search options, other text containers, replacement-specific autofit recalculation and exact native Mac UI/behavior remain outstanding; this is not full parity.

Search/replace now exposes Match case and Find whole words only. These options apply consistently to Find Next/Previous, single replacement validation, Replace All and journal replay. Changing options clears the prior result position. Queries/options/replacement fields are disabled while a dialog operation is pending. Whole-word boundaries currently treat Unicode letters, numbers, combining marks and connector punctuation as word characters; exact locale-dependent Mac segmentation is not yet verified. A title-bar search menu now opens Find and Replace, matching the entry point documented for macOS at https://support.microsoft.com/en-us/powerpoint/find-and-replace-text . The option semantics/defaults are documented by https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textrange.replace ; that API reference does not verify the native dialog's precise option layout.

Validation: 37 backend/search tests and one browser integration pass. Coverage includes case-sensitive whole-word replacement and replay, literal punctuation, Unicode/astral boundaries, option switching, count feedback and the title-bar menu entry. Typecheck, lint, dev build and whitespace checks pass; the updated dialog screenshot was inspected. Native capture was retried and failed with ScreenCaptureKit -3811; the reference document was not changed. Exact native UI and locale matching, additional text containers and replacement autofit remain pending.

Replace All now recomputes normal and shape autofit for every matched text shape across all slides. A revision-checked preview supplies the post-replacement rich text model; the existing browser measurement path calculates font scales or bounds. Measurements carry slide identity because shape IDs can repeat across slides, and text plus fit results are persisted in the same journal entry. Fit targets, duplicate measurements, modes, bounds and scales are validated before replacement mutates text. Attached connectors refresh after the combined operation.

Validation: 36 editor tests and two search browser tests pass, including both fit modes on separate slides with repeating shape IDs, preservation of fixed bounds for shrink-to-fit, growth for resize-to-text, cross-slide one-step Undo/Redo, PPTX round-trip, replay and invalid-target rejection without mutation. Typecheck, lint, dev build and whitespace checks pass. Exact native font metrics, preset text rectangles, additional searchable containers and full native parity remain outstanding.

The existing full canvas-editor browser regression also passes after the shared measurement change (39 tests across the backend and three browser cases in this increment), covering ordinary text/format/autofit editing, history and reload persistence.

Edit now exposes Cut, Copy and Paste alongside the existing commands. During direct text editing, the menu and ribbon use the browser text clipboard and the saved rich-field selection. Cut writes successfully before deleting; Paste inserts plain text through existing edit tracking, retaining local Undo. Empty clipboard text leaves the selection intact. Clipboard failures surface an error without deleting selected text, and asynchronous operations check that the session, value and selection have not changed before applying destructive edits. Pending clipboard actions are disabled.

Validation: the new clipboard browser case and existing Page Setup/menu regression pass (two tests), covering partial Copy/Cut/Paste, multiline insertion, Undo, empty clipboard, denied writes, caret-only disabled commands, ribbon entry points and save. Typecheck, lint, dev build and whitespace checks pass. The test queries the rich editing field through its shadow root using Playwright locators. Native Mac menu placement/appearance remains unverified. Rich clipboard formats, OS object clipboard, immutable object snapshots and immediate object Cut remain outstanding; object commands currently reuse the existing internal clipboard behavior.

Direct-text right-click now offers text Cut/Copy/Paste/Delete/Select All and restores focus to the editing field. It no longer shows object Duplicate/Delete while editing text. Clipboard enablement follows the text selection and browser clipboard availability. Regression coverage also explicitly delays Cut and Paste promises, changes the selected range, and verifies that stale results leave both text and the new selection intact and re-enable commands. The text context Delete path is checked with local Undo. Clipboard browser regression, TypeScript, lint and whitespace checks pass; the existing Page Setup regression also passes. Exact native context-menu inventory and typography remain unverified.

Shape copy now clones referenced chart parts and their owned dependency graphs, including embedded workbooks, instead of sharing editable chart state with the original. Copies within one slide and across slides remain independently editable after PPTX save/reload. Shape relationship targets are resolved against the source slide before transfer, retaining valid references when paths differ. Image/media sharing remains unchanged. Other editable object dependency types, immutable clipboard snapshots and immediate object Cut remain pending.

Validation: seven copy/dependency tests plus 45 grouping/media/group-text regression tests pass (52 tests). Root typecheck uncovered an existing optional `anchorCenter` assignment; it now supplies the existing false default explicitly and typecheck passes. Lint, formatting, whitespace checks, root build and dev build pass. This increment fixes shared chart state in the copy primitive; it does not yet remove the UI clipboard's source-token restriction or establish full Mac parity.

Added the core `importShape` primitive needed to paste from a frozen presentation snapshot. It transfers explicit shape XML and referenced package dependency graphs into the destination package, allocates noncolliding part names, preserves external relationships, and leaves source bytes untouched. The existing `copyShape` same-package contract is preserved. Cross-package imports clone image and theme dependencies too, rather than assuming those parts exist in the destination. Inherited slide layout/theme formatting still follows the destination, and native hyperlink-to-slide semantics remain unverified.

Validation: nine copy/dependency tests pass, including image bytes across packages and repeated imports from a saved snapshot after editing both the original and a prior import. The resulting chart values remain independent after PPTX save/reload and presentation validation passes. Root typecheck, lint, formatting, whitespace checks, root build and dev build pass. This is the transfer primitive only: the UI clipboard and history do not yet store immutable snapshots, and full Mac parity remains incomplete.

The object clipboard now captures revision-checked PPTX bytes at Copy/Cut. Paste stores the immutable snapshot in its journal command and imports selected shapes from it, so later source edits/deletion and changes to earlier pasted copies no longer invalidate the clipboard. Cut deletes immediately after a successful snapshot capture with the original revision guard; Undo restores it, and the clipboard remains available for later pastes. Legacy live-source paste journal entries still replay through the existing path. Snapshot paste never deletes source objects and rejects the legacy cut flag. `loadPresentationBytes` provides synchronous byte loading for deterministic journal replay.

Validation: 37 backend tests and two clipboard browser tests pass, covering source deletion, changed originals, repeated pastes, JSON journal round-trip, undo cursor replay, immediate Cut, Undo of Cut/Paste, reload persistence and existing text clipboard operations. Root/dev typecheck, lint, root/dev builds and whitespace checks pass. Remaining limits: snapshots currently include the entire deck per paste (26-million-character base64 cap); compaction/deduplication is pending. This remains an internal clipboard rather than the OS object clipboard. Explicit XML and dependencies are retained, but inherited layout/theme formatting follows the destination; native paste options, theme semantics and exact Mac UI remain incomplete.

The full canvas/history/notes/reload browser regression and Page Setup/menu regression also pass after snapshot integration (41 tests total across the backend and four browser cases in this increment).

Snapshot paste now restores selected objects in source stacking order, independent of the order in which they were selected. Legacy live-source journal entries retain their prior ordering for deterministic replay. Source IDs are resolved before any copy is appended. Regression coverage pastes a group and a connector selected in reverse order into another slide after deleting the originals, verifies unique IDs and attachment to the copied group child, and checks saved PPTX and JSON journal replay. Pasting the connector alone clears its original attachment instead of binding to an unrelated destination object with a matching ID.

Validation: all 38 editor backend tests, dev TypeScript, lint, formatting, whitespace checks and dev build pass. This validates the internal copy behavior; native Mac visual comparison remains blocked by ScreenCaptureKit capture failure and full UI/operation parity remains incomplete.

Edit completion now checks the destination slide identity before updating selection or restoring editing focus. If the user navigates during a pending operation, the saved edit stays on its original slide without selecting unrelated objects or stealing focus on the newly displayed slide. A delayed-Paste browser regression exercises navigation while the request is pending.

Validation for the navigation guard: the object clipboard browser case and full canvas/history/notes/reload regression pass (two tests). The clipboard case confirms that the pending paste changes only its original slide, keeps the newly selected thumbnail focused, and leaves the destination slide's objects unselected and unchanged. An initial test click targeted an obscured back object after reload; it was corrected to click the front object. Dev TypeScript, lint, build and whitespace checks pass. Native UI capture was retried and still returns ScreenCaptureKit -3811; native equivalence remains unverified.


Insert now includes a Table menu with a 10-column by 8-row grid, arrow-key navigation, dimension feedback and Escape cancellation. An Insert Table dialog accepts integer dimensions from 1 to 75. Insertion creates an editable PPTX table with the existing theme style, header row and banding, and uses the shared selection, history, save and replay paths. Backend validation rejects invalid dimensions before mutation.

The Mac section of [Microsoft’s screen-reader instructions](https://support.microsoft.com/en-us/accessibility/powerpoint/use-a-screen-reader-to-insert-and-edit-pictures-and-tables-in-powerpoint) confirms the Insert → Table menu and keyboard dimension selection. Exact Mac grid dimensions, numeric dialog defaults/limits, insertion geometry and visual styling remain unverified; the current values are implementation choices. Cell editing, row/column operations and table contextual ribbons remain outstanding.

Validation: all 39 backend tests and the new table browser integration pass. Coverage includes dimension validation, keyboard selection and cancellation, numeric form validation, selection after insertion, Undo/Redo, journal replay and PPTX save/reload. Dev build, TypeScript, lint and whitespace checks pass. Grid, dialog and saved-table screenshots were inspected for layout. This increment does not establish full Mac UI or operation parity.

Table cell editing now has a range-preserving core primitive, `replaceTableCellTextRange`, sharing the existing DrawingML range replacement implementation with shapes. It retains surrounding runs, paragraph properties and untouched XML rather than flattening the whole cell. The editor model now includes row heights, column widths and cell text, spans, literal paragraph formats, margins, anchor and direction. Theme-derived cell formatting is not resolved in this model yet.

The `table-cell-text` journal command targets one cell and validates the entire sequential edit list, including UTF-16 surrogate boundaries, before mutation. Covered portions of merged cells are rejected; editing targets the merge anchor. The command participates in ordinary history replay and PPTX save/reload. This is the data and persistence foundation: direct cell selection, on-canvas typing, Tab navigation and contextual ribbon wiring are still pending, so table editing UI parity is not yet achieved.

Validation for the cell-editing foundation: eight core text-range tests, 40 editor backend tests and the table insertion browser regression pass (49 tests). Coverage includes rich runs and paragraph-end formatting across multiline replacement, Japanese text, surrogate-boundary rejection without mutation, merge-anchor targeting, sequential edit validation, undo/replay and PPTX round-trip. Root/dev TypeScript, root/dev builds, lint, formatting and whitespace checks pass. Native visual comparison and direct cell editing remain unverified and unfinished.


Table cells can now enter the shared on-canvas rich text editor through double-click, preserving sequential range edits in the table-cell journal command. Tab and Shift+Tab commit the current cell and move among visible cells, skipping merged continuations. Cell overlay bounds account for unequal row/column sizes, merged spans, table scaling and rotation about the table center. Text clipboard and local text Undo reuse the existing editing session. Font/paragraph ribbon commands remain unavailable for tables; unsupported character shortcuts are prevented from applying browser-only formatting.

Remaining limits: click-to-place-caret and single-click entry, last-cell Tab row insertion, table-style/theme formatting in the editing overlay, full paragraph-property inheritance, row autofit, multi-cell selection and table contextual ribbons. Cell editing currently uses the existing select-all entry behavior and white editing background. Exact Mac appearance and interaction equivalence remain unverified.

Validation for direct cell editing: the table browser case covers Japanese/emoji entry, forward/backward cell navigation, Undo/Redo and exported PPTX text after reload. The existing full canvas/history/notes/reload regression also passes. A geometry unit test verifies unequal dimensions, merged anchors/covered cells and rotation. Dev TypeScript, build, lint, formatting and whitespace checks pass. Saved and active-cell screenshots were inspected; the active-cell theme/background mismatch is recorded above.

Cell editing overlays now reuse the preview renderer's resolved cell background and default text color/font through per-cell SVG metadata. Explicit run formatting remains higher priority. Transparent cells use the renderer's white table backdrop; merged cells use their anchor's appearance. This removes the white rectangle previously covering styled header cells while typing. The renderer's existing table-style approximation is unchanged, and this does not establish native table-theme fidelity or resolve scheme colors in individual runs.

Validation: the cell projection test verifies inherited defaults and explicit run overrides; the table browser test compares computed header and editing backgrounds and still covers Japanese/emoji text, Tab navigation, Undo/Redo and PPTX reload. Both pass. Preview/dev builds, dev TypeScript, lint and whitespace checks pass. The active-cell screenshot was inspected and retains the blue header background. Native Mac UI parity, cell formatting controls and contextual ribbons remain unfinished.

Cell text sessions now enable character-format ribbon controls (font, size, color, bold/italic/underline, strike, baseline, spacing and highlight). Selected ranges use the `table-cell-format` journal command; collapsed selections set insertion formatting for subsequent text. Formatting reopens the same cell and restores the selected UTF-16 range. Command/Ctrl+B/I/U uses this path instead of being suppressed. Paragraph, text-frame and whole-table formatting controls remain unavailable; multi-cell selection and native contextual ribbons are still outstanding.

`setTableCellTextRangeFormat` shares the range-formatting primitive with shape text. It splits literal runs without flattening unaffected formatting or surrounding XML; field runs retain the existing atomic formatting behavior. Cell text journal entries now persist optional insertion formats, with sequence/range/format validation before text mutation.

Validation: nine core range tests, 41 editor backend tests and two browser tests pass (52 total). Coverage includes partial-cell formatting without affecting neighboring text/cells, invalid range/format rejection, Japanese text, emoji boundaries, caret insertion formatting, retained selection after ribbon formatting, JSON history replay/undo, saved PPTX reload and the existing full canvas regression. Root/dev typecheck, builds, lint, formatting and whitespace checks pass. These checks prove this implementation's behavior, not exact native Mac appearance or complete operation parity.

Cell sessions now enable horizontal paragraph alignment controls: left, center, right, justified and distributed. `setTableCellTextRangeAlignment` targets the paragraphs touched by a half-open UTF-16 range, or the caret's paragraph for an empty selection. It preserves surrounding paragraphs and run XML. The `table-cell-align` journal command persists this operation, and the editor restores its cell selection afterward. Distributed alignment in the table's browser preview now justifies the last line as well.

Validation: 10 core range tests, 42 editor backend tests and two browser cases pass (54 total). Coverage includes paragraph boundaries, empty paragraphs, embedded line breaks, surrogate-range rejection, partial paragraph selection, untouched adjacent cells, JSON replay, Undo/Redo, exported PPTX reload and existing canvas behavior. Root/dev TypeScript, root/preview/dev builds, lint, formatting and whitespace checks pass. The first browser attempt exposed editing entry during a pending rebuild; direct text entry now waits for the build to be idle, and the test waits for readiness after reload. Native visual comparison, paragraph spacing/bullets in cells, multi-cell selection and table contextual ribbons remain incomplete.


Cell sessions now enable the existing line-spacing ribbon menu (1, 1.5, 2, 2.5 and 3). The `table-cell-line-spacing` journal command applies the value only to paragraphs touched by the current selection, restores that selection after saving, and supports history replay. `setTableCellTextRangeLineSpacing` also supports fixed point spacing and clearing the explicit setting, validates before mutation, and retains surrounding paragraph/run XML. Cell paragraph reading now exposes literal paragraph properties; the editing projection and both preview text paths consume authored line spacing and before/after paragraph spacing. Before/after spacing has no new cell UI yet. Bullet/indent rendering and controls, inherited paragraph defaults, multi-cell selection and contextual ribbons remain incomplete.

Validation: four core table-range tests, 43 editor backend tests, one cell projection test and the table browser regression pass (49 tests). The added checks cover surrogate/range rejection without mutation, percentage and fixed-point spacing, clearing, selection boundaries, unchanged adjacent paragraphs, selection restoration, Undo/Redo, JSON replay and exported PPTX reload. Root/dev TypeScript, root/preview/dev builds, lint and whitespace checks pass. Native capture still fails with ScreenCaptureKit -3811; no reference document edits were made. Exact Mac UI/behavior parity remains unverified and unfinished.


Cell editing now enables the existing bullets and numbering buttons. The `table-cell-bullets` command applies bullet, number or explicit none only to paragraphs touched by the selection, preserving the selection after save. Cell projections consume literal bullet/indent/level properties. Both table preview paths render ordinary character and automatic-number bullets with authored hanging indents; numbering counters are shared with shape rendering. Picture bullet relationships, explicit bullet font/size overrides, numbering start/restart controls and full inherited defaults in cells remain incomplete. This is not a claim of native ribbon or interaction parity.

The shared bullet mutator now removes superseded image bullet and font-follow-text elements, and inserts bullet font/marker before tab stops, default run properties and extensions. This avoids conflicting bullet choices and incorrect DrawingML child order when editing imported paragraphs.

Validation: 30 core/preview tests, 44 editor backend tests, one cell projection test and one browser regression pass (76 total). Coverage includes selected-paragraph numbering in both preview modes, unchanged adjacent paragraphs/cells, invalid range/style rejection before mutation, marker replacement and XML ordering, selection restoration, bullet/number/off toggling, Undo/Redo, journal replay and exported PPTX reload. Root/dev TypeScript, root/preview/dev builds, lint and whitespace checks pass. Full Mac parity remains incomplete, including multi-cell selection, table contextual ribbons, inherited/table-style typography, cell row autofit and direct comparison with the native app.


Tab in the last visible cell now commits text, appends a row through the `table-row-append` history command, and opens its first cell for typing. Shift+Tab navigates backward without inserting. The appended row uses the last row's declared height, while the table's outer height grows proportionally to preserve existing rendered row heights. Bounds compensate for rotation so the top edge stays fixed. Completion checks both slide index and stable slide key before moving focus, preventing a pending save from pulling editing onto another slide.

The last-cell Tab trigger is documented by Microsoft for Mac: https://support.microsoft.com/en-us/office/add-or-delete-rows-or-columns-in-a-table-in-word-or-powerpoint-for-mac . Native capture was retried and still failed with ScreenCaptureKit -3811; no reference document was modified. Exact row-height choice, direct cell formatting inheritance, merged-row insertion behavior and native undo grouping remain unverified. New rows currently use the core blank-cell defaults and the table's existing style flags, without copying direct formatting or merge spans from the previous row. The command has an implementation limit of 10,000 rows; this is not a claimed PowerPoint limit.

Validation: 45 backend tests and the table browser case pass (46 total). The added checks cover unequal row heights in a rotated table, unchanged existing text, invalid-target rejection, JSON replay, undo, last-cell typing, Tab insertion, immediate new-cell typing, backward navigation, separate text/row Undo/Redo and PPTX reload. Dev TypeScript/build, lint, formatting and whitespace checks pass. Full Mac parity remains incomplete, including the row-formatting and merge details above, multi-cell selection, contextual ribbons and other outstanding operations.

### 表セルの文字方向と上下配置（2026-09-23）

- 既存 Home の Text Direction / Align Text をセル編集中にも有効化。Horizontal / Rotate all text 90° / Rotate all text 270° / Stacked と Top / Middle / Bottom を対象セルに保存し、キャレット・選択範囲を復元して編集を続けられる。
- `table-cell-layout` は変更対象が表の有効なセル・結合アンカーか確認し、方向と配置の両方を検証してから変更する。隣接セルの文字・書式や表の外形は変更しない。履歴の再生・Undo・PPTX 再読み込みでも反映する。
- プレビューの表セル描画が `a:tcPr/@vert` を参照するよう修正。SVG は既存の縦書きレイアウトエンジン、foreignObject は writing-mode / text-orientation / 270 度方向の反転を使用する。編集中の既存 rich-text フィールドも同じセル属性を利用する。
- 確認: core/preview 17 件、editor backend 46 件、browser table 1 件、合計 64 件成功。Dev / Preview 型チェック、両パッケージ build、対象 lint、diff check 成功。スクリーンショット `/tmp/pptx-table-cell-direction.png` を確認（短い行高での 270 度方向の折返しも表示）。
- Mac PowerPoint の画面取得を再試行したが ScreenCaptureKit -3811 で失敗。参考に Microsoft Support の「Set text direction and position in a shape or text box in PowerPoint」「Change the look of a table」を確認したが、現行 Mac の表専用リボン位置・ラベル・メニュー構成の完全一致は未検証。既存 Home メニューの有効化であり、Table Layout リボン追加や行・列操作 UI の実装完了を意味しない。
- 残件: 複数セル選択、セル・行の autofit、未指定 margin の preview / editor 差、SVG と CSS の縦書きの厳密な配置差（特に非対称余白・東アジア文字・オーバーフロー）、ネイティブ実機比較。全操作・UI の完全一致は引き続き未完了。

### Contextual Table Layout ribbon

- Added a Table Layout tab visible only for a single selected table. It exposes the existing cell paragraph alignment, vertical anchoring, text direction, and Arrange commands. Cell text selection survives ribbon/menu operations.
- Ribbon arrow-key navigation skips hidden contextual tabs. Removing the selected table while its tab is active returns the ribbon to Home; other user-selected tabs remain selected when tables appear.
- Verified the browser table workflow including contextual visibility, hidden-tab keyboard navigation, Undo fallback, cell layout edits, selection preservation, and saved/reloaded OOXML. Dev build, TypeScript, targeted lint, and diff checks passed. Inspected `/tmp/pptx-table-cell-direction.png`.
- This is a partial contextual ribbon, not complete native parity: Table Design, row/column commands, merge/split, size/distribution, and cell margins are still missing. Exact Mac arrangement/icons and selection-triggered default-tab behavior remain unverified because native capture is unavailable. Existing short-row clipping remains.

### Table Design style options

- Added contextual Table Design before Table Layout, with Header Row, Total Row, Banded Rows, First Column, Last Column, and Banded Columns checkboxes bound to the table's OOXML style flags.
- The table-style command validates the whole supplied options object before mutation, preserves omitted flags and cell content, and supports journal replay, Undo/Redo, and PPTX export/reload. UI state retains the pending checkbox value during saving and restores the active cell's text selection afterward.
- Validation: 47 backend tests and the browser table workflow pass; dev TypeScript, build, targeted lint, and diff checks pass. Inspected `/tmp/pptx-table-design.png`.
- Partial parity remains: style gallery, shading, borders, effects, and exact Mac visual comparison are not complete. Flag rendering still uses the existing approximate table-style renderer, not full theme/table-style inheritance. Row autofit and clipping remain outstanding.

### Table shading

- Table Design now exposes Shading with solid standard colors and No Fill. While editing a cell it changes that anchor cell; selecting the table frame applies the fill to all cells. Active text ranges are restored after saving.
- `setTableCellFill(cell, null)` writes explicit `a:noFill`; `isTableCellNoFill` distinguishes this from inherited shading. Preview rendering respects that override. `clearTableCellFill` continues to remove the override and restore table-style inheritance.
- Validated cell/whole-table targeting, unchanged neighboring cells/text, invalid-input rejection, serialization, journal replay and Undo in backend tests. Browser validation covers color/transparent rendering, text selection preservation and exported no-fill state. 48 backend + 17 core/preview + 1 browser tests pass; root/dev/preview type checks, all three builds, targeted lint and diff checks pass.
- The Shading menu is still partial: native theme-color grid, recent/custom colors, eyedropper, and image/gradient fills remain unimplemented. Multi-cell selection and exact native Mac appearance still need work; this does not establish full parity.

### Table borders

- Table Design now offers outside/inside/all/no borders, individual sides, inside horizontal/vertical borders, and both diagonals. Pen controls select color, 0.25–6 pt width, and solid/dash/dot/dash-dot lines. Commands target the active cell or the whole table and retain the active text selection.
- Shared edges update on both neighboring cells. Merged perimeters retain independent physical segments, so changing a neighbor beside the lower half of a merged cell leaves its upper edge segment unchanged. Explicit `a:noFill` suppresses a border independently of removing its override. Preview draws merged perimeter segments and honors no-fill; materialized empty border properties retain the default grid appearance.
- Validation covers malformed input, merged anchors, partial merged edges, serialization, journal replay, Undo, browser ribbon commands, saved pen settings and no-border state. Native theme inheritance, compound lines, drawn borders/eraser, multi-cell selection, menu icons and exact Mac appearance remain incomplete. Inspected `/tmp/pptx-table-design.png`; corrected a clipped Pen Color control by moving it beside the line selectors.
- Checks pass: 49 backend + 20 core/preview + 1 browser tests (70 total), root/dev/preview type checks and builds, targeted lint and whitespace checks. Full parity is still incomplete.

### Table row/column distribution

- Added Distribute Rows and Distribute Columns to Table Layout. A selected table frame distributes its complete grid; an active cell distributes only the grid rows/columns spanned by that cell. Text selection is restored after the command.
- Grid totals remain exact (rounded boundaries differ by at most one EMU), while table bounds, rotation, cell content and merge spans remain unchanged. Invalid axes, covered cells and invalid coordinates are rejected before mutation.
- Backend coverage uses unequal rows/columns and a rotated table with a merged cell; checks untouched columns, exact totals, saved/reloaded dimensions, journal replay and Undo. 50 backend tests and the browser table workflow pass. Browser coverage exercises both buttons while preserving a cell text range; dev type check/build, targeted lint and whitespace checks pass. Inspected `/tmp/pptx-table-distribution.png`.
- Microsoft documents distribution on the Mac Table Layout tab: https://support.microsoft.com/en-us/word/resize-a-table-in-word-or-powerpoint-for-mac . Exact native behavior for a single merged-cell selection remains unverified. Arbitrary multi-cell/range selection, row-height/column-width input, text-driven minimum row heights and native icon/layout matching remain outstanding; full parity is incomplete.

### Numeric table cell dimensions

- Table Layout now has Height and Width numeric controls in inches. A table-frame selection applies the size to every row/column; an active cell resizes its row/column span, proportionally retaining the internal grid ratios of a merged cell. Mixed frame-selection values display an empty field with a Mixed placeholder.
- Commands convert authored grid dimensions to displayed dimensions before resizing, preserve neighboring row/column sizes, grow the frame, and compensate its origin for rotation. Input and resulting bounds are validated before mutation. Active cell text ranges are restored after committing.
- 51 backend tests plus the browser table workflow pass. Tests include merged-column targeting, unchanged neighbor width, rotated-origin stability, invalid values, PPTX reload, journal replay/Undo and whole-table row sizing. Browser coverage sets the first cell height to 0.5 inches, checks the saved row height and retained text range. Dev type check/build, targeted lint and diff checks pass; inspected `/tmp/pptx-table-distribution.png`.
- Native minimum row height/autofit, arbitrary cell-range selection, unit/localization preferences, spin-control appearance, and exact merged-span behavior remain unverified/incomplete. Controls currently accept 0.001–56 inches as an implementation range, not a claimed PowerPoint limit. Full Mac operation/UI parity remains incomplete.
- Visual review caught dimension controls using the synthetic cell editing bounds as the table extent. Controls now use the selected table's real frame; the browser test additionally asserts the 0.5-inch input readout after saving.

### Table cell margins

- Table Layout now opens Cell Margins with Top, Bottom, Left and Right numeric inputs in inches. Editing a cell targets that cell; selecting the table frame targets all cells. Mixed values remain blank, and unchanged fields preserve the original EMU precision. Cancel and Escape restore the active text range without applying changes.
- The command validates all sides and coordinates before mutation, preserves omitted margins and neighboring cells, and supports PPTX reload, journal replay and Undo. The current accepted range of 0–56 inches is an implementation constraint, not a verified PowerPoint limit.
- Backend coverage checks partial updates, invalid values, whole-table application, serialization, replay and Undo. Browser coverage checks Cancel, retained text selection, applied margins and exported neighboring-cell values. Native margin presets, arbitrary cell-range selection, unit localization and exact Mac dialog/menu appearance remain unverified; full parity is incomplete. Native capture still fails with ScreenCaptureKit -3811. Inspected `/tmp/pptx-table-margins.png`.
- Validation status for this increment: corrected targeted margins backend test passes; the preceding full run passed the other 51 tests (the margins assertion incorrectly expected absent neighbor margins instead of the authored 91440 EMU default). Dev build/type check, targeted lint and whitespace checks pass. Browser reruns are not green: one timed out reopening a cell, and the latest hit its 120-second total limit after switching that operation to the measured cell center. The fresh full backend run was terminated when subsequent row/column edits superseded its source state; all 10 table-related backend tests pass after those edits. The focused browser workflow added below now passes the margins checks, while native parity remains unverified.

### Table row/column insertion controls

- Added Insert Above, Insert Below, Insert Left and Insert Right in Table Layout's Rows & Columns group. Active-cell insertion uses the outer boundary of the cell's merged span; a frame selection inserts at the corresponding table edge. Successful commands focus the new cell after committing any active text edit.
- Insertions retain displayed sizes of existing rows/columns, copy the adjacent grid size for the new row/column, expand the frame and compensate the origin for rotation. Core row/column insertion now extends merges crossed by the inserted grid line and writes the new covered-cell markers; boundary insertions preserve existing spans and text.
- Command coverage checks all four placements, rotated bounds, crossed merges, invalid selection/placement, PPTX reload, replay and Undo. Core tests cover each insertion boundary and interior for a 2×2 merge with reload. Microsoft documents these four controls on Mac Table Layout: https://support.microsoft.com/en-us/office/add-or-delete-rows-or-columns-in-a-table-in-word-or-powerpoint-for-mac . Exact native merged-span insertion behavior, inherited direct formatting, multi-cell/range insertion and native visual layout still require comparison. Delete rows/columns and cell selection remain incomplete; this is not full parity.
- Validation: 12 core row/column tests and 10 table-related backend tests pass. A focused browser workflow passes all four insertion controls, pending-text commit, new-cell focus, margin Cancel/Apply, text selection restoration, Undo/Redo and exported/reloaded cell values. Root/dev builds, both TypeScript checks, targeted lint and whitespace checks pass. Inspected `/tmp/pptx-table-structure.png`.
- The focused browser test waits for the inserted table's selection UI before beginning a direct edit; model/build completion alone precedes the command's selection update. Measured cell click targets are scoped to `#stage` to avoid thumbnail SVG cells. Previous broad browser coverage remains separately tracked; its earlier timeout is not a failure of the now-passing focused margins workflow.
- The broad `table-insert.test.mjs` rerun still fails separately: it times out waiting for `!` to persist after the existing rich-text italic/insertion sequence (line 186 at this revision), before the new margins/insertion sections. The focused structure/margins browser test passes. Investigate that rich-text sequence independently; do not report the broad browser suite as green.


### Save completion and continuous table editing

- Fixed a save/watch race: edits previously acknowledged an immediately published revision before a delayed history-file notification initiated replay. The next user action could hit a rebuilding document, including after Undo or before reopening a cell. Saves and Undo/Redo now explicitly await replay; the watcher ignores notifications whose history matches the in-flight save or current history. External history changes still trigger rebuilding.
- Added a focused browser regression for reopening a saved Japanese/emoji cell, applying bold to a range, toggling italic at the caret, immediately typing, and checking persisted text and character formats. It also changes the journal externally and verifies that the preview updates.
- The previously failing broad table browser workflow now passes, including its caret-format/input sequence. The focused structure/margins workflow also passes: three distinct browser tests pass after this fix. Dev build/type check, targeted lint/formatting and whitespace checks pass. Inspected the latest `/tmp/pptx-table-structure.png`. This resolves the earlier recorded browser failures; it does not establish complete Mac UI or operation parity.

### Table deletion controls

- Table Layout now has a Delete menu with Delete Columns, Delete Rows and Delete Table, following Microsoft's documented command path: https://support.microsoft.com/en-us/powerpoint/add-or-delete-table-rows-and-columns . Active-cell row/column deletion targets that cell's merged span; frame selection targets the whole grid. Pending text commits before deletion, and a surviving table resumes editing in the adjacent cell/merged anchor. Deleting the final row or column removes the table.
- Commands normalize grid sizes to displayed dimensions, retain surviving sizes, shrink the frame and preserve its rotated top-left position. Core removal repairs intersected merges, shrinks their spans, and promotes the existing anchor with its text/formatting when its physical row/column is removed. Invalid noninteger/out-of-range indices are rejected before mutation.
- Validation: 16 core grid tests, one backend test covering both axes, merged-span/crossing-merge deletion, rotated bounds, invalid commands, PPTX reload, replay and Undo, and one browser workflow covering all three menu actions, pending text, resumed editing, Undo/Redo and saved text pass. Root/dev builds and type checks, targeted lint/formatting and whitespace checks pass. Inspected `/tmp/pptx-table-delete.png`.
- Native capture was retried and still fails with ScreenCaptureKit -3811. Exact native anchor-content handling, merged-span deletion selection, post-delete caret behavior and menu/icon appearance therefore remain unverified. Arbitrary multi-cell selection and its deletion semantics remain incomplete. Full Mac UI/operation parity is not achieved.

### Editing font fallback

- Text fields, formatted runs and placeholder prompts now use the same Calibri/Helvetica Neue/Arial/sans-serif fallback sequence as the slide renderer. Previously a missing document font could fall back to the browser's serif default only during editing.
- The table deletion browser workflow verifies fallback on both the field and typed runs and passes, including deletion, resumed editing, persisted text and Undo/Redo. Inspected `/tmp/pptx-table-delete.png`: the entered text now displays sans-serif. Dev build/typecheck, targeted formatting/lint and whitespace checks pass. This aligns preview/editing fallback; it does not verify native Mac font substitution or complete UI parity.

### Table merge command groundwork

- Added an opt-in `coveredText: 'append'` core merge mode: nonempty cells contribute their original paragraph XML in row-major order to the anchor, preserving run and paragraph properties; covered text bodies are removed. Existing keep/drop modes retain their behavior. All range/overlap validation precedes mutation.
- Added a `table-merge` editor command with an explicit rectangular range, using append mode and the existing edit journal. Core merge tests (17) and a backend reload/replay/Undo test pass, as do root/dev builds, both typechecks and targeted lint/format/whitespace checks.
- Microsoft's operation reference is https://support.microsoft.com/en-us/powerpoint/merge-or-split-table-cells-in-powerpoint . This is backend groundwork, not a completed user-facing merge/split workflow. Range selection, ribbon commands, arbitrary subdivision and existing-merge handling remain to implement. Native empty-cell/paragraph ordering details remain unverified.

### Table range selection and Merge Cells UI

- Dragging inside an already selected table now selects a rectangular cell range, highlights it, and expands intersections to whole merged cells. The outer five-pixel frame and handles retain table move/resize behavior; a click inside starts cell editing. Selection is cleared when slide, revision or shape selection changes. Table Layout exposes Merge Cells for multi-cell ranges without pre-existing merges and resumes editing in the resulting anchor.
- Delete Rows/Columns now accepts the selected rectangle, validates its bounds and deletes only its rows/columns. Browser coverage passes for drag selection, ribbon merge, editing/saving the merged result, unchanged frame bounds, two Undo steps, and deleting only selected columns. The existing table-deletion/editing browser workflow also passes. Dev build/typecheck, targeted lint/format/whitespace checks pass. Inspected `/tmp/pptx-table-merge.png`.
- One range-deletion test initially dragged before Undo had re-enabled controls; it now waits for the enabled Delete control and asserts the range before deletion, and passes. Full native parity remains unverified. Range-aware formatting, keyboard clear/copy/paste, selection from active text editing, existing-merge consolidation, split cells and precise native selection visuals still require work.

### Clearing a selected cell range

- Delete/Backspace and the selected table's context-menu Delete now clear the selected cell range's text while retaining the table, merges, cell properties and grid dimensions. Successful clearing retains the range highlight. Explicit Table Layout > Delete Table still removes the table.
- The command validates the complete range before mutation and rejects ranges cutting through merged cells. Empty text nodes now serialize canonically, avoiding slide-token changes caused solely by XML parse/save normalization.
- Validation: the backend regression passes invalid-range atomicity, retained fill/merges/dimensions, save/reload, history replay and Undo. The browser merge workflow passes range clearing, retained selection and Undo restoration. Thirteen core text-range tests pass; root/dev builds and targeted lint/format/whitespace checks pass.
- Exact Mac key/menu and selection appearance remain unverified because native capture has failed with ScreenCaptureKit -3811. Range-aware formatting/copy/paste, selection from active text editing, consolidation of existing merges and Split Cells remain incomplete; this increment does not establish full PowerPoint parity.

### Formatting selected cell ranges

- Selected cell rectangles now receive font formatting, paragraph alignment/bullets/line spacing, vertical anchoring, text direction, shading and margins. Commands retain the cell selection, validate bounds and complete merged-cell coverage before mutation, and leave neighboring cells unchanged. Supported Home controls and font inputs are enabled during range selection.
- Whole-cell font formatting now includes paragraph ends, empty paragraphs, fields and breaks while retaining existing unrelated formats. The range command shares the validated merged-cell targeting used by Clear.
- Validation: range-format backend tests cover merged/empty cells, invalid-operation atomicity, preserved neighbors, save/reload, replay and Undo. Seven core table-text tests pass. The browser workflow verifies range bold/shading in the downloaded PPTX, retained selection, Undo, then merge/edit/clear/Undo and selected-column deletion. Root/dev builds, both typechecks and targeted lint/format/whitespace checks pass.
- Native visual comparison remains unavailable (ScreenCaptureKit -3811); exact mixed-format indication and native toolbar defaults need comparison. Range borders/distribution/sizing, clipboard, keyboard range navigation, existing-merge consolidation and Split Cells remain incomplete. Full PowerPoint parity has not been achieved.

### Borders on selected cell ranges

- The Borders menu now uses the selected rectangle for all twelve modes, retaining the range after applying a border. Outside/inside and side-specific borders use selection edges rather than the whole table. Shared edges update both neighbors; merged-cell interiors are skipped and existing merged perimeter segments remain preserved.
- Range bounds and full merged-cell coverage are checked before writes. Backend coverage checks every grid edge for all twelve modes, unaffected edges, diagonal targeting, invalid/partial-merge atomicity, save/reload, replay and Undo. All 21 table/cell backend tests pass.
- The browser workflow passes Outside Borders on a selected rectangle, exported shared-edge equality, selection retention, Undo and subsequent merge/edit/clear/delete operations. Dev build, typecheck, targeted lint/format and whitespace checks pass. pnpm launcher failed switching to its configured 12.5.1 executable (ENOEXEC), so installed local tsdown/tsc/oxfmt/oxlint binaries were used directly.
- Exact Mac appearance and mixed-format indication remain unverified. Range sizing/distribution, clipboard, keyboard range navigation, consolidation of existing merges and Split Cells remain incomplete; full parity is not established.

### Selected row/column sizing and distribution

- Height/Width and Distribute Rows/Columns now target the selected cell rectangle's grid interval, retain the range, and preserve unselected grid dimensions. A single selected merged cell uses its full span for the requested dimension. Rotation compensation retains the table's rotated upper-left origin during size changes.
- Height/Width display the selected grid size and show Mixed for differing sizes; a single merged cell displays its combined span. Native behavior for mixed merged/unmerged selections and minimum content-driven row heights still needs implementation/comparison.
- Microsoft's Mac reference documents selecting rows/columns before distribution: https://support.microsoft.com/en-us/word/resize-a-table-in-word-or-powerpoint-for-mac . Native screenshot capture remains unavailable; this is not a claim of full visual or behavioral parity.
- Validation: all 23 table/cell backend tests pass, including both sizing/distribution axes, rotated origin, untouched grid dimensions, invalid partial-merge selection, merged width, save/reload/replay/Undo. The browser workflow passes selected-column Width input, neighboring-column retention, range highlight, Undo and subsequent formatting/merge/delete actions. Dev build/typecheck, targeted lint/format and whitespace checks pass using installed binaries directly.

### Consolidating existing table merges

The editor now enables Merge Cells for a selected rectangle containing multiple
visible cells, including fully contained existing merges. The core API exposes
this through `allowContainedMerges`; its default rejection behavior is preserved
for existing callers. Partial intersections are rejected before any XML changes.
Append mode copies visible anchor paragraphs, preserving rich formatting without
reintroducing hidden covered-cell text. Old span flags are cleared before the new
merge is written.

Validation: 19 core merge tests pass, including save/reload and atomic rejection;
the editor merge test covers both ordinary and previously merged cells through
save/reload, replay and Undo. Typecheck and builds pass. Exact Mac PowerPoint
visual and interaction parity remains unverified because native capture is
unavailable.

### Split Cells dialog and grid subdivision

Table Layout now provides Split Cells for a single visible cell, including an
existing merged cell. The dialog accepts row and column counts and uses Insert
and Cancel, matching the documented Mac operation:
https://support.microsoft.com/en-us/powerpoint/merge-or-split-table-cells-in-powerpoint .
The core `splitTableCell` API retains existing grid boundaries, adds equally
spaced split boundaries, and extends neighboring cells across the new grid lines
without changing the table frame or those cells' geometric extents. The original
rich text stays in the first cell; the other new cells are empty. Invalid counts
and covered-cell coordinates are rejected before mutations. Editor journal
commands preserve the operation through save/reload, replay and Undo.

Validation: 23 core split/merge tests pass, including OOXML schema validation;
21 editor tests selected by `table|cell` pass. A dedicated browser test exercises
cancel, two-axis subdivision, preserved edited text, PPTX export and Undo. Root
and editor typechecks, builds, targeted lint/format and diff checks pass. The
browser screenshot `/tmp/pptx-table-split.png` was visually inspected.

Limitations at initial implementation: native Mac pixel comparison is still
unavailable. Short split rows initially clipped text; the content minimum-height
work below addresses horizontal text. Exact native minimums, inherited
styles/banding and detailed border treatment need further comparison and work.
The dialog currently limits each requested dimension to 100; the native limit
has not been verified. Multiple-cell split selection is not implemented. This
addition does not establish full UI or operation parity.

### Cell content minimum height

- Horizontal table text now grows rows after text/format/layout, margin, cell-size, merge and split edits. Measurement uses the preview text layout engine, including wrapping, paragraph spacing, font sizes, margins and empty paragraph end formatting.
- A merged cell constrains the combined height of its row span. Growth preserves the table's rotated top edge and does not shrink existing rows.
- Verified: 23 table/cell editor tests, including merged-span growth, wrapping, empty-cell font size, margins, reload, replay and Undo; Split Cells browser test and screenshot inspection. The previously clipped Japanese split-cell example now fits inside its cell. Preview/dev typechecks and builds, targeted lint and diff checks pass.
- Remaining: default font measurement is approximate; no native Mac visual comparison is claimed. Vertical text fitting is not covered by this fitting policy; direct frame resizing was added below. Exact native minimum heights and split-cell style/border inheritance remain open.

### Minimum height across table grid operations

- Row/column insertion, row append and deletion now also apply content minimum heights. Deleting part of a merged cell cannot leave its retained text in a shorter frame than the layout measurement requires.
- Distributing columns recalculates wrapped text height. Distributing rows grows the entire selected row group when needed, retaining equal row heights (within EMU rounding), unaffected outer rows and the rotated top edge.
- Validation: 25 table/cell editor tests pass, including constrained distribution, narrowed columns, merged-row deletion, save/reload, replay and Undo. Native font metrics and exact Mac behavior remain unverified.
- Browser validation: insertion/formatting/distribution/export, structural insertion with margins, and deletion tests pass. The insertion test now verifies the content-constrained height instead of expecting a three-paragraph cell to shrink to 0.5 inches. Dev typecheck/build, targeted lint/format and diff checks pass.

### Table frame resizing and fixed edges

- Resizing a table through drag positions or numeric shape bounds now enforces horizontal content minimum heights. Moving/rotating without changing dimensions leaves authored row heights alone.
- North-side drag handles keep the opposite bottom edge fixed when minimum height limits the resize; south-side handles and numeric sizing retain the top edge. This also applies to rotated frames and live preview commands.
- Validation: 27 table/cell editor tests pass, including both anchor directions at 30 degrees, wrapping, unchanged translation/rotation dimensions, invalid anchor rejection, save/reload, replay and Undo. Dedicated browser test drags the north handle, checks fixed bottom edge and text minimum height in exported PPTX, and undoes. Screenshot /tmp/pptx-table-resize.png inspected. Dev typecheck/build and targeted lint/format pass.
- Native Mac visual/behavioral parity remains unverified. Vertical text fitting, full native font metrics, and group-resize policies remain incomplete.

### Native table context-menu comparison

- Native AX and screen capture are available again. On the installed Mac PowerPoint, temporarily inserted a 5-column, 2-row table into the authorized support reference, inspected its cell context menu, and undid the insertion (Undo disabled, Redo Table visible afterward).
- Confirmed native labels and order: Insert → Columns Left, Columns Right, Rows Above, Rows Below; Delete → Columns, Rows, Table; Select → Table, Column, Row; Merge Cells; Split Cells. Native also includes text formatting, language, hyperlink, comment and device commands.
- Added Insert/Delete submenus and Merge/Split commands for cell editing and selected cell ranges, using the existing editor operations. Ribbon and context deletion share one menu implementation and retain the surviving cell focus. Direct table right-click starts editing the cell at the pointer.
- Scope remains partial: native Select and remaining text/context commands are not yet mirrored, range insertion still follows the existing insertion implementation, and full native visual parity is not claimed.
- Validation: extended Split Cells browser test passes through right-click split, export/Undo, direct right-click row insertion and row deletion with focus restoration. Dev typecheck/build and targeted lint pass. Direct hit lookup uses the editor shadow root, verified by the browser regression.

### Table selection menus

- Added Select Table, Select Column and Select Row to the cell context menu and a Select control at the start of Table Layout. The ribbon control is disabled for a table frame without an active cell/range. Already selected axes are disabled in the menu.
- Selection commits pending cell text, retains the current row/column extent when expanding an existing range, and shares merged-cell expansion with drag selection. Keyboard clearing and structural deletion consume that selected range.
- Compared the installed Mac app directly: Select is disabled on the table frame, enabled after entering a cell; Select Row highlights the entire row and disables Select Row; selecting columns from that row highlights the entire table. The temporary table was undone, leaving Undo disabled and Redo Table available.
- Corrected floating-point noise in table minimum-height measurement so an exact integer height does not gain an extra EMU when rounded upward. Regression covers an empty default cell with exactly 365760 EMU minimum height.
- Validation: two cell projection/selection unit tests, 10 preview table tests, 27 table/cell editor tests and the drag/merge browser regression pass. Broader native parity remains incomplete, including the rest of the table context commands and ribbon visual details.

### Range insertion following table selection

- The Select menu browser regression passes: pending text is committed, row clearing retains the grid, column selection drives column deletion, and the ribbon selects the entire table.
- Insert Above/Below/Left/Right now consumes the selected cell range. It inserts the selected number of rows or columns at that range's corresponding edge and focuses the first inserted cell. Cell-only and table-frame insertion retain their existing behavior.
- Multiple insertion count follows Microsoft's documented behavior: https://support.microsoft.com/en-us/powerpoint/add-or-delete-table-rows-and-columns . Inserted dimensions retain the existing adjacent-cell sizing policy; exact native sizing for unequal selections still needs comparison.
- Native comparison could not continue in this session because the computer-use transport closed. The earlier temporary reference table had already been undone.
- Validation: dev typecheck/build, targeted lint, diff checks and all 28 table/cell editor tests pass. The new regression checks four insertion directions, retained original cell contents, invalid ranges, export/reload, journal replay and Undo.
- Extended browser regression also passes through selecting the entire two-row table, inserting two rows below, restoring focus to row three, and exporting a four-row table.

### Clipboard on first table-cell context click

- Direct right-click on a table cell now shows Cut/Copy/Paste immediately when it starts the editing session. Previously these entries appeared only when the user right-clicked an already active cell editor.
- Both paths share clipboard enablement and actions. A collapsed text selection disables Cut/Copy; Paste inserts into the active cell, and text Cut participates in local Undo.
- Validation: extended table-selection browser regression checks first-click Paste, selected-text Cut, local Undo and subsequent structural operations. Existing text clipboard race/error/selection regression also passes. Dev typecheck/build, targeted lint/format and diff checks pass.
- Native transport remains closed; the Font/Paragraph and other missing native context commands and precise visual comparison remain open. This does not establish full context-menu parity.

### Font dialog from text and table context menus

- Added Font... to text-edit and table-cell/range context menus. The dialog edits font family, regular/bold/italic style, size, color, single/double underline, strikethrough, capitalization and baseline position through the existing character-format commands.
- Only changed controls are submitted, preserving unrelated mixed formatting. The editing selection is captured before opening and restored for Cancel, Escape, and successful application. Modal inertness is released before the editor rebuilds the text field; blur into this dialog retains the edit session.
- Browser regression passes through Cancel with a changed size, applying size 24 plus bold/italic to only `Context` in `Context paste`, retained selection and unchanged suffix formatting in the saved PPTX, followed by column deletion and multi-row insertion. Existing text-clipboard regression passes. Dev typecheck/build, targeted lint/format and diff checks pass. Screenshot `/tmp/pptx-font-dialog.png` inspected.
- This is an initial implementation, not native visual parity. Mixed values are initialized from the active run rather than displayed as indeterminate; full native tabs, character spacing/kerning, custom offsets, theme-color/font pickers and exact control arrangement remain unfinished. Reference for native Font effects: https://support.microsoft.com/en-us/powerpoint/format-text-as-superscript-or-subscript . Native app transport is still unavailable for direct comparison.

### Font character spacing and position controls

- Added Font / Character Spacing tabs with keyboard tab navigation, normal/expanded/condensed tracking, explicit point spacing, kerning threshold and custom superscript/subscript percentage offsets. Character Spacing → More Spacing opens the spacing tab directly.
- Only changed properties are applied. Added the missing validated kerning field to the editor format conversion so it survives range formatting and PPTX export.
- Tracking currently follows the editor's ±20 pt limit. Kerning is stored in PPTX, but the preview measurement/rendering engine does not yet apply its threshold; native visual parity and mixed-value presentation remain incomplete. Native comparison remains unavailable because the computer-use transport closed.
- Validation: dev typecheck/build, targeted lint and diff checks pass. Extended table-selection browser regression verifies selected text exports with baseline 40%, condensed spacing 1.5 pt and kerning threshold 14 pt while preserving the unselected suffix and selection.

### Kerning in editing and preview rendering

- Connected explicit kerning thresholds to rich-text editing, HTML preview and SVG text pieces. Thresholds use the effective font size after autofit; zero disables kerning and absent values retain the renderer default.
- Fontkit measurement now honors explicit kerning state. Measurement cache keys and adjacent SVG run grouping preserve differing kerning states; the audit measurer receives the same state.
- Validation: 52 targeted layout/audit/table-preview tests pass, including real glyph width changes, threshold boundaries and adjacent differing states. Preview and dev typechecks/builds, targeted lint and diff checks pass. Table-selection browser regression passes and verifies computed kerning on the selected edited text alongside exported values and existing structural operations.
- References for underlying renderer APIs: https://github.com/foliojs/fontkit#fontlayoutstring-features---- and https://github.com/linebender/resvg/blob/main/CHANGELOG.md (0.26.0 font-kerning support). Full native UI comparison remains outstanding. Custom baseline offsets are saved but current preview still renders generic super/sub positions; that rendering gap remains open.

### Authored superscript/subscript offsets in rendering

- Propagated the authored baseline fraction into SVG text pieces and HTML/editing CSS rather than reducing all positive/negative values to one generic position. SVG adjacent runs with different offsets remain separate, and wavy underlines use the same baseline shift.
- SVG width measurement now uses the displayed super/subscript glyph size, and line extents account for the shifted glyphs. HTML text uses an explicit reduced size derived from the run's effective size rather than an em relative to the paragraph; editing now applies that same size ratio.
- Validation: 54 targeted preview/layout/audit tests pass, including +40%/-20% offsets, display size, distinct adjacent offsets and glyph measurement. Preview/dev typechecks and builds, targeted lint and diff checks pass. The table-selection browser regression passes and verifies the edited span's baseline-to-font-size ratio together with saved values and selection preservation.
- Native computer-use was retried and still returns Transport closed. The existing 0.65 script glyph-size ratio is retained and needs installed Mac verification, as do exact line-height/offset interaction and full UI geometry. Source for the Offset operation: https://support.microsoft.com/en-au/powerpoint/format-text-as-superscript-or-subscript . This remains partial progress toward full UI/operation parity.

### Paragraph dialog and paragraph settings

- Added Paragraph... to text/table context menus and Line Spacing Options... to the line-spacing menu. Controls cover alignment, left/right indentation, first-line/hanging indentation, spacing before/after, single/1.5/double/multiple and exact point line spacing. Only changed controls are submitted; Cancel/Escape preserve the editing selection.
- Added validated paragraph-setting APIs for shapes and cell text ranges. XML patches preserve runs, bullets and untouched properties, with spacing children in schema order. Editor commands support selected shape paragraphs, cell paragraphs and selected cell ranges; shape autofit measures the resulting paragraph settings.
- Validation: 8 core range/serialization tests, 6 targeted editor/history regressions and the table-selection browser regression pass. The browser checks Cancel, selection restoration, applying hanging indentation and exact line spacing, retained inline formatting and exported paragraph properties. Root/dev typechecks, dev build, targeted lint and diff checks pass.
- Native visual parity remains unverified while computer-use transport is unavailable. Dialog geometry, labels, unit preferences, native spacing choices and mixed-value display still require installed Mac comparison. Initial values use the active paragraph; no indeterminate display for mixed paragraphs yet. This completes this paragraph-settings increment, not overall PowerPoint parity.

### Mixed paragraph values across the actual selection

- Paragraph dialog now derives initial values from every touched text paragraph, selected shape paragraph or selected table cell, rather than only the first paragraph. Differing values appear blank with an accessible mixed-value description. Equivalent alignment tokens share the same displayed value.
- Blank mixed values are not submitted as zero and do not prevent unrelated edits. Explicit zero is submitted. Indentation direction/amount and line-spacing mode/amount remain coupled; selecting a mode supplies an editable amount.
- Validation: dev typecheck/build, targeted lint and diff checks pass. Two browser tests cover differing numeric values, shared alignment, explicit zero, coupled hanging-indent/exact-spacing edits, Escape, and a selected table column with mixed paragraph settings. Export confirms only the changed spacing property is unified and existing per-cell spacing is retained.
- Retried native computer-use documentation entry point: Transport closed. Native arrangement, units, mixed-value styling and all-operation parity remain unverified/incomplete; no changes were made to the reference presentation.

### Mixed font values across text and cell selections

- Font dialog now aggregates intersecting text runs, selected shapes and selected cell ranges. A caret continues to use its active insertion format. Mixed font names/sizes/styles, underline, capitalization, positions and spacing show blank values; mixed strikethrough/kerning use indeterminate checkboxes. Mixed colors show a patterned swatch with an accessible mixed-value description.
- Unchanged mixed values remain untouched. Blank numeric controls do not block unrelated edits or become accidental zeroes. Explicit Normal position/spacing and cleared effects still apply their reset values.
- Validation: dev typecheck/build, targeted lint and diff checks pass. Two browser regressions cover changing only size with mixed formats, resolving indeterminate effects, baseline/spacing/kerning resets, Escape and full-range mixed size/style display in the live table editor. Existing range formatting and exported PPTX checks still pass.
- This is functional selection consistency, not proof of installed Mac visual parity. Native color/font pickers, exact mixed-control appearance, defaults inherited from themes and the complete operation/UI inventory remain outstanding.

### Format menu and Mac dialog shortcuts

- Added Format to the application menu bar with Font... and Paragraph... linked to the existing dialogs. Entries are disabled without an editable text target or during an edit request. The menu uses the shared keyboard/hover menu-navigation system.
- Connected Command+T and Command+Option+M for text editing and selected text shapes/cells. Option-modified M uses the physical key code as well as the character, accommodating the Mac µ character. Dialog fields, thumbnails, slide sections, composition and unrelated modifier combinations retain their own behavior.
- Validation: dev typecheck/build, lint and diff checks pass. Browser regression checks both menu entries, cancellation, selection preservation and cancelable shortcut events in a cell editor, followed by the existing formatting/export/table operations.
- Shortcut tests establish behavior once the page receives the event; they do not establish interception of OS/browser-reserved combinations. Format's full native item inventory and exact menu geometry remain incomplete. References: https://support.microsoft.com/en-us/accessibility/powerpoint/use-keyboard-shortcuts-to-create-powerpoint-presentations and https://support.microsoft.com/en-us/powerpoint/set-or-clear-tab-stops-in-powerpoint .

### Mac paragraph alignment and script-position shortcuts

- Connected Command+E/J/L/R to center/justify/left/right paragraph alignment and Command+Control+Equal / Command+Control+Shift+Equal to subscript/superscript toggles. Existing range-aware edit commands retain the text selection and support table cells. Command+[ / Command+] promote/demote ordinary shape paragraphs; table paragraph levels remain unsupported and are explicitly excluded.
- Modifier combinations and composition are checked before consuming the shortcut. Physical Equal handles the shifted plus character. Existing Font and Paragraph dialog shortcuts remain supported.
- Validation: dev typecheck/build, targeted lint and browser regression pass. Tests cover shape paragraph levels, untouched composition/unrelated modifiers, all four cell paragraph alignments, script toggle/reset, preserved text ranges and saved PPTX run boundaries; existing dialog/table regressions still pass.
- Native computer-use transport remains closed. This verifies received browser events, not OS/browser interception or full installed Mac appearance/operation parity. Overall parity remains incomplete.

### Table paragraph level editing

- Added range-aware table paragraph level adjustment, independently clamping each touched paragraph to 0–8. Invalid ranges/deltas are rejected before mutation; runs, bullets and other paragraph settings are preserved.
- Connected table cell editing and selected cell ranges to Increase/Decrease List Level ribbon commands and Command+[ / Command+] shortcuts. Removed the prior table shortcut exclusion and enabled the corresponding ribbon controls. Existing table text fitting and selection restoration apply to these commands.
- Validation: core serialization/range checks cover mixed paragraph levels, both bounds, invalid input, untouched runs and neighboring cells. Browser regression covers cell text selection, both shortcut directions, selected-column ribbon editing, exported levels and retained cell selection. Root/dev typechecks, builds, targeted lint and diff checks pass.
- Native level-dependent indentation/default inheritance and exact Mac UI comparison remain unverified. This closes the missing table edit-command path, not overall parity.

### Selected-text case conversion and Format Shape shortcut

- Fixed lowercase/uppercase commands modifying an entire shape while a text range was selected. Changes now target the selected range, retaining each source character's formatting; selected cells and cell text editing use the same conversion. Whole-selection conversion preserves contextual Unicode casing, and expanded results such as ß → SS restore the adjusted selection length.
- Document Undo/Redo invoked during text editing now reopens the same surviving shape/cell after applying history. Local uncommitted text history remains first in precedence.
- Connected Command+Shift+1 to the existing Format Shape pane for selected objects. This exposes the current Size & Properties pane; it does not complete missing native pane categories.
- Validation: dev typecheck/build, lint/diff checks, case helper test, two editor serialization tests and table browser regression pass. Coverage includes partial selection, mixed formatting, Unicode expansion/final sigma, unmodified suffixes, saved cell formatting, Undo restoring cell editing, and the pane shortcut.
- Remaining: sentence/title/toggle case menu entries, Shift+F3 cycling, caret word behavior, native locale-specific rules, complete Format Shape pane and native visual comparison. Microsoft references: https://support.microsoft.com/en-us/word/change-the-capitalization-or-case-of-text and https://support.microsoft.com/en-us/accessibility/powerpoint/use-keyboard-shortcuts-to-create-powerpoint-presentations .

### Complete Change Case menu choices

- Added Sentence case., Capitalize Each Word and tOGGLE cASE alongside lowercase/UPPERCASE, routed through the same selection-aware shape and cell commands. Sentence/word boundaries use Intl.Segmenter; punctuation and uncased characters are retained. Lowercasing still uses the full selection for contextual forms such as Greek final sigma.
- Validation: two helper tests and the table browser regression pass, including quoted sentences, newlines, accented words, apostrophes, Unicode expansion, selection boundaries and consecutive native-named menu choices with restored text selection. Dev typecheck/build, lint and diff checks pass.
- Native computer-use retry still reports Transport closed. Locale-specific PowerPoint segmentation/titlecasing, Shift+F3 cycling, caret word behavior and exact native menu appearance remain unverified or incomplete; full parity is not achieved.

### Basic Fill & Line pane operations

- Added Fill & Line / Size & Properties category tabs to Format Shape, with keyboard navigation. Fill supports No fill / Solid fill and color; Line supports No line / Solid line, color, width in points and preset dash styles. Controls apply through existing selection/history commands; unsupported objects disable paint controls, and mixed widths remain blank.
- Added editor stroke commands and paint state. Corrected the core width-only stroke setter to preserve the current fill choice, including color transforms/opacity.
- Validation: dev typecheck/build and lint/diff checks pass; two core suites (five tests), editor export/replay/undo regression and the integrated browser regression pass. Browser coverage includes pane tabs, fill/line type, color, width preserving color, dash and disabled controls after No line. Core regression verifies opacity preservation through export after width changes.
- Native computer-use retry still fails with Transport closed. Exact Mac layout, native color palettes, theme/style color resolution, transparency controls, gradient/pattern/picture fills, effects and remaining outline controls are incomplete. This adds basic working controls and does not establish exact visual or behavioral parity.
- Microsoft reference consulted: https://support.microsoft.com/en-us/powerpoint/add-shapes (Mac Format Shape > Fill & Line).

### Solid fill and outline transparency

- Added paired 0–100% transparency number inputs/sliders for explicit solid fill and line paint. Mixed values leave the number field blank; inherited/non-solid paints disable these controls until supported.
- Added public setShapeFillOpacity/setShapeStrokeOpacity APIs with 0–1 validation. These replace alpha/alphaMod/alphaOff with an absolute alpha while retaining other color transforms. Editor color changes retain existing opacity; width changes continue preserving it.
- Validation: root build, dev typecheck/build, lint/diff pass. Four core opacity tests, editor export/replay/undo regression, and integrated browser regression pass. Checks cover replacement of existing alpha transforms, 0/1 endpoints, invalid numeric values, exported color/opacity, numeric fill transparency and outline slider synchronization.
- Remaining: inherited/theme paint materialization, gradient/pattern/picture transparency semantics, native palette/geometry and full Mac comparison. Full UI/operation parity remains incomplete.

### Line cap, join and compound controls

- Connected Cap type, Join type and Compound type selectors to existing shape stroke APIs. Changes retain color, width and opacity; mixed/unspecified values remain blank rather than guessing inherited defaults. No line disables these controls.
- Validation: dev typecheck/build, lint/diff, editor export/replay/undo and integrated browser tests pass. Exported cap/join/compound values persist, SVG includes square caps and bevel joins, and browser changes retain outline opacity.
- Compound preview remains incomplete: the renderer only approximates double lines and does not accurately paint all compound variants. Native gallery icons, inherited defaults and exact Mac appearance are unverified. Saving a compound value is not proof of full rendering parity.
- Reference for PowerPoint control names: https://download.microsoft.com/download/5/9/1/5914ff35-1227-4dc4-a011-7251db29bd1b/AF102264791_en-us_mhillpowerpointqs_ch07.pdf . This older reference is not evidence of current Mac visual parity.

### Symmetric compound outline rendering

- Replaced the double-line width-only approximation with transparent SVG masks for double and triple outlines. Shape/custom geometry and straight/bent/curved connectors share the renderer; masks preserve underlying fill/background and outline opacity. Arrow markers are emitted separately so the gaps do not cut them out.
- Validation: preview typecheck/build, lint/diff and five focused raster/custom-geometry tests pass. Raster samples verify both bands/gaps for double and triple rectangles and horizontal connectors, including translucent outlines over a colored background. The wider render suite has 29 passes and two text-overlay rotation failures; both failures reproduce with compound rendering bypassed, so they are independent of this change and remain unresolved.
- Remaining: asymmetric thick/thin variants, detailed cap/dash/corner/arrow fidelity, picture borders and comparison against native Mac rendering. Full UI/operation parity is still not complete.

### Text rotation regression coverage

- Resolved the two text-overlay test failures: centered text intentionally uses SVG text for measured placement, while the old assertion required a foreignObject. Actual rotation output was correct; no renderer change was needed.
- Expanded coverage across centered/default, HTML and SVG text layouts, testing all four horizontal/vertical flip combinations at 30 degrees. Assertions verify text content, the expected rendering path and independent text rotation.
- Validation: all 37 tests in the render SVG, compound stroke and custom geometry suites pass. Native computer-use retry still reports Transport closed, so exact Mac visual comparison remains unavailable and full parity remains incomplete.

### Begin and end arrow controls

- Added Begin/End Arrow type and size controls to Fill & Line, covering all six OOXML end types and nine width/length combinations. No line disables the controls; an absent arrow disables its size control. Mixed selections leave differing values blank.
- Editor commands merge only the changed arrow fields per selected shape, retaining individual sizes when changing type and preserving the opposite endpoint and stroke paint. Arrow state is included in the editor model; export, replay and undo use existing transactions.
- Validation: dev typecheck/build, lint/diff, two focused editor tests and integrated browser regression pass. Checks cover multi-object size preservation, opposite-end edits, exported SVG markers, replay/undo, browser type/size changes and disabled size after No Arrow.
- Remaining: native icon galleries and exact Mac pane layout, inherited arrow style resolution and native comparison. Current text-based selectors provide working operations but do not establish exact UI parity. Full parity remains incomplete.

### Arrow preview galleries

- Replaced the text-only arrow type/size selectors with vector preview galleries. Begin and end previews face their corresponding endpoint; the size gallery shows all nine width/length combinations for the current arrow type. Selected values are highlighted, with accessible option names and mixed-state display.
- Added arrow-key/Home/End navigation, Enter selection, Escape cancellation and trigger focus restoration. Popovers use the browser top layer and viewport bounds to avoid pane clipping. Disabled controls close an open gallery. Excluded gallery controls from the editor's capture-phase shape shortcuts so arrow keys cannot move the selected object.
- Validation: dev typecheck/build, lint/diff and screenshot inspection pass. Integrated browser coverage exercises type selection, the nine preview options, row navigation, selection persistence, cancellation/focus and No Arrow disabling, followed by existing table operations.
- Exact native Mac gallery dimensions, order, colors and focus behavior remain unverified because native comparison is unavailable. This improves visual selection but does not establish full UI parity.

### Triple outline band proportions

- Corrected triple outlines from three equal-width bands to thin/thick/thin, with line/gap/line/gap/line proportions of 1:1:2:1:1. Added raster samples inside the expanded center band and narrowed outer bands for rectangles and connectors, retaining transparent gaps.
- Evidence: Microsoft's DrawingML CompoundLineValues documentation identifies `tri` as thin/thick/thin (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.compoundlinevalues?view=openxml-3.0.1). The numerical ratio follows Microsoft's archived VML stroke documentation (https://learn.microsoft.com/en-us/windows/win32/vml/web-workshop---how-to-use-vml-on-web-pages-----stroke--element); this is a rendering reference, not verification of the installed Mac version.
- Validation: 37 compound/custom-geometry/render SVG tests, preview typecheck/build and focused formatting/lint checks pass. Native tool reset still fails with Transport closed. Asymmetric compound variants and exact native comparison remain outstanding; full UI/operation parity is incomplete.

### Asymmetric outlines on closed geometry

- Added thick/thin and thin/thick compound rendering for rectangles, ellipses, polygons and explicitly closed path subpaths. A silhouette clip identifies the interior without depending on path winding; it removes the appropriate quarter-width gap while retaining the full outside or inside half-width band. Even-odd fill rules are carried into the clip. Open paths retain their existing rendering until directional offsets are implemented.
- Evidence: inspected Microsoft's published thickthin rectangle sample (https://learn.microsoft.com/en-us/windows/win32/vml/images/thickthin.gif), which puts the thick band outside, and used the 2:1:1 / 1:1:2 ratios documented in the archived VML reference above. These references do not prove current Mac rendering fidelity.
- Validation: all 39 compound/custom-geometry/render SVG tests pass, with new raster checks for both asymmetric styles on rectangles and ellipses over a contrasting background, with and without shape fill. Preview typecheck/build, focused formatting/lint and diff checks pass.
- Remaining: asymmetric open connectors/paths, native Mac comparison, detailed cap/dash/corner/arrow fidelity and picture borders. Full UI/operation parity remains incomplete.

### Picture outline rendering

- Fixed the picture render path dropping outlines entirely. Pictures and image-filled shapes now use the same geometry/stroke renderer as ordinary shapes, including solid paint, width, stroke opacity, dash/cap/join and compound styles. The shared geometry helper retains custom/preset geometry and unknown-geometry flags for normal shapes.
- Outlines are drawn after the image in a separate transformed group, outside its crop clip and image filter/opacity. Missing-image placeholders also retain the authored outline. No line omits the border.
- Validation: 57 compound/custom-geometry/render SVG/imported-artwork tests pass, along with preview typecheck/build, focused lint/formatting and diff checks. Added raster coverage for cropped pictures with solid, double, triple, both asymmetric and removed borders; image and stroke transparency remain independent.
- Native comparison remains unavailable: the reference /tmp/pptx-macro-audit directory is absent, and an alternate read-only System Events query failed with error -10827. Nonrectangular image clipping, picture effects, asymmetric open paths and exact Mac UI/operation parity remain outstanding.

### Picture geometry clipping

- Added an independent silhouette mask to picture/image-fill rendering using the same preset/custom geometry as the outline. Crop and image corrections are applied within that shape; the outline and text remain outside the mask. Missing-image placeholders also use the silhouette.
- Fixed preset/custom geometry readers excluding picture objects despite their spPr geometry. Shape masks ignore authored fill/stroke transparency and retain custom path fill rules.
- Validation: 87 focused picture-geometry, compound, custom-geometry API/render, SVG and imported-artwork tests pass. New raster cases cover cropped pictures and image fills using ellipse/roundRect, with outlines, without outlines and at 90-degree rotation. Root and preview builds, preview typecheck, focused formatting/lint and diff checks pass.
- Remaining: native Crop to Shape UI, picture effects, custom geometry masking edge cases, asymmetric open paths and full Mac UI/operation comparison. The shared mask supports custom paths structurally, but native custom-picture fidelity is not yet verified.

### Picture Crop to Shape command

- Added a picture-only contextual Picture Format tab with Crop → Crop to Shape and the existing categorized shape gallery. Nested galleries support grid arrow navigation and Escape focus restoration. The tab hides and returns to Home when the picture selection ends.
- Added setShapeImageCropShape and the imageCropShape edit command. Geometry replacement preserves picture bytes, source crop, bounds, rotation and outline; exported files and history replay/undo retain the operation.
- Validation: 87 related renderer/API cases pass; a new editor export/replay/undo test and a browser test covering insertion, crop selection, gallery keyboard navigation, undo and contextual tab visibility pass. Core/dev builds, dev typecheck, focused formatting/lint and diff checks pass.
- The Picture Format ribbon remains incomplete. Crop handles, Fit/Fill, aspect ratios, picture effects and exact installed Mac appearance/behavior still need implementation/comparison. Native CUA remains unavailable; this is not full PowerPoint parity.

### Picture Fill / Fit

- Added Crop → Fill / Fit and a setShapeImageFit API. Centered scaling keeps source proportions and the existing picture frame; Fill crops overflow and Fit leaves space outside the source image. Shape outlines, rotation and other picture settings remain intact.
- Enabled signed source crop fractions in the API and renderer for Fit margins. Validation now runs before XML mutation and rejects nonfinite values, overflow and empty source regions without deleting previous crop settings.
- Added imageCrop to the editor model for observable crop state. PNG/JPEG dimensions are inferred; API callers can supply decoded dimensions for other formats. The editor still needs automatic dimension detection for other formats.
- Validation: 52 focused crop/picture/outline/SVG tests, all 75 editor tests, and the expanded browser crop test pass. Coverage includes wide/tall Fill and Fit pixel output, export, replay, undo, preservation of the frame and invalid edit atomicity. Core/preview/dev builds, preview/dev typechecks, focused lint/format and diff checks pass.
- Reference: Microsoft's macOS Fit/Fill instructions at https://support.microsoft.com/en-us/office/graphics-visuals/crop-a-picture-to-fit-in-a-shape . Native crop handles, source-image dragging, aspect-ratio menu and exact Mac visual comparison remain outstanding. This implementation does not yet reproduce the complete native crop interaction.

### GIF and WebP picture proportions

- Extended shared image dimension reading to GIF87a/GIF89a logical screens and WebP VP8, VP8L and VP8X canvases. Insert/replace contain sizing and Crop Fill/Fit now infer proportions for these formats. Animated images use canvas dimensions; RIFF chunk padding and offset byte views are supported.
- Rejects truncated dimension headers, out-of-bounds RIFF chunks, invalid lossless versions and invalid canvas sizes. This reads dimensions only; it does not validate compressed pixels.
- Validation: 39 focused image sizing, crop, format and Fill/Fit rendering tests pass. New header fixtures cover endian/bit-field decoding, malformed input, insertion/replacement, Fill/Fit and saved geometry/crop. Core/preview/dev builds and preview/dev typechecks pass; focused lint/format checks pass. GIF/WebP fixtures in these tests are header-only; decoded display fidelity and animation playback are not verified by this change.
- References: https://www.w3.org/Graphics/GIF/spec-gif89a.txt and https://developers.google.com/speed/webp/docs/riff_container . Other image formats, native crop interaction and complete Mac UI parity remain outstanding.

### Picture crop aspect ratio

- Added Crop → Aspect Ratio with square, portrait and landscape ratio choices. Ratio submenus use ordinary menu rows rather than the shape-gallery grid; existing shape-gallery keyboard handling remains covered by the browser test.
- Added setShapeImageCropAspectRatio and the imageCropAspectRatio edit command. The implementation crops a centered rectangle within the current frame, retaining source scale, asymmetric crop offsets, rotation and image bytes. Invalid/nonfinite/empty ratios are rejected before mutation. Saved geometry and crop percentages respect OOXML integer precision.
- Validation: 14 focused crop API/render tests, a new editor export/replay/undo test and the expanded browser crop test pass. Core/dev builds, dev typecheck, focused lint/format and diff checks pass.
- Reference: https://support.microsoft.com/en-us/office/graphics-visuals/crop-a-picture-in-office documents the Crop → Aspect Ratio operation. Native CUA was retried and still fails with Transport closed. Exact installed Mac menu contents, ordering, metrics and frame adjustment behavior are unverified. This command does not yet enter an interactive crop mode; handles, source dragging and ratio-constrained drag adjustment remain required for parity.

### Interactive picture crop foundation

- Added Crop mode for a single selected picture. Eight black edge/corner handles change the picture frame and source crop together without scaling the source. Dragging inside the frame pans the source image. Pointer motion follows local rotated axes and source crop honors flips.
- Uses the existing server-rendered drag preview and history commit path. Pointer cancellation or Escape during a drag discards that drag; Escape while idle closes crop mode. Selecting another object/slide exits the mode. Undo restores committed frame/crop changes.
- Added the imageCrop edit command and isolated crop geometry calculations in picture-crop.ts.
- Validation: four geometry tests (corners, edges, pan, rotation, flips, existing signed crop, outward crop and minimum extent), four picture editor tests and the expanded browser test pass. Browser coverage includes handle drag, source pan with unchanged bounds, both Undo operations, cancellation, mode exit and the existing shape/aspect/Fill/Fit menus. Dev typecheck/build, focused lint/format and diff checks pass.
- Remaining: native split Crop ribbon button, dimmed full-source overlay outside the frame, ratio-constrained crop dragging, modifier behavior, multi-picture crop behavior, source resize handles and exact Mac appearance/interaction comparison. This is an interactive foundation, not complete crop fidelity.

### Dimmed source outside the crop frame

- Added a crop-only SVG overlay that reuses the source image while masking out the active rectangular frame. The remaining source is dimmed and follows current crop, rotation and flips. Existing image opacity/filter references are retained. The overlay updates during pointer movement and is removed on mode exit; it never enters saved slide content.
- Validation: the expanded browser crop test passes, including screenshot pixel checks (white source stays white inside the frame and becomes gray outside), source/frame extent checks, mode cleanup and existing drag/pan/Undo/Escape interactions. Visually inspected /tmp/pptx-crop-source-overlay.png. Dev typecheck/build and focused lint/format/diff checks pass.
- The screenshot covers an unrotated opaque picture. Rotated/flipped, transparent and corrected images still need visual comparison. Native dimming level and exact overlay/handle geometry remain unverified. Ratio-constrained drag, source resizing and complete Mac UI parity are still outstanding.

### Ratio-constrained crop dragging

- Selecting an aspect ratio now enters crop mode for the same single selected picture after the command succeeds. All eight frame handles preserve that ratio while retaining source scale. Corners anchor the opposite corner; side handles anchor the opposite edge and center the perpendicular axis. Source panning retains the fixed frame.
- Ratio constraints last for the crop session; exiting crop mode or explicitly choosing Crop clears them. Minimum frame dimensions prevent crossed handles from producing an empty crop.
- Validation: all five crop geometry tests and the expanded browser crop test pass. Browser coverage verifies automatic mode entry, a constrained corner drag and Undo, followed by the existing free crop, source pan, overlay pixels and Escape checks. Dev typecheck/build, focused lint/format and diff checks pass.
- Native Mac ratio persistence, handle anchoring, modifier keys and exact menu/visual behavior remain unverified. Source resizing and full PowerPoint UI/operation parity remain outstanding.

### Crop split ribbon control

- Split Crop into an icon button that toggles interactive crop mode and a separately accessible Crop options menu button. The main button exposes pressed state, supports native button keyboard activation and is disabled unless one picture is selected. Entering/exiting mode leaves the document unchanged; the menu retains shape, aspect ratio, Fill and Fit commands.
- Reference: https://support.microsoft.com/en-us/office/graphics-visuals/crop-a-picture-to-fit-in-a-shape describes using the arrow under Crop for options and the Crop button or Escape to finish.
- Validation: the expanded browser test passes, including click entry/exit, Space activation, Escape, pressed state, unchanged picture model and all existing crop interactions. Dev typecheck/build, lint and diff checks pass. Visually inspected the updated screenshot and corrected the split control height to fit the ribbon.
- Native CUA again returns Transport closed. Exact Mac ribbon metrics and appearance, multi-picture behavior, source resize handles and full UI/operation parity remain unverified or incomplete.

### Picture source resizing during crop

- Added an outline and four round corner handles around the full source image in crop mode. Dragging a source handle scales the image proportionally about its opposite source corner while preserving the crop frame. Source handles update along with the dimmed overlay; frame crop handles remain above them.
- Source resizing uses local rotated axes and maps visual edges through horizontal/vertical flips. It reuses the existing crop preview, commit and history path. Exiting crop mode removes the source frame and handles.
- Validation: six crop geometry tests pass, including all source corners with four flip combinations at 90-degree rotation, proportional sizing, fixed opposite corners and crossed-handle minimum size. Expanded browser test verifies source drag, unchanged frame, Undo and subsequent source pan. Dev typecheck/build, lint/format and diff checks pass; updated source/frame screenshot visually inspected.
- Native source handle dimensions, overlap priority, modifiers and anchoring behavior have not been compared with the installed Mac app. This adds source resize capability but does not establish complete crop or overall PowerPoint parity.

### Fill/Fit crop mode transition

- Fill and Fit now enter interactive crop mode after a successful edit for the same single selected picture. Frame handles, source resizing and source panning are immediately available; Escape exits. Shared the guarded transition with the aspect-ratio command, retaining its session-specific ratio constraint while clearing that constraint for Fill/Fit.
- Reference: https://support.microsoft.com/en-us/office/graphics-visuals/crop-a-picture-to-fit-in-a-shape explicitly describes crop handles appearing after Fill/Fit. Multi-selection edits still apply without entering the single-picture interactive mode.
- Validation: expanded browser test independently verifies Fill and Fit mode entry, source handles, pressed state and Escape exit, alongside existing crop/gallery/Undo coverage. Dev typecheck/build, focused lint/format and diff checks pass. Exact native Mac behavior and full UI/operation parity remain unverified/incomplete.

### Picture transparency gallery and pane

- Added Picture Format → Transparency with picture thumbnails for 0, 15, 30, 50, 65, 80 and 95 percent transparency. Picture Transparency Options opens the Format Picture pane, with numeric and slider input from 0 to 100 percent. Pane tabs support keyboard navigation; controls use the existing update/history path and support picture multi-selection.
- Added imageOpacity to the editor model and edit command, with validation before mutation and command-specific history conflict detection. Existing history fingerprints remain unchanged for unrelated commands.
- Reference: https://support.microsoft.com/en-gb/office/graphics-visuals/make-a-picture-transparent documents the presets, options pane, slider and numeric input.
- Validation: focused editor export/reload/replay/conflict/invalid-input/Undo test passes. Expanded browser crop test passes, including gallery selection, numeric input, slider ArrowRight, synchronized numeric value and Undo, alongside existing crop interactions. Dev typecheck/build, focused lint/format and diff checks pass. Gallery and pane screenshots visually inspected.
- Installed Mac ribbon/pane appearance and precise behavior have not been compared; native CUA remains unavailable. Image corrections/effects and full PowerPoint UI/operation parity remain incomplete.

### Picture correction fields

- Added brightness and contrast controls (numeric and slider, -100 to 100 percent) to Picture Corrections in the Format Picture pane. Values are exposed through the editor model and update command, using the existing DrawingML luminance APIs. Command-specific history fingerprints detect correction conflicts without changing unrelated legacy fingerprints.
- Reference: https://support.microsoft.com/en-gb/office/graphics-visuals/change-the-brightness-contrast-or-sharpness-of-a-picture documents correction fields. Native Mac CUA was retried and still returns Transport closed.
- Validation: two editor tests cover endpoints, export/reload, replay, conflict detection, invalid values and Undo. Expanded browser test passes for both correction fields, slider synchronization and Undo plus prior crop/transparency operations. Dev typecheck/build and focused lint/diff checks pass. Pane screenshot inspected and numeric percent-unit layout aligned with existing fields.
- Remaining: Corrections ribbon gallery, sharpness and presets, exact installed Mac appearance and comparison of rendering. Overall UI/operation parity remains incomplete.

### Corrections ribbon gallery and contrast rendering

- Added Picture Format → Corrections with a five-by-five brightness/contrast thumbnail grid (-40, -20, 0, 20, 40 percent on each axis), keyboard row navigation, accessible percentage labels and Picture Corrections Options linking to the pane. Choosing a preset applies both adjustments in one undoable edit. Thumbnails reuse the source image and opacity.
- Fixed an existing rendering mismatch: signed contrast adjustment was being used directly as a multiplier (positive adjustments darkened colors; negative adjustments could clamp them to black). Preview now uses a neutral multiplier of one and scales around the midpoint in sRGB. Gallery thumbnails use the same linear mapping. This corrects adjustment direction and neutrality; exact Office transfer curves and combined brightness/contrast behavior still require native comparison.
- References: https://support.microsoft.com/en-gb/office/graphics-visuals/change-the-brightness-contrast-or-sharpness-of-a-picture describes gallery axes and options. https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/52c3867f-20ad-47b1-b0c9-88082bf94676 specifies the signs of brightness/contrast adjustments, but not the exact transfer curve used here.
- Validation: 38 focused preview tests pass, including a new saved/reloaded RGB raster regression for neutral, increased, decreased and minimum contrast, plus positive/negative brightness. Browser test passes for 25 thumbnails, ArrowDown navigation, combined preset selection, Undo, options pane and prior crop/transparency controls. Preview/dev typechecks/builds, focused lint and diff checks pass. Gallery screenshot visually inspected.
- Outstanding: native preset values/metrics verification, Sharpen/Soften section, hover preview on the slide, preservation of other effects in gallery thumbnails, and full Mac UI/operation parity.

### Corrections live preview lifecycle

- Added disposable server-rendered slide previews for Corrections gallery hover and keyboard focus. Uses the existing non-persisting preview endpoint, so picture effects, cropping and surrounding slide content follow the normal renderer. Temporary SVG IDs are namespaced to avoid paint-server collisions.
- Candidate changes cancel the previous request and restore the saved slide; menu close/replacement, pointer exit and editor render clear the preview. Returning to the focused candidate restarts it. Abort and revision/slide/selection checks prevent obsolete responses from installing a preview. Choosing a candidate continues through the existing one-edit commit path.
- Validation: expanded browser test passes for visible preview, unchanged revision/model, pointer leave/re-entry, Escape restoration, delayed-response cancellation, and existing preset selection/Undo, options, crop and transparency checks. Dev typecheck/build, focused lint and diff checks pass.
- Remaining: exact installed Mac hover/focus timing and appearance, Sharpen/Soften, native correction transfer curves and the broader PowerPoint parity scope. This does not establish full Mac fidelity.

### Transparency live preview

- Transparency preset hover and keyboard focus now reuse the disposable server-rendered picture preview. Changing candidates replaces the preview; Escape and menu close restore the saved slide; selection commits through the existing undoable update.
- Validation: expanded picture browser test passes, checking rendered opacity for two candidates, unchanged editor/revision before commitment, Escape restoration, preview removal after commitment, and existing transparency Undo/crop/correction flows. Dev typecheck/build and focused lint pass.
- Native CUA was retried and still reports Transport closed. The Microsoft transparency reference documents the preset gallery and options but does not establish exact Mac hover/focus behavior: https://support.microsoft.com/en-gb/office/graphics-visuals/make-a-picture-transparent . Installed Mac fidelity remains unverified.
- Reset Picture / Sharpen-Soften research: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/f10da8b6-40f3-4527-9c3d-9f5c2e9cff45 shows a14 image processing parameters with a separate original-image relationship and a processed fallback bitmap. Reset and sharpness remain unimplemented; simply adding the parameter would not establish compatible rendering or original-image restoration. Full UI/operation parity remains incomplete.

### Change Picture from File

- Added Picture Format → Change Picture → Picture from File using the existing file chooser. The chosen picture replaces the selected picture's embedded image while retaining shape ID, transform, crop and supported formatting. Captured slide/revision/selection checks prevent applying an asynchronous file choice to a different target.
- Added a picture-replace command and source-image hashes for replay conflict detection. `setShapeImage` now accepts `isolated: true`, allocating both a new media part and relationship so replacements leave other pictures sharing either resource unchanged; default API behavior is preserved.
- Validation: two editor tests pass for same-format/shared-relationship isolation, cross-format/shared-media isolation across slides, retained model properties, saved/reloaded bytes, replay, source conflict detection, malformed input and Undo. The expanded browser test passes for file selection, rendered MIME change, preserved picture model and Undo, alongside prior crop/correction/transparency operations. Existing 12 image-fit tests pass; core/dev builds, dev typecheck, focused lint and diff checks pass.
- Reference: https://support.microsoft.com/en-us/office/graphics-visuals/replace-or-delete-a-picture-in-microsoft-office documents Change Picture. Remaining: exact native differing-aspect-ratio crop behavior, processed/original-image extension handling, context-menu placement, clipboard/stock/photo-browser sources, native file chooser equivalence and full UI/operation parity. This implementation does not establish native Mac equivalence.

### Root typecheck cleanup

- Converted crop aspect-ratio bounds to branded EMU values through the public unit helper, and corrected unit types in related test fixtures. Omitted undefined text-layout options to satisfy exact optional-property typing.
- Validation: root typecheck passes; all 65 tests across seven affected crop/image/preview suites pass; focused lint passes.

### Picture context menu entry points

- Added Change Picture → Picture from File and Format Picture to the picture context menu. The ribbon and context menu share a single file-picker implementation, including size checks, asynchronous target guards and cleanup. Non-picture selections do not expose these commands.
- Validation: expanded picture browser test passes for right-clicking an unselected picture, keyboard submenu expansion, file replacement, retained model properties, Undo, opening/closing Format Picture and absence on a text shape. Existing crop, correction and transparency flows pass after explicitly reselecting the picture/tab. Dev typecheck/build, focused lint and diff checks pass.
- Native reference retry returned Transport closed. Context-menu order, exact labels and current installed Mac visibility rules remain unverified; the prior context-menu gap is now implemented but not certified as exact Mac parity. The complete UI/operation scope remains incomplete.

### Stale SVG resources on image replacement

- Image replacement now removes the old SVG alternate resource from the target blip. Office can otherwise retain a different image resource from the replacement raster. Namespace-based matching supports arbitrary prefixes. Unrelated extension data, including local DPI and empty private extensions, remains unchanged. Shared SVG media and relationships remain available to other pictures.
- Validation: both isolated editor-style and legacy same-format API paths pass saved/reloaded regressions for target-only SVG removal, retained DPI/private extensions/crop and shared original resource preservation. All 14 focused image tests pass; root typecheck/build and focused lint pass.
- References: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/2451f45e-5d77-4661-86d1-0a017fced779 and https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/68e0150d-6a01-4ba5-ac4d-5a18d685229b define the SVG alternate resource. This correction does not resolve a14 original/processed-image effects: preserving those requires regenerating the processed bitmap along with updating its original-image source. Simply deleting processing metadata would lose formatting. Full Mac operation/UI parity remains incomplete.


### Object hyperlink persistence groundwork

- Extended object click actions to groups and added optional ScreenTip text. Click links are inserted before hover links and extensions in nonvisual metadata; removal preserves those other children and group child links.
- Internal slide destinations are validated against the current presentation before modifying an existing link. Foreign and deleted slide handles are rejected without changing XML or relationships.
- Validation: 17 click-action/grouping tests pass, including saved/reloaded group links, escaped ScreenTips, child-link preservation, replacement/removal, metadata ordering and invalid-target atomicity. Root/dev typechecks and builds, focused lint and diff checks pass. Dev build retains existing source-map warnings.
- Reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.nonvisualdrawingproperties?view=openxml-3.0.1 documents group nonvisual properties and hyperlink children. Editor hyperlink commands/dialog, text-range behavior, presentation activation and exact installed Mac comparison remain outstanding. This is persistence groundwork, not a completed hyperlink UI or full parity.

### Undoable object hyperlink commands

- Added serialized object link state to the editor model and an `object-link` command for setting, replacing and removing URLs, internal slide destinations and preset navigation with ScreenTips. Internal destinations use slide part names rather than transient slide positions. Existing journal fingerprints stay unchanged for older operations; link commands additionally fingerprint current links and destination order for conflict detection.
- Added a dedicated object ScreenTip reader so text-run links cannot override the object link's displayed metadata. Object link removal leaves text-run links intact. Invalid destinations, empty URL values, invalid actions, tips and selections are rejected atomically by the editor.
- Validation: all 83 editor tests pass, including two new save/reload, replay, partial/full Undo, conflict and atomic validation regressions. Six core click-action tests, root/dev typechecks/builds, focused lint and diff checks pass. Existing dev source-map warnings remain.
- Native CUA retry still returns Transport closed. Dialog, ribbon/context menu/shortcut entry points, text-range links and link activation remain unfinished; full Mac UI/operation parity is not established.

### Object hyperlink dialog and entry points

- Connected single-object links to Insert → Link, Cmd/Ctrl+K and context-menu edit/removal. The dialog supports Web Page or File address entry, This Document slide/preset destinations, Email Address/Subject and a nested ScreenTip dialog. Draft edits cancel without mutation; save failures restore the dialog; stale selection/revision guards prevent applying to another object. Existing mailto parameters survive subject edits.
- Added tab keyboard navigation and disabled context-menu removal during a pending save. Text editing and table-cell selections remain outside this object-only path pending range-link implementation.
- Validation: dialog and live editor browser tests pass for URL/slide/email/ScreenTip edits, cancellation, removal, retry, ribbon/shortcut/context entry points, Undo and reload persistence. Dev typecheck/build and focused lint pass. Initial integration failures exposed asynchronous save timing and accidental entry into text editing; tests now await save completion and preserve object selection.
- Reference: https://support.microsoft.com/en-au/word/create-or-edit-a-hyperlink-in-office-for-mac (PowerPoint section) documents destination tabs and ScreenTip. Exact installed Mac layout, labels, file Select browser, recent email addresses, text-to-display editing, text-range links and activation remain incomplete/unverified. Full operation/UI parity remains outstanding.

### Character-range link persistence

- Added `setShapeTextRangeClickAction` for URL, internal-slide and preset-navigation links with ScreenTips on selected UTF-16 text ranges. Runs split only at selection boundaries while surrounding links, character formatting and object click actions remain intact. Removal affects the selected range only; empty ranges are no-ops and surrogate-splitting ranges are rejected.
- Shared the object click-link relationship builder with text links and extracted the existing range property mutation traversal. Internal destination validation remains before mutation. Fields retain the existing atomic range semantics.
- Validation: 22 tests across new range-link, range-format, table-range and object-click suites pass, including saved/reloaded Unicode/multi-paragraph links, range removal, internal targets, all navigation presets, invalid-range/foreign-slide rejection and object-link preservation. Four affected editor tests, root/dev typechecks/builds, focused lint and diff checks pass; existing dev source-map warnings remain.
- Editor text selection/dialog connection, link rendering/activation and exact Mac comparison remain pending. This is the persistence layer and does not complete text hyperlink UI or the full parity objective.

### Character-range link editor history

- Added `getShapeTextRangeClickActions`, reading UTF-16 ranges for runs, fields and breaks with resolved destinations and ScreenTips.
- Editor model now exposes `textLinks`; `text-link` commands validate a single shape and nonempty range and preserve object-level actions.
- History fingerprints include text links and slide destinations so replay detects changes to source hyperlinks. Existing object-link fingerprints remain unchanged.
- Verified save/reload, replay, undo, partial removal, source conflicts and invalid/UTF-16-splitting ranges: 3 focused editor tests and 8 core link tests passed. Root/dev typechecks, builds, lint and diff checks passed.
- Remaining: connect selected text to the link dialog and entry points; native Mac visual/interaction comparison remains unavailable. This does not establish full PowerPoint parity.

### Selected-text hyperlink dialog integration

- Insert ribbon Link, Command-K and the text context menu now open hyperlink editing for nonempty shape text selections. Pending text is committed before resolving link ranges; dialog completion/cancel restores selection.
- Existing uniform linked selections populate destination and ScreenTip. Selected text can be changed in Text to Display; replacement and hyperlink form one history command, preserving surrounding text and object actions. Removal preserves display text.
- Browser validation covers insertion, populated edit, cancel/selection restoration, removal, undo/reload, display text replacement with a different UTF-16 length, and a single undo restoring previous text and link. Both hyperlink browser suites passed; 3 focused editor history tests, dev typecheck/build, lint and diff checks passed.
- Remaining: caret-only insertion/editing, table-cell links, links spanning paragraph gaps, link activation and further native Mac comparison. Full UI/operation parity remains incomplete.

### Caret hyperlink insertion and editing

- Link remains available at a text caret. A caret inside an existing hyperlink expands the edit target to the contiguous hyperlink range, while cancel restores the original caret. Matching linked runs separated only by paragraph separators are treated as one link for dialog initialization.
- At an unlinked caret, Text to Display is editable and insertion creates the text and link in one history command. Empty display text defaults to the entered address, email address, or chosen slide label. Collapsed removal and insertion without text/destination remain invalid.
- Browser tests verified caret editing, cancel restoring the caret, address-based insertion, the resulting UTF-16 selection and single-command undo; both hyperlink browser suites passed. Three focused editor history tests, dev build/typecheck, lint and diff checks passed.
- Native reference retry returned `Transport closed`; these changes have not been verified against the installed Mac PowerPoint. Table links, activation and broader parity remain outstanding.

### Table-cell hyperlinks

- Added cell text click-action getters/setters using the same UTF-16 range mutation and relationship resolution as shape text. Preserves surrounding runs, formatting, neighboring cells and table-level object links.
- Editor model carries per-cell link ranges; text-link commands now support cell targets, display text replacement, replay/undo and source-link conflicts. Invalid/merged continuation cells are rejected before mutation.
- Cell text editing supports the Link ribbon command, Command-K and a context-menu entry. Dialog completion restores the cell and text selection; insertion, editing, removal, undo and reload were verified in a live browser, alongside existing table structure/margins operations.
- Focused validation: 11 core tests, 4 editor hyperlink history tests, 2 browser suites passed. Root/dev typechecks, builds, lint and diff checks passed. Native Mac visual/interaction parity and link activation remain unverified/incomplete; full parity is not achieved.

### Slideshow hyperlink activation

- The slide shadow root now handles numeric and preset slide destinations while presenting. Link clicks stop propagation so they do not also trigger the normal slide advance; editing mode suppresses navigation. External links retain separate-window behavior with a protocol allowlist.
- Fixed internal destination rendering to compare slide part names: click-action getters can return a distinct handle for the same slide, so object identity previously dropped valid links. Added all four preset run navigation anchors.
- Object action wrapping now covers pictures, tables and groups as well as ordinary shapes. Object links/ScreenTips are read independently from text links, avoiding an external text link accidentally linking the whole shape.
- Validation: 2 renderer tests cover explicit internal run/object destinations, all presets, independent run/object links, and picture/table/group save/reload. The integrated browser hyperlink suite verifies presentation navigation without the extra advance/end-screen. Root/dev typechecks, preview/dev builds, lint and diff checks passed.
- Remaining: table-cell/field link rendering, edit-mode Open Hyperlink affordances, native Mac interaction/visual comparison and the broader parity backlog. Native automation was unavailable in the last attempt; no claim of full parity is made.

### Cell and field hyperlink rendering

- Table cell paragraph models now retain hyperlink destinations and ScreenTips using UTF-16 source offsets, including paragraph separators and surrogate pairs. Browser cell text shares shape-text link styling and anchor rendering.
- Shape text now resolves links through the range reader, allowing linked fields to retain their actions while displayed slide-number fields update independently of cached text length.
- Pure SVG output now emits anchors and ScreenTips around linked spans; token grouping preserves separate ScreenTips instead of merging them.
- Validation: 57 tests across click actions, slide numbers, table text and existing SVG text modes/rendering passed. Added Unicode/multiple-paragraph table link save/reload coverage for both text render modes and linked live slide-number fields. Browser testing physically clicks a cell link, observes its popup destination using an intercepted response, and verifies that the presentation stays on its original slide. Root/dev typechecks, preview/dev builds, lint and diff checks passed.
- Mac reference access retried and returned `Transport closed`; no native comparison was possible. Remaining work includes edit-mode Open Hyperlink affordances and the broader PowerPoint parity backlog. Full parity remains incomplete.

### Open Hyperlink in the editor

- Added Open Hyperlink to the object, shape-text and table-cell text context menus when the current selection resolves to a link. Reused contiguous range resolution from the link dialog to keep caret and multi-run selections consistent.
- External destinations open separately with a protocol allowlist. Internal/preset destinations select the target slide in editing mode; unavailable destinations leave the current slide unchanged. Text edits are committed before internal navigation.
- Pending typing is committed and selection restored before displaying the text context menu, so link resolution uses current UTF-16 offsets rather than stale saved ranges.
- Browser coverage verifies object navigation without entering slideshow and a cell external-link popup after pending typing, plus existing insert/edit/remove/undo/reload/slideshow checks. Dev typecheck/build, lint and diff checks passed.
- Microsoft reference: https://support.microsoft.com/en-us/powerpoint/training/add-a-hyperlink-to-a-slide documents normal-view testing through Open Hyperlink. Native Mac menu placement/appearance remains unverified because the last automation attempt returned Transport closed. The full parity goal remains incomplete.

### Last-viewed and end-show actions

- Added native `lastslideviewed` and `endshow` click actions to object, text and cell readers/writers and editor command validation. Imported actions remain visible in the hyperlink dialog without silently changing their destination.
- Slideshow links now navigate to the previously viewed slide or exit presentation mode. History resets on each presentation start; a last-viewed action without history does nothing. These presentation-only actions do not accidentally select the preceding slide in editing mode.
- Validation: 16 focused core/render tests passed, including save/reload and exact action XML for objects, text and cells in both renderer modes. The browser hyperlink suite verifies empty-history behavior, back-and-forth history navigation (including returning to a later-numbered slide), and end-show without an extra advance. Root/dev typechecks, preview/dev builds, lint and diff checks passed.
- Microsoft reference: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slideshowview.lastslideviewed documents the immediately previously viewed slide. Native Mac comparison remains unavailable in the last attempt (`Transport closed`). Action-button gallery/settings and broader UI parity remain incomplete.

### Action-button gallery and navigation defaults

- Added all 12 action-button presets at the end of the Shapes gallery. Navigation buttons carry previous/next/first/last/last-viewed actions through insertion, export and journal replay; non-navigation buttons remain unassigned. Crop to Shape excludes the new category.
- Object hyperlinks now use their SVG bounding box for pointer hit testing. Browser testing uncovered that the hollow center of the forward button previously let clicks reach the slide background; real center clicks now activate the link.
- Validation: all gallery geometries/export/replay and 12 button action states passed the editor test; six click-action renderer tests passed. The integrated browser hyperlink test passed gallery insertion, reload and physical button activation alongside existing hyperlink/history/end-show coverage. Root/dev typechecks, preview/dev builds and lint/diff checks passed.
- Reference: https://support.microsoft.com/en-us/powerpoint/add-commands-to-your-presentation-with-action-buttons. The required post-insertion Action Settings dialog (click/hover, destinations, sound and other actions) remains to be implemented. This gallery increment does not complete action-button parity.
- Native Mac access retried and returned `Transport closed`; appearance, ordering and exact default behavior still require native comparison. Full UI parity remains incomplete.
- The broad browser editor regression also passed after updating gallery count/End-key expectations and limiting duplicated ribbon selectors to visible controls.

### Click Action Settings dialog

- Action-button insertion now opens Action Settings after the new shape is saved and selected. Added Insert > Action and an object context-menu entry; selection/revision guards prevent applying a stale dialog to another object.
- Supports None and Hyperlink to, including the six slideshow destinations plus nested Slide and URL destination dialogs. Existing ScreenTips are preserved; Cancel leaves the current action intact, and identical confirmation avoids an extra history entry. Save errors keep the dialog available for review.
- Browser validation passed automatic opening/default selection, cancel, ribbon/context reopening, nested URL destination, action removal, undo and the existing slideshow/link workflow. Dev typecheck/build, lint and diff checks passed.
- Still incomplete: Mouse Over settings, program/macro/OLE actions, sound, custom shows and file destinations. Native dialog appearance and exact controls have not been verified because Mac automation remains unavailable. Full PowerPoint UI/operation parity is not achieved.

### Mouse Over action persistence

- Added object hover action readers/writers and ScreenTip access using `a:hlinkHover`, independently of click actions. Supports URL, explicit slide and the six existing slideshow destinations. Insertion preserves click/hover/extension XML order; replacement/removal changes only the requested trigger.
- Validation: 18 tests across object click/hover and text/cell links passed. Hover coverage includes save/reload, groups and tables, child-action independence, replacement/removal, ScreenTips and foreign-slide rejection without mutation. Typecheck, lint and diff checks passed.
- This is persistence support only: editor commands, Mouse Over settings and slideshow execution remain to be connected. Native PowerPoint access retried and returned `Transport closed`; full UI/operation parity remains incomplete.

### Mouse Over settings in the editor

- Added Mouse Click / Mouse Over tabs with separate draft values and keyboard tab switching to Action Settings. Cancel discards both drafts; confirmation saves both triggers as one history operation. The editor model exposes hover actions independently of click links.
- Added atomic validation of both destinations before mutation, save/reload and journal replay support, and hover-source conflict detection. Existing object-link journal fingerprints stay unchanged.
- Validation: three focused editor tests passed, including invalid hover destinations leaving the click action untouched. The integrated browser hyperlink suite passed tab switching/draft retention and one-step undo of both triggers alongside existing dialog/link checks. Dev typecheck/build, lint and diff checks passed.
- Remaining: slideshow hover execution and renderer metadata, broader action types, and native Mac dialog fidelity. Full PowerPoint parity is still incomplete.

### Slideshow Mouse Over execution

- Renderer emits independent hover destinations on object wrappers, including hover-only shapes, while preserving click destinations. The slideshow runs internal destinations/end-show on pointer entry and ignores editing mode, touch input, and transitions between descendants of the same action wrapper.
- Pointer coordinates suppress repeated action execution when a slide replacement puts another action under the stationary pointer. Moving the pointer re-arms hover actions. External destinations use the existing protocol allowlist and a separate window; browser popup policy can still restrict hover-initiated windows.
- Validation: eight renderer tests passed across SVG/foreignObject modes, including saved hover-only actions and independent click/hover URLs. The integrated browser suite passed saved/reloaded hover navigation and end-show with real pointer movement, edit/touch suppression, and existing click/settings/undo checks. Root/dev typechecks, preview/dev builds, lint and diff checks passed.
- Native Mac hover behavior and visual fidelity remain unverified. External hover popup behavior and broader action types remain outstanding; full parity is not achieved.

### Preserve independent action metadata

- Updating a supported destination now retains its sound child, end-sound/history flags, extension children and namespace declarations; replaced action/id/tooltip and stale invalidUrl metadata are not copied. Explicit removal still removes the complete link.
- The combined Action Settings command now writes only triggers whose exposed values changed. Editing one trigger therefore preserves an unrecognized macro/action on the other trigger instead of deleting it as an apparent null link.
- Validation: 16 core action tests and four editor action tests passed, including save/reload of independent metadata and an imported macro on either unchanged trigger. Root/dev typechecks, root build, lint and diff checks passed.
- Reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hyperlinkonclick?view=openxml-3.0.1 documents hyperlink sound, flags and extension children. Native reference retry still returned `Transport closed`. Unsupported action editing/execution and sound controls remain unimplemented; preserving them does not establish full action-setting parity.

### Embedded action sound persistence

- Added independent click/hover sound read/write APIs for embedded WAV data, sound names and stop-previous-sounds flags. Sound-only actions can be stored without a navigation destination. Setting/removing sound preserves the destination and other trigger.
- Identical audio bytes reuse the media part and slide relationship. Replacing a sound creates/reuses another part instead of mutating shared bytes; getters and setters copy byte arrays. Removed media is retained for other references. Invalid inputs are rejected before mutation.
- Validation: four new sound tests pass for save/reload, package invariants, independent triggers, destination retention, shared-sound isolation, removal, invalid input atomicity and slide duplication. The prior 16 click/hover tests also passed. Root typecheck/build, lint and diff checks passed.
- DrawingML reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hyperlinksound?view=openxml-2.17.1. Editor sound controls, built-in sounds and slideshow playback are not connected yet. Native fidelity and full parity remain incomplete.

### Action sound controls and editor history

- Action Settings now exposes independent click/hover WAV file selection and Stop Previous Sound controls. Tab switching retains separate drafts; confirmation saves sounds and destinations together. Cancel leaves the document untouched.
- Editor models include sound names/data and stop flags. Both triggers are validated before mutation; journal replay and undo restore their settings. Removing navigation preserves an existing sound, and unchanged audio metadata is not rewritten.
- Validation: three focused editor tests passed for save/reload, replay, undo, destination removal and invalid-data atomicity. The integrated browser suite passed WAV selection, independent tab drafts, combined save and undo. Dev typecheck/build, lint and diff checks passed.
- Remaining: slideshow audio playback, built-in sound choices and native Mac visual/behavioral comparison. These settings do not yet establish full PowerPoint parity.

### Slideshow action sound playback

- Object wrappers now carry separate embedded WAV and stop-sound metadata for click and hover, including sound-only shapes with no navigation destination.
- Slideshow actions start sound before navigation, retain ongoing sounds across slide changes, stop all active action sounds for Stop Previous Sound, and release playback on show exit. Edit mode suppresses playback; sound-only clicks consume the click so they do not advance the slide. End/error/rejected playback is removed from the active set.
- Validation: ten renderer tests passed, including saved/reloaded sound-only wrappers in both text rendering modes. The integrated browser suite passed edit suppression, click playback routing without slide advance, hover stop of two active sounds, and show-exit cleanup using media lifecycle instrumentation. A separate real browser Audio playback reached the ended event for the WAV fixture. Root/dev typechecks, preview/dev builds, lint and diff checks passed.
- Native reference access again returned Transport closed. Audible output and exact Mac sound concurrency behavior are not verified; browser autoplay policy can reject hover-initiated playback. Built-in sound choices and broader action types remain incomplete. Full UI/operation parity remains unachieved.

### Retain sound choices while editing

- Fixed mismatches between the displayed sound and the saved value when disabling/re-enabling Play sound, or selecting Stop Previous Sound and returning to the previously selected WAV. Each trigger keeps its selected sound separately from its enabled state; tab changes preserve both.
- Validation: the integrated browser suite passed disable/tab-switch/re-enable and stop/WAV reselection followed by exact saved-data assertions. Reopening retained the saved selection; an invalid WAV left it intact; Cancel after disabling sound preserved the document revision and saved sound. Existing playback, navigation and undo checks also passed. Dev typecheck/build, lint and diff checks passed.
- Native checkbox behavior and visual appearance still require Mac comparison. This fixes internal UI/data consistency and does not prove full PowerPoint parity.

### Custom show document model

- Added getCustomShows/setCustomShows APIs for named, stable-ID slide sequences in standard p:custShowLst markup. Sequences can repeat slides, remain independent of deck order and preserve metadata on surviving show IDs. Replacement validates IDs and slide ownership before modifying the package.
- Removing a slide now removes all its occurrences from custom shows while retaining empty show definitions. Removal resolves canonical internal relationship targets and rejects a handle from another presentation.
- Validation: 15 tests across five core suites passed, covering custom show save/reload, XML schema validation when available, repeated slides, reorder independence, retained extension metadata, invalid-input atomicity, canonical relationship deletion and existing section/sort/delete operations. Root typecheck/build, lint and diff checks passed.
- References: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.customshow?view=openxml-3.0.1 and the repository's ECMA-376 Transitional pml.xsd CT_CustomShow/CT_Presentation definitions.
- Remaining: editor commands/history, Custom Shows and Define Custom Show dialogs, ribbon entry, custom sequence playback, hyperlinks to shows and native Mac comparison. Document support alone does not establish UI or operation parity.

### Custom show editor, dialog, and playback

- Added custom show definitions to the editor model and a revision-checked history command. Definitions save into the PPTX, replay, and undo; missing-slide references and changed-source replay fail without applying the edit.
- Slide Show > Custom Show lists defined shows and opens Custom Slide Show. The manager supports create, edit, copy, delete, and Start Show. Define Custom Show supports name, multi-select Add, Remove, and ordering. Cancel discards definition changes; each confirmed manager operation persists separately. Concurrent document changes report an error instead of overwriting newer definitions.
- Playback tracks ordered slide keys and an independent position, including repeated slides. Next/Previous, Home/End, Shift-Space, navigation actions, end screen, and exit use the custom sequence; starting a normal show clears it. Explicit slide destinations remain available.
- Validation: dev typecheck/build, targeted editor save/replay/undo/conflict test, custom show browser integration, existing hyperlink/action browser integration, targeted lint, and diff whitespace checks passed.
- Still incomplete: custom-show hyperlink destinations and Show and return; full native Mac visual/behavior comparison. Native UI transport remains unavailable. The dialog follows the documented Mac workflow, but exact native parity is not claimed.

### Custom show hyperlinks and Show and return

- Implemented the MS-OI29500 custom-show action URI (`ppaction://customshow?id=...&return=true`) for shape click, mouse-over, text runs/ranges, and the shared table text reader. Destination validation rejects absent shows before changing either action. Readers preserve numeric IDs and the return setting through PPTX save/reload.
- Action Settings now provides Custom Show with a destination picker and Show and return. Insert/Edit Hyperlink lists custom shows under This Document and exposes the return checkbox. Existing custom-show links reopen with their selections intact.
- The renderer emits custom-show action destinations. Slideshow playback can enter a show from a link and restore the originating slide, parent sequence position, and last-viewed state on natural completion. Nested returning shows retain a return stack. Non-returning links finish at the end screen; starting normal playback or exiting clears return state.
- Validation: root/preview/dev typecheck and builds; 27 tests across custom-show, action rendering, hover, and text-range suites; 2 targeted editor save/replay/undo/atomicity tests; browser custom-show manager/link/nested-return/non-return playback plus existing hyperlink and link-dialog integrations; targeted lint and diff checks passed.
- Native exactness remains unverified. In particular, compare Escape/explicit End Show behavior inside returning custom shows against installed PowerPoint; these currently exit the full presentation. Mac UI transport remains unavailable. Overall all-operation UI parity remains incomplete.
- Reference: https://learn.microsoft.com/ja-jp/openspecs/office_standards/ms-oi29500/a65b76db-6abc-4989-8cd1-baa9a3500f6f

### Presentation-wide slideshow settings storage and history

- Added public `getSlideShowProperties` / `setSlideShowProperties` APIs for show mode (present / browse / kiosk), all / range / custom show selection, loop, narration, animation, and timings. Reads OOXML defaults, follows the properties relationship, preserves unrelated print/web settings, pen color and extension children, and creates missing properties without overwriting an occupied conventional part name.
- Invalid values, missing custom shows, external properties relationships and malformed/wrong-root XML reject before package mutation. Core tests cover save/reload, schema validation, arbitrary related part locations, preservation and rejection atomicity.
- Added `EditorModel.showProperties` and the `show-properties` edit command with history replay/undo and conflict detection against original settings, slide identities and custom shows.
- Fixed transition introspection: timing-only `p:transition` now returns `effect: 'none'` with its timing instead of null; sound actions/extensions are not treated as effects; XML `true` / `false` booleans are read correctly.
- Corrected the prior custom-show render test's option from `textMode` to `textLayout`, so it actually exercises both SVG and foreignObject modes.
- Validation: root typecheck/build, dev typecheck/build, focused editor history tests (2), and 31 core/preview/transition tests passed; lint and diff whitespace checks passed. Native Mac comparison remains unavailable because the UI transport reports `Transport closed`.
- This increment is storage/history infrastructure only. Set Up Show dialog, applying settings to playback, automatic advance, kiosk behavior, narration and animation playback remain pending. No claim of native UI or complete slideshow parity.

### Apply slideshow settings to playback

- Playback now resolves all visible slides, a configured inclusive range, or a persisted custom show into an ordered sequence. Play from Start and its keyboard shortcuts explicitly use the configured sequence's first item; Play from Current uses the current slide's position when present. Repeated custom slides remain distinct positions.
- Loop restarts the sequence after its last slide; nested Show and return still restores its parent before considering a loop. Browse mode starts without requesting browser fullscreen. Fixed the slideshow context menu's Previous enablement to use sequence position.
- Editor slide models now expose parsed transition options. Auto advance uses stored milliseconds when Use timings is enabled, honors click advance flags for stage clicks, and cancels timers on exit or end screen. Long durations are chunked within the browser timeout limit. Manual navigation reschedules the active slide's timer; manual settings suppress auto advance.
- Browser coverage verifies range endpoints, loop restart, saved custom-show selection after reload, repeated sequence order, timing progression, manual override, disabled stage-click advance and timer cancellation on exit. Timing is injected into the browser model in this test to isolate playback; core transition tests separately prove persisted timing extraction. Existing custom-show return/manager and hyperlink scenarios passed.
- Validation: dev typecheck/build, lint and diff checks passed. New configured-show browser test and existing custom-show test passed after fixing the custom-show Play from Start regression discovered by the new assertion.
- Remaining: Set Up Show UI, kiosk input restrictions/restart, browse scrollbar behavior, animation/narration playback, exact native interaction and visual verification. Timing playback does not yet wait for animated effects or media, whose playback is not implemented.

### Set Up Show dialog and kiosk navigation restrictions

- Added Slide Show ribbon > Set Up Slide Show, with a modal restoring show mode, all/range/custom selection, loop, narration/animation flags and manual/timing advance from the presentation. Range fields enable only for range selection; missing custom shows disable selection; invalid reversed ranges keep the dialog open. Preserves existing browse scrollbar/kiosk restart values when editing other fields.
- OK uses the existing history command and revision guard; Cancel/Escape leave the presentation unchanged. Kiosk mode checks and disables the loop control, matching Microsoft's documented automatic loop selection. Narration/animation flags are persisted; those playback engines remain unimplemented.
- Kiosk playback now suppresses background clicks, ordinary navigation keys, the generic context menu and on-screen navigation controls; action links, timer navigation and Escape remain available. Inactivity restart and complete native kiosk behavior are still pending.
- Browser tests cover dialog enablement, invalid ranges, cancel, save/reopen, undo, kiosk loop control and kiosk navigation suppression/Escape, alongside prior range/custom/timing playback cases. Dev typecheck/build, lint and diff checks passed. Existing custom-show test timed out once at a shape-link selection, then passed on rerun; no deterministic regression established.
- Visually inspected a 1600×1000 browser screenshot. Fixed inherited generic dialog CSS that collapsed the settings columns and range inputs; re-render verified readable two-column layout and footer alignment. This is browser visual verification, not evidence of exact installed Mac PowerPoint geometry. Native UI tool retry still returned `Transport closed`.
- Reference: https://support.microsoft.com/en-us/powerpoint/training/create-a-self-running-presentation (show modes and automatic kiosk loop selection). Full native Set Up Show controls and exact visuals remain to be compared; pen color/monitor settings are not exposed by this dialog yet.

### Kiosk restart clock

- Implemented the persisted kiosk restart interval with a monotonic clock, bounded browser timeout chunks, cancellation on exit, and reset at automatic sequence looping. Restart clears nested return state and returns to the original show's first slide, including configured custom shows. A custom show entered without return becomes the new restart sequence.
- Corrected the previous entry's “inactivity restart” terminology: the OOXML restart interval is elapsed time since presentation start/restart, not time since the last pointer or keyboard input. Reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.kioskslidemode?view=openxml-3.0.1 . Zero intervals use a minimum 16 ms scheduling delay to avoid a busy loop.
- Browser clock tests verify input does not postpone restart, custom-show start destinations, nested Show and return retaining the original deadline/destination, timer cancellation after exit, and reset at automatic looping. The configured-show and custom-show browser suites both passed. Dev typecheck/build, focused lint and diff checks passed.
- Native Mac comparison remains unavailable; this verifies browser behavior against the serialized setting, not exact native UI parity. Browse scrollbars, full Set Up Show controls, narration/animation/media playback and the wider parity gaps remain open.

### Browse-mode scrollbar

- Added Show scrollbar to Set Up Show with browse-only enablement, restoration and history-backed saving. Playback now honors this persisted flag with a browser-native vertical scrollbar and reserves its width from the fitted slide area. It is hidden outside enabled browse playback.
- Scroll positions follow the configured sequence, including repeated custom-show entries. Keyboard/action slide changes synchronize the position; scrolling selects the corresponding sequence position. Resize recalculates the scroll range.
- Validation: dev typecheck/build and both configured-show/custom-show browser suites passed. Added coverage for setting restoration, disabled controls, enabled/disabled playback, real wheel input, repeated sequence positions and keyboard synchronization. Inspected screenshots of the dialog and browse playback; the OS/browser controls native thumb visibility. Lint and whitespace checks passed.
- Reference: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slideshowsettings.showscrollbar documents slide browsing via the scrollbar. This establishes behavior, not exact Mac geometry. Native UI retry again failed with Transport closed; installed Mac comparison and the broader outstanding parity requirements remain open.

### Transition advance timing editing

- Added Transitions ribbon > Advance Slide controls for On Mouse Click and After, with minutes:seconds input and validation. Editing persists through history and reload and feeds the existing playback timing model. The Transitions ribbon remains incomplete: effect gallery, duration, sound, preview and Apply to All are still pending.
- Added public `setSlideAdvanceTiming`, which updates only `advClick` and `advTm`, preserving visual effects, sound actions, extensions and other transition attributes. Null delay removes automatic advance; invalid values fail before mutation. Added editor `slide-advance` command using normal slide conflict detection.
- Browser testing caught a pending-save checkbox flicker; pending timing state now keeps controls stable and disabled until save finishes. Screenshot inspection caught vertical clipping; reduced timing-row spacing and added a geometry assertion, then visually rechecked the corrected ribbon.
- Validation: root/dev typechecks and builds passed, 17 core transition/introspection tests passed, editor history persistence/replay/undo/conflict test passed, and configured-show browser test passed including input validation, save/reload, undo and clearing automatic advance. Lint and diff whitespace checks passed.
- Reference: https://support.microsoft.com/en-us/PowerPoint/set-the-timing-and-speed-of-a-transition documents the Transitions timing controls. Exact installed Mac placement remains unverified. Playback still lacks transition effects and waiting for animations/media before advancement; this increment does not establish full transition or PowerPoint parity.

### Transition Apply To All

- Added Transitions > Apply To All and a history-backed editor command. Copies the complete source transition, including visual effect, timing, sound actions and extension metadata; a source with no transition removes transitions on other slides. Destination slides retain their other content.
- Added public `applySlideTransitionToAll`. Relationship references are remapped per destination, matching relationships are reused, and inherited namespace declarations are retained. All copies are prepared and checked before package mutations. Foreign source slides and missing referenced relationships/parts are rejected.
- Validation: root/dev typechecks and builds passed; 8 core tests and 2 focused editor history tests passed; configured-show browser suite passed with Apply To All and Undo coverage. Sound relationship remapping, repeated application without duplicate relationships, reload, removal and atomic missing-relationship failures are covered. Focused lint and whitespace checks passed. Inspected `/tmp/pptx-transition-apply-all.png` for ribbon clipping.
- Reference: https://support.microsoft.com/en-us/powerpoint/training/add-change-or-remove-transitions-between-slides describes Apply To All. Native Mac appearance remains unverified due to the unavailable UI transport. Effect gallery, duration, sound selection, preview, transition playback and wider PowerPoint parity remain outstanding.

### Transition effect editing

- Added ribbon tiles for None, Fade, Push, Wipe, Split and Cover, plus a More menu for the complete existing core effect catalog. Effect Options exposes supported directions, split orientations and fade/cut through-black choices. Selection restores after reload and edits use normal history/Undo.
- Added public `setSlideTransitionEffect` and editor `transition-effect` command. Effect-only changes retain advance timing, speed, sound actions and extension metadata. Choices validate before mutation; None removes the visual effect while retaining other settings.
- Validation: root/dev typechecks and builds, 9 focused core tests, an editor persistence/replay/undo test and configured-show browser suite passed. Browser coverage includes selection state, direction, reload, more-menu choice, option disablement and None/Undo. Focused lint and whitespace checks passed. Screenshot `/tmp/pptx-transition-gallery.png` inspected; temporary capture removed.
- Reference: https://support.microsoft.com/en-us/powerpoint/training/add-change-or-remove-transitions-between-slides (Mac instructions for gallery and Effect Options). Native retry still returns Transport closed. The current More menu, icons, catalog and placement are not verified as identical to installed Mac PowerPoint. Modern effects such as Morph, full gallery layout, automatic/manual preview, duration and sound controls, effect playback and broader parity requirements remain incomplete.

### Transition duration and compatibility containers

- Added a Duration input (seconds) and history-backed `transition-duration` command. Public `setSlideTransitionDuration` and `TransitionOptions.durationMs` read/write Office 2010 `p14:dur` milliseconds with local namespace/markup-compatibility declarations. Duration changes preserve effect, sound, speed and advance settings. Invalid/noninteger/out-of-range API values fail before mutation; UI validates seconds before saving.
- Official documentation exposed a gap in the previous transition implementation: PowerPoint can wrap transitions in `mc:AlternateContent`. Transition reading now locates Choice/Fallback containers; effect/timing/duration edits update their transition branches, Apply To All copies the full container, and removal removes that container. Unknown modern effect decoding and capability-based Choice selection remain incomplete.
- Validation: root/dev typechecks and builds passed; 11 focused core tests and 3 editor history tests passed; configured-show browser suite passed with duration input, invalid input, reload and preservation across effect changes. Core tests cover wrapped transitions, both branch updates, full-container copying/removal, namespace round-trip and validation atomicity. Tests caught and fixed a missing `mc` declaration and fractional-millisecond acceptance. Focused lint and whitespace checks passed. Inspected `/tmp/pptx-transition-duration.png`; removed temporary capture.
- Reference: https://learn.microsoft.com/en-us/office/open-xml/presentation/how-to-add-transitions-between-slides-in-a-presentation documents `p14:dur` milliseconds and the Choice/Fallback structure. Duration is blank when absent in source XML; native effect-specific defaults still need verification. This does not implement transition playback/preview or establish exact installed Mac UI parity. Sound controls, modern effects, full gallery and the broader outstanding requirements remain open.

### Transition preview and initial playback

- Added the Transitions Preview command and automatic preview after changing an effect or its options. Slideshow entry/navigation now plays supported effects using the stored duration, and schedules automatic slide advance after playback completes. Navigation, exit, screen blanking and changed SVG content cancel active playback. Separate shadow roots isolate duplicate SVG IDs in outgoing/incoming layers.
- Initial playback covers fade (including through black), push, wipe, split, cover, pull, cut, zoom, circle and diamond. Unsupported effects switch immediately. Missing duration currently uses a provisional 500 ms default. Exact native geometry, easing, effect-specific defaults and semantics remain unverified; Plus and the remaining effects still need playback implementations.
- Validation: dev typecheck/build, focused lint and whitespace checks passed. Browser tests cover halfway push geometry, overlay bounds, SVG isolation, cancellation without stale completion, preview cleanup, unsupported-effect fallback, and existing configured-show ranges/looping/timers. The configured-show UI test also exercises Preview with a saved 1.75-second Push transition without changing the selected slide.
- This is partial playback support, not full Mac PowerPoint parity. Native reference automation remains unavailable from earlier Transport closed errors; gallery layout, modern effects, transition sound controls and the broader outstanding operations remain open.

### Plus and Blinds playback

- Extended transition playback to Plus and both Blinds directions. Plus expands a twelve-point cross; Blinds reveals eight isolated strips, with horizontal/vertical movement from the saved direction. Both use stored duration and the existing cancellation/completion lifecycle.
- Browser validation passed for half-progress clipping in both Blinds directions, Plus cross geometry and completion cleanup. Inspected `/tmp/pptx-transition-plus.png` at half progress. Dev typecheck/build, focused lint and whitespace checks passed.
- Reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.blindstransition?view=openxml-3.0.1 describes directional bar wipes and explicitly leaves exact rendering to the application; https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.plustransition?view=openxml-3.0.1 identifies the Plus transition. Bar count and exact growth remain provisional until native comparison. Native retry again returned Transport closed. Full Mac UI/operation parity remains incomplete.

### Wheel effect options and playback

- Added `TransitionOptions.spokes` reading/writing and effect-only editing, with unsigned-integer validation before mutation. Wheel Effect Options offers 1, 2, 3, 4 and 8 spokes. Stored spokes survive save/reload and Apply To All; timing is retained on effect edits.
- Wheel playback now reveals clockwise radial sectors with the selected spoke count. Fixed-topology polygon keyframes interpolate angular progress. Default is four spokes; other unsigned values are preserved in files but currently switch immediately during playback rather than allocating unbounded polygon data.
- Core validation passed (six transition tests), including invalid-spoke atomicity, save/reload, timing retention, Apply To All and stripping wheel-only attributes from other effects. Root/dev typechecks and builds, focused lint and whitespace checks passed. Browser coverage adds all five playable spoke counts and Wheel selection → 3 Spokes → reload → Preview.
- Reference: https://learn.microsoft.com/en-au/dotnet/api/documentformat.openxml.presentation.wheeltransition?view=openxml-2.9.0 documents clockwise radial wipes and spoke examples. Native effect geometry and exact Mac menu appearance remain unverified, and the broader full-parity goal remains incomplete.

### Dissolve, Random Bars and Checkerboard playback

- Added randomized square reveals for Dissolve, randomized horizontal/vertical bars for Random Bars and directional alternating-cell wipes for Checkerboard. These now use the saved duration in preview/slideshow instead of switching immediately. A ResizeObserver rebuilds pixel-based clip paths when the slide display changes size and is disconnected on cancellation.
- Browser checks sample real clipped hit regions at start, halfway and end, verifying zero/half/full cell coverage for all five effect/direction combinations. Checked resizing and existing cancellation/cleanup paths. Configured-show suite passed; its navigation/timing cases now explicitly clear the visual effect to isolate their timer expectations. Dev typecheck/build, lint and whitespace checks passed.
- Inspected `/tmp/pptx-dissolve-midpoint.png`. This caught fine artifacts caused by zero-width connecting edges in compound polygons; independent path subpaths removed them. Temporary screenshot capture was removed from the test.
- References: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.dissolvetransition?view=openxml-3.0.1 ; https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.randombartransition?view=openxml-3.0.1 ; https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.checkertransition?view=openxml-3.0.1 . Exact native cell counts, default durations and interpolation remain unverified. This does not establish full Mac PowerPoint UI/operation parity; remaining effects and broader outstanding operations still require work.

### Comb and Wedge playback

- Added Comb playback using eight alternating strips entering from opposite edges, with horizontal/vertical direction support. Added Wedge playback using two radial edges opening from the top in opposite directions. Both use stored duration and the existing preview, cancellation and completion lifecycle.
- Browser tests passed for actual half-progress hit regions in both Comb directions and both sides of the Wedge reveal. The configured-show browser suite also passed. Dev typecheck/build and focused lint passed.
- References: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.combtransition?view=openxml-3.0.1 ; https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.wedgetransition?view=openxml-3.0.1 . Strip count and precise native geometry/timing remain provisional pending comparison with installed Mac PowerPoint. Strips, Newsflash and Random playback, modern effects, exact gallery UI and the broader full-parity requirements remain outstanding.

### Newsflash and Random playback

- Added Newsflash playback: the incoming slide grows around its center while turning counter-clockwise over the previous slide. One revolution and linear interpolation are provisional pending native comparison.
- Random now chooses a supported effect on each playback, retaining the saved duration without changing the document's Random setting. Effect-specific parameters reset to defaults when resolving Random. The current candidate set is the implemented catalog, including Cut; exact native candidates and weighting remain unverified.
- Browser tests passed for Newsflash start/intermediate/end transforms and deterministic Random selections across two plays, including frozen input settings and duration retention. Dev typecheck/build and focused lint passed.
- References: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.newsflashtransition?view=openxml-3.0.1 describes counter-clockwise growth; https://learn.microsoft.com/th-th/dotnet/api/documentformat.openxml.presentation.randomtransition?view=openxml-2.8.1 describes choosing from available renderer effects. Native reference retry returned Transport closed. Strips playback, modern effects, exact Mac gallery/controls and wider full-operation parity remain incomplete.

### Strips playback

- Added staggered strip wipes for all four corner directions. Eight horizontal bands reveal from the appropriate edge with staggered starts, completing within the stored duration. Random can now select Strips too. The existing core gallery catalog now has playback implementations; this is not the complete native Mac catalog.
- Browser tests passed for actual clipped hit regions at start, midpoint and finish for all four directions, including reversing the band order and horizontal reveal edge, plus cancellation cleanup. An initial test expectation incorrectly used the same midpoint row for upward and downward wipes; corrected the sample to follow direction.
- Reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.stripstransition?view=openxml-3.0.1 describes staggered bar wipes. Band count and precise geometry/timing remain provisional until native comparison. Modern effects, exact Mac gallery and control layout, sound controls and broader operation parity remain incomplete.

### Transition option selection and omitted directions

- Effect Options now marks the active option, using stored values or ECMA-376 defaults when attributes are absent. Defaults cover direction, Split orientation, through-black and Wheel spokes. Unknown explicit values are not falsely marked as a standard choice.
- Inspecting the bundled authoritative schema (`references/ecma-376-5th/ECMA-376/OfficeOpenXML-XMLSchema-Transitional/pml.xsd`, lines 38–64) exposed two playback bugs: omitted Zoom direction must be `out`, and omitted Strips direction must be `lu`. Corrected both, also affecting effects resolved by Random.
- Added browser checks for Push default selection, saved Push/Wheel selection after reload, and omitted-versus-explicit default playback geometry for Zoom and Strips. Dev typecheck/build passed. Native Mac menu appearance and effect timing remain unverified; broad full-operation parity remains incomplete.

### Transition sound persistence

- Added public `SlideTransitionSound`, `getSlideTransitionSound` and `setSlideTransitionSound` APIs for no sound, stopping previous sound, or embedded WAV playback with looping. Audio bytes are copied, duplicate media/relationships are reused, and replaced media remains available to other consumers.
- Sound-only changes preserve visual effects, duration and advance settings, insert sound actions before extensions and update every existing Choice/Fallback transition branch. Invalid WAV input is rejected before package mutation; missing/external audio is not fetched.
- Eight focused transition tests passed, including sound save/reload, Apply To All, duplicate relationship prevention, detached returned bytes, stop/removal, branch updates, extension retention and atomic invalid input. Root typecheck/build, focused lint and whitespace checks passed.
- Authoritative structure: bundled `references/ecma-376-5th/ECMA-376/OfficeOpenXML-XMLSchema-Transitional/pml.xsd`, CT_TransitionStartSoundAction and CT_TransitionSoundAction. This increment implements persistence only; the ribbon sound selector, editor history integration and runtime playback remain to be connected. Full Mac UI/operation parity remains incomplete.

### Transition sound editor history

- Added serialized transition sound data to each editor slide and a `transition-sound` command for WAV playback/looping, stop and none. Commands validate kind, loop and bounded WAV/base64 data before mutation and use normal slide conflict detection.
- Focused editor test passed for sound save/reload, history replay, Undo, stop/removal, retaining visual settings, conflicting slide changes and atomic rejection of invalid data. Dev typecheck/build, focused lint and whitespace checks passed.
- Ribbon sound selection and runtime playback still need to consume this model/command. No native Mac UI parity claim is established by this backend integration.

### Transition sound ribbon controls

Connected the transition sound selector to the editor history API: No Sound,
Stop Previous Sound, Other Sound (WAV upload), embedded sound name, and Loop Until
Next Sound. Uploads validate WAV signatures and the 20 MB limit. Async upload is
bound to the originating slide; saving retains the pending checkbox state.
Browser coverage verifies upload, loop, reload persistence, stop and removal.
Dev typecheck/build, focused lint, browser show-properties test and diff checks pass.
Playback is not yet connected; exact Mac ribbon positioning and built-in sound
choices remain unverified. This does not establish full PowerPoint parity.

### Transition sound playback

Connected embedded WAV settings to slide entry, show start and ribbon Preview.
Looping audio survives visual transition completion and silent slide entries;
new transition/action audio and Stop Previous Sound stop it. Show exit, kiosk
restart, preview retrigger and editing navigation release the audio source.
Ended, error and rejected play promises release resources, with stale completion
callbacks unable to stop a newer sound.

Validation: dev typecheck/build and focused lint pass. Transition playback browser
suite covers sound lifecycle with an instrumented Audio implementation, including
loop, silent continuation, replacement, stop, rejected playback and stale events.
Show-properties integration covers ribbon preview, show start, silent next slide
and Escape cleanup. Both suites pass; diff check passes. These verify lifecycle
and wiring, not audible fidelity. Native reference access was retried and still
returns Transport closed. Built-in sound choices and exact Mac visual/audio
behavior remain unverified; full parity remains incomplete.

### Transition sound history conflict coverage

New transition-sound and transition-apply-all journal records now include a
separate transitionSoundBefore fingerprint. This detects externally replaced WAV
bytes even when the slide XML and relationships stay identical. Apply All checks
all affected slide sounds. The optional record field preserves replay support for
existing version-1 journals; their historical conflict coverage is unchanged.
Focused editor test verifies normal replay, old-record replay, byte-only external
changes and preservation on conflict for both commands. Dev typecheck/build,
focused lint and diff check pass.

Duration investigation found that the local ECMA schema defaults spd to fast,
contrary to the builder's old documentation; corrected that comment. Exact legacy
speed-to-duration conversion is not established, so no guessed mapping was added.

### Transition compatibility branch reading

Transition/effect and sound readers now select a usable Choice in document order,
resolving Requires prefixes from inherited/local namespace declarations, or use
Fallback. Unknown requirements and unsupported modern visual children no longer
mask a stored legacy effect as None. p14 duration on standard effects remains
supported. A wrapper without a usable Choice or Fallback returns no transition.
Mutators continue retaining/updating all alternatives; Apply All preserves the
original modern markup and inherited declarations.

Reference for selection order and namespace requirements:
https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf
(section 8.3.4.2.1). This is transition-specific partial capability handling,
not a general MCE processor or implementation of modern PowerPoint effects.

Validation: 10 focused core tests and 4 transition editor tests pass; root/dev
checks and builds, focused lint and diff check pass. Cases include multiple
requirements, inherited aliases, first supported Choice, absent fallback,
consistent sound branch, preservation/copy/save/reload and changing effects.
Full modern effect/UI parity remains outstanding.

### Reverse wheel transition

Added p14:wheelReverse reading/writing, a provisional Reverse Wheel gallery entry
with spoke options, and counterclockwise playback. Newly authored reverse effects
use mc:AlternateContent with a p14 Choice and standard wheel Fallback. Effect
changes preserve timing, sound and extensions. Apply All and save/reload retain
the modern effect. The fallback is our compatibility choice, not native-verified.

Schema references:
https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/76223734-7aa7-4053-b666-acd5f73d1d9a
https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/22ebe6b5-2ade-43d9-977a-98fa194725c2

Validation: 11 focused core tests, 4 editor tests and 3 browser tests pass,
including opposite half-circle coverage, ribbon spoke selection, reload and
preview. Root/dev type checks and builds, focused lint and diff check pass.
Native comparison remains unavailable (CUA Transport closed); gallery label,
order, icon, timing/easing and visual fidelity are not claimed identical to Mac
PowerPoint. Full UI and operation parity remains open.

### Slideshow pointer controls

Added Pointer Options to the show toolbar/context menu: Automatic, Hidden, Arrow,
Laser Pointer and Red/Green/Blue laser color choices. Automatic hides the cursor
after three seconds without movement; moving restores it. Mac shortcuts Command
L/A/I/U select laser/arrow/hidden/automatic, and Control H hides the pointer.
The visual laser follows the pointer within slide bounds, is suppressed for
menus and blank/end screens, survives slide navigation and is cleared on exit.
Cursor overrides also apply inside the slide shadow DOM. These settings remain
show-session UI state and do not create document edits.

References (Mac sections):
https://support.microsoft.com/en-us/powerpoint/turn-your-mouse-into-a-laser-pointer
https://support.microsoft.com/en-us/accessibility/powerpoint/use-keyboard-shortcuts-to-deliver-powerpoint-presentations

Validation: expanded show-properties browser integration passes (15 seconds),
covering shortcuts, actual SVG cursor style, color selection, idle timeout,
movement restoration, navigation and exit cleanup. Dev typecheck/build, focused
lint and diff check pass. A 1600x1000 screenshot was visually checked for the
laser dot and toolbar. Exact installed Mac menu appearance, pointer geometry,
click behavior, pen/ink drawing and recording still need implementation or native
comparison; this is not a full pointer or overall PowerPoint parity claim.

### InkML serialization foundation

Added an internal constant-width pen trace builder using the Office InkML subset.
It emits explicit context/brush references, integer absolute samples, physical
X/Y resolution (360000 samples per cm), RGB brush colors and physical pen width.
Stroke bounds include the pen radius, including single-point marks. Invalid
coordinates, widths, colors and empty traces are rejected before serialization.

Reference:
https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/096dacae-0d2c-4861-bc4d-c8e4c6405ad3

Validation: three focused tests, root typecheck/build and focused lint pass.
This is an internal prerequisite, not a user-visible ink feature. Content-part
packaging with picture fallback, rendering, pen gestures, erase, Keep/Discard
on exit and native round-trip comparison remain to be implemented. The builder
has not yet been validated in the installed PowerPoint application.

### Ink content-part packaging

Added internal `appendSlideInk`, which stores InkML and a caller-rendered PNG,
creates slide relationships, and appends a p14 contentPart with a picture
fallback under mc:AlternateContent. Both branches share their identifier,
name and transform. Shape identifier allocation now scans nonvisual properties
in the XML tree, preventing collisions with ink and other unexposed content.
Invalid strokes or fallback headers fail before package mutation.

Validation: five ink builder/package tests pass, including save/load, relationship
resolution, matching branch geometry, repeated insertion and failed-input
atomicity. Ten existing shape-copy/slide-introspection tests, root typecheck,
build and focused lint pass. Native CUA retry returned `Transport closed`.

The append function is deliberately not exported in the public facade yet:
the typed reader/preview still omit ink content. Reading/rendering, editor
integration, pen gestures and Keep/Discard remain required. The PNG caller is
responsible for rendering the same traces cropped to the computed ink bounds;
the packaging test uses a tiny fixture and does not prove visual/native parity.

### Ink shape reading and fallback preview

The typed reader now exposes native ink content (including an AlternateContent
wrapper) as one `ink` shape. The wrapper remains the identity for copy, deletion
and stacking. Geometry reads use the p14 transform; position, size, rotation and
flip writes synchronize the fallback picture transform. PNG fallback bytes are
available through image reads and render in the preview. Copying remaps the same
nonvisual ID consistently across the native and fallback branches. Summary shape
counts include ink. Ink aspect-ratio locking is explicitly unsupported for now.

Validation: 23 focused ink, shape-copy, slide-introspection, presentation-summary
and picture-render tests pass. Ink tests cover ZIP round-trip, transformations,
copy/delete/stacking and rendered pixel colors. Root typecheck/build, dev typecheck,
focused lint and diff check pass. The pixel test uses a solid PNG fixture to prove
placement and stacking; it does not validate trace-to-PNG fidelity or native
PowerPoint behavior.

Remaining ink work includes native trace rendering when no fallback exists,
generating the matching PNG from captured strokes, public authoring/editor
commands, pen gestures, erase and Keep/Discard on exit. The internal writer is
still not publicly exported. Native comparison remains unavailable after the
previous CUA transport failure.

### Ink persistence and edit history (2026-09-24)

- Exposed `InkStroke`, `getInkBounds`, and `addSlideInk`; generated transparent SVG/PNG compatibility artwork from the same validated stroke coordinates as native InkML.
- Raster fallback is capped at 2048 pixels and preserves nonzero dimensions for extremely thin strokes.
- Added the editor `ink` command: validates all slide targets and strokes and prepares all PNGs before insertion; retains individual strokes as selectable ink objects across slides in one history entry.
- History fingerprints include every annotated slide. Tests cover export/reload, replay, undo, changed-target conflicts, and invalid late annotations without partial insertion.
- Validation: 9 ink/renderer tests, all 97 editor tests, root/preview/dev TypeScript checks, preview/dev builds, focused lint and diff checks passed. Dev build emits the existing generated sourcemap warning.
- Still pending: connect slideshow pen input and Keep/Discard UI to this command; native Mac comparison remains unverified while CUA transport is unavailable. This is persistence infrastructure, not completion of all PowerPoint operations.

### Slideshow pen UI connected (2026-09-24)

- Added Pen to Pointer Options and Cmd+P switching, pen colors, round constant-width pointer strokes in slide coordinates, per-slide temporary ink, and Shift+E clearing of current-slide temporary annotations.
- Ink is retained across navigation and hidden on blank/end screens. Pen input suppresses slide advance and hover/click actions.
- End Show and fullscreen exit lead to Keep/Discard when annotations exist. Keep commits the multi-slide ink history command; Discard leaves the presentation unchanged. Failed saves keep the dialog and ink available.
- Browser verification passed: multi-slide draw/navigation, keep, one-step undo/redo, erase, discard, save failure; existing slideshow properties/pointer/timing regression passed. Typecheck, dev build and focused lint passed. Screenshot `/tmp/pptx-show-ink.png` inspected: red line correctly placed on the slide.
- Native Mac UI appearance (palette, pointer glyph, exit dialog dimensions/text) and exact default stroke width still require comparison; no claim of complete PowerPoint parity. Additional cases still to verify include browser-controlled fullscreen exit and existing native ink erasure semantics.

### Slideshow ink lifecycle verification (2026-09-24)

- Verified actual browser fullscreen exit prompts Keep/Discard and leaves presentation mode.
- Fixed ongoing ink capture when switching tools, navigating, blanking, or leaving presentation mode; pointer release includes the final sample.
- Browser test exposed a second defect: switching from Pen to Arrow during a drag caused mouse release to advance the slide. Suppress the click belonging to that ink gesture while preserving subsequent pointer gestures.
- Fullscreen exit now reaches ink handling even if the browser rejects the fullscreen request.
- Extended browser ink test passed (fullscreen exit, tool switch, no unintended slide advance), alongside dev TypeScript/build and focused lint/diff checks.
- Native comparison retried: CUA still returns `Transport closed`; native palette/pointer/dialog fidelity remains unverified.

### Saved ink editing and compatibility metadata (2026-09-24)

- Confirmed retained ink is selectable, draggable and resizable through the normal canvas handles. Browser reload preserves the resulting geometry through journal replay.
- Fixed ink metadata mutations so native content and compatibility picture agree on selection-pane name, alt description/title and hidden state. Clearing metadata also clears the fallback attribute, while preserving its identity and unrelated properties.
- Validation: 5 native ink tests, 9 copy/ink-renderer/builder tests, extended slideshow browser test, root typecheck/build and focused lint passed. Browser test waits for Undo/Redo completion before initiating a drag; state publication alone precedes editor readiness.
- Full PowerPoint operation/UI parity remains incomplete; native comparison still unavailable.

### Ink aspect locking and action compatibility (2026-09-24)

- Added native ink aspect lock read/write through `a14:cpLocks/@noChangeAspect`, synchronized to the compatibility picture lock. Existing Format Shape controls now support retained ink without an unsupported-operation error.
- Synchronize click/hover actions, including relationship references, from native ink properties into the compatibility picture; clearing an action removes it from both branches.
- Verified saved/reloaded locks and actions, clearing, and browser corner resizing with the lock enabled. The browser test waits for the resize handle to settle after closing the sidebar before dragging.
- Validation: 26 ink/action regression tests and the extended slideshow browser test passed; root typecheck/build and focused lint/diff checks passed.
- Format reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.drawing.contentpartlocks?view=openxml-3.0.1
- Native Mac UI comparison remains unavailable; complete UI and operation parity is still pending.

### Slideshow keyboard action selection (2026-09-24)

- Added Tab / Shift+Tab cycling within visible slide hyperlinks and click-sound actions, including text-range hyperlinks. Return activates the focused action without also advancing the slide.
- Skip display/visibility-hidden content and suspend action traversal on black/white/end screens, menus and text controls. Numeric slide navigation retains precedence when a number has already been entered; kiosk links remain keyboard-operable.
- A slide without actions consumes Tab without navigating or moving focus into unrelated editor/browser controls.
- Native comparison retried: CUA still returns `Transport closed`. Mac Shift+E is documented; saved-ink erasure semantics and additional eraser UI remain unverified, so no unsupported Windows-only eraser controls were added.
- Reference: https://support.microsoft.com/en-us/accessibility/powerpoint/use-keyboard-shortcuts-to-deliver-powerpoint-presentations (Mac slide-show control table).
- Validation: browser hyperlink test covers forward/backward cycling, wraparound, Return navigation, a slide without links, and repeated keyboard activation of a sound-only action without slide advance. Hyperlink and slideshow ink browser regressions, dev typecheck/build, focused lint and diff checks passed. Existing dev sourcemap warnings remain.
- Full native UI and operation parity remains incomplete.

### Mac slideshow start, exit and hidden-slide shortcuts (2026-09-24)

- Command+Return now starts from the current configured slide, while Command+Shift+Return starts at the beginning. Modifier checks prevent unrelated Control/Option combinations from starting playback; handled events are not processed again.
- Added Command+Period and unmodified Hyphen exits through the existing fullscreen and Keep/Discard lifecycle. Escape remains supported, including kiosk shows.
- H reveals the immediately following hidden slide when one exists. Normal advancement still skips hidden slides. Returning from a revealed slide restores the current show-sequence position rather than getting stuck at its beginning.
- Reference: https://support.microsoft.com/en-us/accessibility/powerpoint/use-keyboard-shortcuts-to-deliver-powerpoint-presentations (Mac start/end and hidden-slide shortcuts).
- New browser coverage exercises current/beginning start, both exit shortcuts, ordinary hidden-slide skipping, H reveal, forward/backward resumption, and pending-ink discard after Command+Period.
- Native UI/operation parity remains incomplete; installed Mac comparison remains unavailable following the CUA transport failure.
- Validation: new shortcut and existing slideshow-range/loop/timer browser tests pass after the backward-navigation correction; hyperlink browser regression also passed during this change. Dev typecheck/build, focused lint and diff checks passed; existing sourcemap warnings remain.

### Hidden-slide navigation resumption (2026-09-24)

- Number jumps to hidden slides now anchor ordinary playback to the displayed slide instead of the previous playback position. Previous-button availability uses the same anchor.
- Starting from a selected hidden slide retains that slide; advancing then follows the configured visible sequence. An empty visible sequence can end normally.
- Custom shows retain their explicit ordering, including nested return and kiosk restart state.
- Validation: shortcut browser coverage includes backward number jumps to hidden slides, forward/backward resumption and current-hidden-slide start. Shortcut and custom-show tests pass sequentially; range/loop/timer test passed separately in the initial batch. Fullscreen completion is now awaited in the shortcut test. Initial concurrent execution failed in later ink/link actions; sequential rerun passed. Dev typecheck/build and focused lint passed.
- Installed Mac comparison remains unavailable due to CUA transport failure. Current-hidden-slide start is not yet verified against the installed app. Full UI/operation parity remains incomplete.

### Escape releases slideshow drawing tools (2026-09-24)

- Unmodified Escape while the pen or laser is active now releases the tool and retains the running show. A subsequent Escape exits through the existing annotation Keep/Discard flow. Command+Period still exits directly. Open menus retain their own Escape handling.
- Releasing a pen during pointer capture finalizes the current stroke, preserves accumulated annotations, and suppresses the release click so the slide does not advance.
- Reference: https://support.microsoft.com/en-us/powerpoint/training/start-the-presentation-and-see-your-notes-in-presenter-view (macOS, Use the controls in Presenter view). The installed app is still unavailable: CUA returned `Transport closed` in this turn. Actual macOS/browser fullscreen Escape interception and native pointer appearance remain unverified.
- Validation: all three sequential slideshow ink, show-properties and Mac-shortcuts browser tests passed. They cover pen/laser release without ending playback, continued transition sound, retained ink, release during drawing, subsequent exit and save/discard, and direct Command+Period exit. Dev typecheck/build, focused lint and diff checks passed; existing sourcemap warnings remain. Full UI and operation parity is not complete.

### Presenter view workspace (2026-09-24)

- Added the Slide Show ribbon Presenter View command and Option+Return entry. The current canvas remains the live slideshow canvas; a right pane shows the next slide and editable speaker notes, with text-size controls and a draggable/keyboard-resizable divider.
- Added elapsed timer pause/resume/reset, local clock, centered previous/next and position controls, and a bottom slide navigator including hidden slides. Navigation uses the existing show sequence; next-slide lookup handles custom-show order, nested return and looping.
- Reuses the existing notes textarea and persistence/history command. Use Slide Show restores the ordinary show; End Show and existing exit shortcuts restore the notes control to its editing pane. Black/white/end overlays are constrained to the current-slide area. Presenter buttons retain native keyboard activation without double slide advancement.
- Compared the rendered layout with Microsoft's published Mac screenshot: https://support.microsoft.com/en-us/powerpoint/media/ppt365mac-presenterview-fullscreenview.png . Reference instructions: https://support.microsoft.com/en-us/powerpoint/training/start-the-presentation-and-see-your-notes-in-presenter-view (macOS). Clock, next-slide/notes pane, timer, bottom navigator and Use Slide Show placement/labels were adjusted from this reference. This is not an installed-version pixel comparison.
- New browser test covers ribbon and Option+Return entry, hidden-slide skipping/selection/resumption, notes persistence through reload, note size, timer pause/resume/reset, keyboard pane sizing, Enter activation, blank-screen recovery, Use Slide Show and full exit. Rendered screenshot /tmp/pptx-presenter-view.png inspected. Presenter, Mac shortcuts and hyperlinks passed; presenter and ink regressions passed after waiting for the existing post-save rebuild in the ink revision assertion. Dev typecheck/build, focused lint and diff checks passed; existing sourcemap warnings remain.
- Remaining: separate audience output and Swap Displays, automatic multi-monitor entry, presenter preferences, native toolbar icons/Tips, full navigation keyboard behavior, native pane sizes and installed Mac visual/interaction validation. The current view is a single-display implementation, not complete Presenter View parity. Full application UI and operation parity remains incomplete.

### Presenter audience output foundation (2026-09-24)

- Presenter View opens a separate audience popup when `screen.isExtended` reports multiple displays. Where available, Window Management permission allows placement on a display other than the current one. Single-display behavior is unchanged.
- The audience document receives current slide SVG, temporary ink, laser position, and black/white/end screens. Notes, next-slide preview, thumbnails, and editor controls are excluded. Navigation keys are routed through existing show handlers, and output closes on End Show, Use Slide Show, or page exit. Closing the audience popup leaves the presenter running.
- Browser coverage uses a simulated extended-screen flag and a real second page: slide navigation in both directions, private-note exclusion, black/white screens, ink coordinates/erase, and exit cleanup. It does not establish physical monitor placement parity.
- Remaining: native Swap Displays and display preferences; popup-blocked recovery UI; screen-change handling; transition and animation playback in the audience document; automatic audience fullscreen. Opening the popup consumes transient browser activation, so presenter fullscreen may be denied; the presenter workspace remains usable in its window. Native installed-Mac comparison is still unavailable (CUA transport closed).

### Audience transition synchronization (2026-09-24)

- Audience output now mirrors the current transition's isolated SVG layers and animation clock. It copies the existing KeyframeEffect directly, preserving coordinate precision and random effect ordering; it does not choose a second random transition. Source-sized layers are scaled to the audience viewport, including pixel-based clip paths.
- Transition resize revisions update copied effects after the source ResizeObserver recalculates clip paths. Replacing, cancelling, blanking, or ending the show removes the audience transition; the animation loop stops when no transition is active.
- Browser coverage verifies fade, push, strips, checker, and dissolve at two seek positions, exact computed clip/transform parity, isolated SVG paint IDs, a different output viewport, resize updates, cancellation, and window cleanup. Existing presenter and transition playback/sound tests pass. Visual fixture: `/tmp/pptx-audience-transition.png`.
- Remaining overall scope includes Swap Displays, physical monitor and fullscreen behavior, slide-object animation playback, popup recovery, and installed Mac UI comparison. This closes the static-only transition limitation of the preceding audience-output entry; it does not establish all-operation PowerPoint parity.

### Presenter Swap Displays command (2026-09-24)

- Added the Mac-labelled `Swap Displays` command when a live audience window and permission-granted multiple-screen details are available. It requests presenter fullscreen on the audience screen, verifies the browser's current screen changed, then places audience output on the former presenter screen. A denied or silently ignored screen request reports the failure without moving audience output.
- Display-disconnection events refresh availability. Ending the show removes the screen listener and display state. A blocked audience popup now has a presenter-side status message explaining how to retry.
- Browser tests simulate screen details and fullscreen placement: swap twice without changing the current slide/notes, denied placement leaves output untouched, and a single remaining display hides the command. Existing audience-transition and presenter lifecycle tests pass. Actual physical monitor placement is NOT established by these simulated tests.
- Native CUA comparison retried this turn and still returns `Transport closed`. Remaining display work includes automatic audience fullscreen, manual window moves, reconnection and more-than-two-display selection, and matching installed Mac appearance/behavior. Full PowerPoint operation parity remains incomplete.
- Follow-up verification also covers silently ignored fullscreen screen requests. Audience slide layers now form a stacking context so temporary ink and laser overlays cannot appear above the black/white/end screen; browser coverage checks black-screen hit testing after an ink stroke.

### Audience slide action clicks and keyboard focus (2026-09-24)

- Audience slide artwork now preserves clickable hyperlinks and sound-only actions. Clicks are mapped to the corresponding live slide action so existing internal navigation, external target behavior, and sound handling run once, without an extra slide advance. Plain clicks keep using the slideshow handler; pen-mode clicks do not accidentally activate links or advance.
- Tab/Shift-Tab continue through the shared Mac show-hotspot handler, with focus reflected on the audience link. Return uses the same focused action. Links and sound actions have a pointer cursor in the audience window.
- Presenter browser coverage adds an in-document link click, Tab/Return activation, sound-only click without slide advancement, and an external link served entirely through a browser route fixture. Authoring/serialization is covered separately by the existing hyperlink suite; these audience checks inject rendered SVG action fixtures to exercise delegation.
- Remaining audience interactions include hover actions, direct pen drawing and laser movement from the audience window, and context-menu placement. Installed Mac UI/behavior comparison and the full-operation parity goal remain incomplete.

### Audience hover actions (2026-09-24)

- Audience artwork forwards mouse-over hyperlinks and action sounds to the corresponding source artwork, sharing the existing action execution and repeated-position suppression.
- Related targets are mapped across documents so movement within one action does not retrigger it. Pointer movement releases the position suppression. Blank screens, pen mode and touch input suppress audience hover actions.
- Presenter browser coverage verifies audience hover navigation and touch/pen suppression. Dev typecheck/build, targeted lint, diff whitespace check and the hyperlink, presenter-view and audience-transition browser tests pass (3 tests).
- This is browser verification, not installed Mac PowerPoint verification. Native UI transport remains unavailable. Audience pen/laser input, context menus, physical display verification and broader operation/UI parity remain unfinished.

### Audience laser input and pointer appearance (2026-09-24)

- Audience mouse movement now maps its letterboxed slide coordinates into the shared slideshow pointer state. The laser therefore follows the audience mouse with the same relative position on both render surfaces, despite differing dimensions.
- Audience cursor appearance follows automatic, hidden, arrow, laser and pen modes. The laser overlay does not intercept pointer input. Leaving or blurring the audience window clears the pointer; blank screens continue to suppress it.
- Presenter browser coverage verifies a 25%/75% laser position, hidden laser cursor, black-screen suppression, pointer exit and restoration of the arrow cursor. Dev typecheck/build, targeted lint, whitespace check and both presenter-view/audience-transition browser tests pass.
- Native comparison was retried but CUA still returns `Transport closed`. Direct audience pen input and context menus, physical display behavior and complete Mac UI/operation parity remain unfinished.

### Audience pen input and annotation persistence (2026-09-24)

- Extracted slideshow pen surface binding so the presenter slide and audience slide use the same stroke collection, color, coalesced samples, click suppression and Keep/Discard lifecycle.
- Input coordinates are normalized against the surface receiving the event. Active strokes track the owning surface and release that surface's pointer capture on finish, mode/slide changes and audience close; another surface cannot accidentally extend a stroke with a matching pointer ID.
- Presenter browser coverage draws directly in the audience window, checks normalized coordinates and matching paths on both surfaces, verifies no accidental navigation, returns to the stroke after navigating away, and keeps the annotation across reload.
- Dev typecheck/build, targeted lint and whitespace checks pass. Presenter-view, audience-transition and existing full ink lifecycle browser tests pass (3 tests), including existing multi-slide save, undo and discard coverage.
- Audience context menus, physical display validation and complete installed Mac UI/operation parity remain unfinished. Native app transport is still unavailable as recorded in the previous entry.

### Audience context menus (2026-09-24)

- The audience window now opens the shared slideshow context menu via right click or Shift+F10. It uses the same actions, enabled states, pointer/color choices, screen controls and End Show command as the presenter surface.
- Menu construction, submenu placement and keyboard focus use the owning document and its viewport. The shared menu returns to the main document on dismissal/close, and audience menu clicks are excluded from slide advancement. Resize dismisses audience menus.
- Browser coverage verifies pointer submenu selection, keyboard Screen submenu activation, clearing a black screen without advancing, Escape/focus restoration, and End Show followed by keeping audience ink. The menu screenshot `/tmp/pptx-audience-context-menu.png` was visually inspected.
- Dev typecheck/build, lint and whitespace checks pass. Hyperlink, ink, presenter and transition tests passed; after lifecycle refinements the presenter and transition tests passed again. A fixture hover race was resolved by moving the physical test pointer outside the slide before synthetic hover checks.
- This shares the current implementation's menu; it does not establish exact installed PowerPoint visual parity. Native app comparison, physical monitor validation and broader missing functionality/UI remain unfinished.

### By Title slideshow navigation labels (2026-09-24)

- By Title now reads title/ctrTitle placeholder text from each slide, collapses whitespace for menu display and retains a numbered Slide fallback for missing/empty titles. Shared presenter/audience actions still navigate by stable slide position, including hidden slides.
- Presenter browser fixture now includes a real centered-title placeholder populated through setSlideTitle. It verifies exact title-menu labels for untitled and titled slides and navigation through the titled entry, followed by normal hidden-slide skipping.
- Dev typecheck/build, lint and whitespace checks pass; presenter-view and hyperlink browser tests pass (2 tests).
- Reference: https://support.microsoft.com/en-us/powerpoint/go-to-a-slide-when-delivering-your-presentation describes Mac By Title navigation by title or number. Exact menu typography, punctuation and remaining installed-Mac behavior remain unverified while native app transport is unavailable; full parity is not complete.

### Fade-out visibility timing (2026-09-24)

- Exported fade-out effects now delay the hidden visibility state until the authored fade duration ends. Previously the exit hid the shape immediately while its opacity animation was still running. Instant disappear and entrance visibility remain immediate.
- Regression coverage saves and reads back five effect/duration combinations, including zero-duration fade-out, and checks visibility timing plus opacity duration. Animation write/read/discovery suites pass (3 files, 16 tests); root typecheck/build, targeted lint and whitespace checks pass.
- References: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.condition?view=openxml-3.0.1 and https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.startconditionlist?view=openxml-3.0.1 document activation delay and start conditions.
- This fixes the saved timing definition; application shape-animation playback/editor controls are still missing. Native rendering of the exported effect and complete Mac UI/operation parity remain unverified and incomplete.

### Main animation sequence metadata (2026-09-24)

- Added getSlideAnimationSequence and exposed it in the editor model. It reads direct main-sequence groups and ordered preset effects, target IDs, click/with/after trigger tokens, effect-node delays and supported opacity durations. Multiple effects on one shape retain separate sequence entries.
- Unknown presets and empty groups remain represented instead of silently collapsing the sequence. Interactive sequences are excluded; non-numeric/missing timing values remain null. This metadata reader does not evaluate general nested timing conditions, paragraph builds, repeats or interactive triggers.
- Save/reload coverage verifies repeated effects, durations, unique timing IDs and clearing. An independent timing-tree fixture verifies unknown presets, mixed triggers, target order, empty groups and interactive-sequence exclusion. Related tests pass (3 files, 18 tests); root/dev typechecks and builds, targeted lint and whitespace checks pass.
- Playback integration, animation editing controls, broader effects and installed Mac behavior comparison remain unfinished. Full operation/UI parity is not established.

### Slideshow shape effect playback (2026-09-24)

- Connected main-sequence metadata to slideshow navigation. Supported appear/disappear/fade-in/fade-out effects now consume click groups before advancing slides. A next action during an active effect completes that group; previous restores the preceding group state. Effect-node delays and with/after ordering within a group are evaluated for this subset.
- Entrance targets begin hidden; fade frames update SVG presentation styles, and teardown restores original style attributes. Slide changes/restarts reset state. Transition snapshots now include the current rendered animation state instead of always using raw artwork. Automatic advance waits while a group runs; presentation navigation buttons reflect remaining effect history.
- Added deterministic browser coverage for fade intermediate values, completion, rewind, grouped delays, restart and style restoration, plus an actual generated-deck/server test for click sequencing before slide navigation and exiting to normal editing. These and existing presenter/audience transition browser tests pass (4 tests). Dev typecheck/build, targeted lint and whitespace checks pass.
- This is an initial supported-effect playback implementation, not a complete PowerPoint timing engine: unknown effects, paragraph builds, repeats, interactive triggers, full native timing semantics and animation editing UI remain outstanding. Audience animation-specific rendering/focus behavior and exact installed Mac interactions still need dedicated verification. Full UI/operation parity remains incomplete.

### Audience animation frames preserve focus (2026-09-24)

- Audience rendering no longer replaces the entire SVG for shape-wrapper style changes. It compares artwork structure without wrapper styles, then copies those styles onto existing audience nodes when the structure is unchanged. Actual artwork/structure changes still replace the SVG.
- This keeps links and keyboard focus alive during fade frames and during restoration of original styles. Closing/reopening output clears the structure cache.
- Extended the popup transition test with real animation playback at 25/50/75/100 percent, exact audience opacity checks, persistent link identity/focus, original style removal and a subsequent non-style artwork change. The animation, presenter and audience transition tests pass (4 tests); dev typecheck/build, targeted lint and whitespace checks pass.
- Installed Mac comparison was retried; native CUA still reports Transport closed. This does not establish exact Mac animation timing/visual parity. Broader animation effects/editor UI and full operation parity remain unfinished.

### Initial animation authoring ribbon (2026-09-24)

- Added an Animations tab between Transitions and Slide Show with Entrance/Exit effect groups for the currently supported Appear, Fade, Disappear and exit Fade. Buttons require selected objects and use the existing busy/edit lifecycle.
- Added animation-add editor commands using setShapeAnimation for selected targets; commands participate in existing persistent history, undo and redo. Existing effects are appended, matching object-selection addition rather than effect-marker replacement (effect-marker selection is not yet implemented).
- The browser integration test now authors both effects through the ribbon instead of injecting them in the source deck, verifies disabled controls without selection, undo/redo and reload persistence, then tests slideshow sequencing and exit restoration. Two animation browser tests pass; dev typecheck/build, targeted lint and whitespace checks pass. Viewed /tmp/pptx-animation-ribbon.png for layout inspection.
- Reference: the macOS section of https://support.microsoft.com/en-us/powerpoint/training/animate-text-or-objects documents separate effect categories and applying an additional effect by selecting the object again. This sparse initial ribbon is not exact installed-Mac parity: effect galleries, markers, pane, selected-effect replacement/removal, timing, reorder, preview and other effects remain outstanding.

### In-place opacity animation duration editing (2026-09-24)

- Added setSlideAnimationDuration and an animation-duration editor command, targeting a timing ID so repeated effects on one object remain distinct. Supported opacity fades update in place while preserving targets, IDs, ordering and other effects. Fade-out visibility switches at the previous endpoint follow the new duration.
- Invalid durations, missing effects and unsupported behaviors are rejected before mutation. Native animEffect/filter fades are not covered by this opacity-specific implementation.
- Related API tests pass (20 tests). An editor regression additionally verifies save/reload, undo/redo and conflict rejection for the second effect on one object. Root/dev typechecks and builds, targeted lint/format and whitespace checks pass.
- This command is not yet connected to effect selection or timing controls in the UI. Native comparison remains unavailable after Transport closed errors; full Mac UI/operation parity remains incomplete.

### Animation effect selection and duration controls (2026-09-24)

- Added an Animation Pane toggle and ordered effect rows, including repeated effects on the same target. Selecting an effect enables the ribbon duration input for supported fades. Selection is scoped to slide key/timing ID and cleared when the effect disappears; row keyboard navigation preserves focus across updates.
- Duration edits use the prior in-place command with text-edit completion, busy guards and seconds-to-milliseconds conversion. Instant effects have disabled duration controls. The pane reserves stage width rather than covering slide content and is hidden during presentation.
- Browser integration now selects separate entrance/exit effects, edits only the exit duration, verifies undo/redo and reload persistence, checks non-overlapping pane/stage bounds and exercises slideshow navigation. Both browser tests pass. Dev typecheck/build, targeted lint and whitespace checks pass; inspected /tmp/pptx-animation-pane.png.
- Native CUA retry still reports Transport closed. This initial pane is not verified exact Mac UI: effect markers, full pane controls, start/delay, reorder/remove, preview and broader timing behavior remain unfinished. Full operation/UI parity remains incomplete.

### Selected animation effect removal (2026-09-24)

- Added removeSlideAnimation and an animation-remove editor command. Removal works on a cloned timing tree, removes only the selected main-sequence effect and empty enclosing wrappers, cleans unreferenced build groups and preserves unrelated effects. Incoming timing references and nested effect structures are rejected before mutation. Empty main sequences accept subsequent effect additions.
- Connected pane Remove and Delete/Backspace to the selected effect, moving selection/focus to a remaining neighbor. Excluded the pane from canvas keyboard shortcuts; browser testing caught and resolved a Delete-key conflict with object deletion.
- API coverage verifies save/reload, middle-effect removal, build cleanup, deleting all/re-adding, preservation of same-group siblings/interactive sequences and atomic rejection of incoming timing references. Related API tests pass (20). Browser coverage verifies effect-only deletion with the shape retained and undo/redo restoring duration (2 tests pass). Root/dev typechecks and builds, targeted lint/format and whitespace checks pass.
- This does not establish complete native deletion semantics or exact Mac pane UI. Reordering, markers, timing triggers, full effect support and the broader operation/UI parity scope remain outstanding.

### Filter-based fade timing support (2026-09-24)

- Extended fade timing recognition and in-place duration edits to p:animEffect with filter="fade" and matching in/out transition, in addition to p:anim/style.opacity. XML stays in its original behavior format. Both metadata reads and writes share the behavior recognition helper; multiple recognized fade behaviors no longer expose an arbitrary duration.
- Custom progress, different filters and direction/preset mismatches are excluded from duration editing. This is not general filter/progress playback support.
- Independent entrance/exit XML fixtures verify duration, target/trigger/delay preservation, save/reload and atomic rejection of unsupported variants. Related API tests pass (23). A browser test loads filter-based timing through the real package reader and verifies 50 percent opacity halfway through the effect; all animation browser tests pass (3). Root/dev typechecks and builds, targeted lint/format and whitespace checks pass.
- Reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.animateeffect?view=openxml-3.0.1 documents animEffect and its cBhvr/cTn duration. Native Mac comparison remains unverified; the previously named /tmp/pptx-macro-audit directory was not present when checked this turn. Broader timing semantics, UI details and full parity remain incomplete.

### Selected animation delay editing (2026-09-24)

- Added setSlideAnimationDelay and the animation-delay editor command, with a ribbon Delay input for the selected effect. Changes preserve effect identity, trigger and behaviors, and participate in undo/redo. Simple numeric start conditions are updated or created; complex event/reference/multiple conditions are rejected before mutation.
- API validation follows the existing unsignedIntMs convention of rounding fractional milliseconds. Coverage checks invalid values, rounding, save/reload, absent start conditions and atomic rejection of complex conditions. Related API tests pass (25).
- All three animation browser tests pass, including delayed fade opacity before/at/after its start and selected-effect delay editing with undo/redo. The integration test waits for timing edits to finish before issuing Delete. Root/dev typechecks, dev build, targeted formatting/lint and whitespace checks pass.
- Native CUA retry returned Transport closed. User authorization for temporary reference-document changes remains in effect, but no reference document was modified this turn. Exact Mac UI verification, complex timing semantics and full operation/UI parity remain incomplete.

### Independent animation reordering (2026-09-24)

- Added moveSlideAnimation and an animation-move editor command. Independent click effects move earlier/later as complete timing subtrees, preserving identifiers, targets, timing and build data. Boundary moves are no-ops; grouped effects, non-linear wrapper structures and timing references across moved subtrees are rejected before mutation.
- Added Move Earlier/Move Later arrows above the Animation Pane list. Selected timing ID and keyboard focus follow the moved effect; boundary buttons are disabled. Undo/redo uses existing editor history. Multiple-effect group reordering still requires timing regrouping and is not implemented.
- Related API tests pass (28), covering saved order, exact XML restoration after inverse movement, boundaries and dependency rejection. All three animation browser tests pass, including pane selection, boundary state and undo/redo. Root/dev typechecks and builds, targeted formatting/lint and whitespace checks pass. Viewed /tmp/pptx-animation-pane.png: arrows and effect rows remain inside the pane with no canvas overlap.
- Reference: https://support.microsoft.com/en-us/powerpoint/change-remove-or-turn-off-animation-effects-in-powerpoint describes Mac pane arrows for reordering. Installed Mac UI remains unverified after the preceding connection failure; this change does not establish complete UI or operation parity.

### On-slide animation number selection (2026-09-24)

- Added numbered buttons beside animated shapes, with separate stacked buttons for repeated effects and click-group numbers matching the pane. Rotated object bounds position the stack; marker size stays in screen pixels. Selecting a number selects the timing ID and target shape(s), updates timing controls and highlights its pane row. Pane selection and reordering update marker highlights.
- Markers appear while the Animations tab or Animation Pane is active, are suppressed during presentation/drawing/cropping and are rebuilt for slide/revision/size changes. Marker pointer events bypass shape dragging; Delete/Backspace removes the effect while retaining the shape, and focus moves to a surviving marker or canvas.
- Three animation browser tests pass, including marker visibility on tab/slide changes, selection synchronization, reordering and marker Delete with undo/redo. Dev typecheck/build, targeted lint/format and whitespace checks pass. Viewed /tmp/pptx-animation-pane.png to verify marker placement and selected styling.
- Reference: https://support.microsoft.com/en-us/powerpoint/training/animate-text-or-objects describes Mac numbered effects and selecting a number to revise an effect. Exact installed-Mac visuals, paragraph build numbering, arbitrary overlapping-shape marker layout and broader full parity remain unverified or incomplete.

### Animation preview in the editing view

- Added ribbon Preview and Animation Pane Play From, Play Selected, and Stop controls. Playback reuses the existing appear/disappear/fade renderer and stored duration/delay; preview does not write slide data or history.
- Preview hides editing overlays and restores original SVG styles on completion, Stop, Escape, slide navigation, slideshow entry, or an edit command. Successive click groups advance automatically during preview.
- Browser coverage exercises selected/from/all entry points, completion, Stop, Escape, navigation cancellation, editing overlay restoration, and unchanged revision/history, alongside slideshow regression coverage.
- Validation: dev TypeScript/build, targeted lint and diff checks; animation browser suite (3 tests). Screenshot review found arrow wrapping and stale selected-preview enablement; both corrected.
- Scope remains incomplete: unsupported effects, interactive triggers and native group/paragraph preview semantics are not implemented. Exact installed Mac appearance and timing remain unverified while native app transport is unavailable. Mac reference: https://support.microsoft.com/en-us/powerpoint/training/animate-text-or-objects and https://support.microsoft.com/en-us/PowerPoint/animate-or-make-words-appear-one-line-at-a-time-in-powerpoint.

### Animation Start control and click grouping

- Added Start choices On Click / With Previous / After Previous to the Animations ribbon, with selected-effect state, pending-edit handling and undo/redo integration.
- Added `setSlideAnimationStart`: retains effect/build subtrees and IDs, rebuilds simple main-sequence click groups, and uses preceding parallel-group end references for After Previous so later duration edits remain effective. Unsupported wrapper attributes/conditions and external references to replaced wrapper IDs reject atomically.
- Leading automatic effects play on slideshow entry after the slide transition, including when slide-advance timings are disabled. Their markers start at 0; shared click groups share a marker number. Rewinding does not immediately restart the automatic group.
- Play From now starts at the selected effect within a group rather than replaying earlier effects in that group.
- Validation: root/dev TypeScript and builds, targeted lint/diff checks, 30 animation API tests, 3 animation browser tests covering Start changes, undo/redo, numbering, reload and first automatic playback. Screenshot review caught the third Timing row clipping; compact field sizing and bounding-box assertions address it.
- Scope remains incomplete: arbitrary native timing wrappers, interactive triggers, full grouped reordering/removal, unsupported effect playback and exact native Mac UI fidelity need further work. CUA reconnection attempt still returned `Transport closed`; exported native playback is not yet verified in PowerPoint.
- References: https://support.microsoft.com/en-us/powerpoint/set-the-start-time-and-speed-of-an-animation-effect ; https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.condition?view=openxml-3.0.1 ; https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.triggereventvalues?view=openxml-3.0.1 .

### Removing effects from automatic animation groups

- Reused the validated simple-sequence reader and builder for removal. Deleting an effect now reconnects subsequent With Previous / After Previous groups without references to deleted wrapper timing IDs, while preserving remaining effect IDs, start modes, duration, delay and build data.
- Complex native timing structures continue through conservative subtree removal and dependency checks. Empty native click stops prevent sequence rebuilding so they are not silently discarded. Unsupported dependent deletions remain atomic.
- Added coverage for deleting each of six positions across mixed click/parallel/sequential groups, save/reload, valid timing references, build cleanup, removing all effects and authoring again. Browser coverage verifies removal of a leading click effect, automatic marker numbering and undo/redo.
- Validation: 36 animation API tests and all 3 animation browser tests pass; root/dev typechecks and builds, targeted lint/format and whitespace checks pass. The browser test waits for selected-effect UI restoration after the command completes.
- Remaining start modes are intentionally preserved; exact native Mac regrouping behavior has not been observed while CUA transport is unavailable. Grouped reordering, arbitrary native timing structures and full UI/operation parity remain incomplete.

### Reordering effects in automatic animation groups

- Move Earlier / Move Later now address adjacent effects in the flattened pane order, including With Previous and After Previous effects within or across click groups. Simple sequences are rebuilt with the reordered effects' existing start modes, IDs, behavior subtrees, delays and durations. Independent click-only sequences retain their existing subtree-swap path.
- Explicit timing references crossing a moved effect's boundary reject before mutation. Complex native wrappers still require further support; the command reports an error without rewriting them.
- Added API coverage for five adjacent swaps across a mixed six-effect sequence, inverse moves, save/reload and atomic rejection of an external dependency. Browser coverage verifies grouped movement, automatic numbering, selected-effect controls, boundary disablement and undo/redo.
- Validation: 41 animation API tests and all 3 animation browser tests pass; root/dev TypeScript and builds, targeted lint/format and diff checks pass. Mac CUA retry returned `Transport closed`; exact native UI and regrouping behavior remain unverified. Full operation/UI parity remains incomplete.

### Changing the selected animation from the ribbon

- Added `setSlideAnimationEffect` and the editor `animation-effect` command. Replacement retains the selected timing ID, start conditions, click grouping and build data, swaps the behavior subtree, remaps new behavior IDs and copies the existing object/paragraph target. Fade-to-fade replacement retains duration, including zero; instant-to-fade uses the existing 500 ms authoring default.
- Ribbon effect buttons replace the selected pane/marker effect; selecting the object directly or in the Selection pane clears the effect selection so the gallery adds another effect. The active effect is highlighted in the ribbon. Explicit references to replaced behavior IDs, nested effects and inconsistent targets reject before mutation.
- Microsoft Mac instructions distinguish pane selection for replacement from object selection for addition: https://support.microsoft.com/en-us/powerpoint/change-remove-or-turn-off-animation-effects-in-powerpoint . Installed native UI comparison remains pending while CUA is unavailable.
- Validation: 46 animation API tests and all 3 animation browser tests pass; root/dev typechecks/builds, targeted lint/format and whitespace checks pass. Browser coverage checks identity/count preservation, ribbon selection, undo/redo and object reselection followed by addition. It exposed missing selection clearing on the shape overlay; corrected in its pointer handler. The test also waits for editing to finish before clicking the canvas.
- Remaining scope includes the full effect gallery, per-effect options, multiselect, native timing variants and exact Mac appearance. This does not establish complete operation/UI parity.

### Animation Pane multiple selection and batch deletion

- Added Shift-click range selection, Command/Ctrl-click toggle, Shift-arrow range extension and Command/Ctrl+A to the Animation Pane list. The list exposes `aria-multiselectable`, selected rows and on-slide markers reflect the full selection, and plain row/marker or object selection resets it.
- Delete/Remove sends selected timing IDs as one edit command and one undo history entry. `removeSlideAnimations` validates IDs and restores the original timing tree plus package data if any removal fails. Remaining simple automatic groups reconnect through the existing removal implementation.
- Validation: 48 animation API tests and all 3 animation browser tests pass; root/dev typechecks/builds, targeted lint/format and whitespace checks pass. Coverage includes mixed-group batch deletion, duplicates, saved results, rollback after an earlier successful removal, keyboard/range/toggle selection, preserved shapes and single-step undo/redo.
- Multiple-selection timing/effect edits, selected-only preview and reordering are not implemented yet; those controls are disabled for a multiple selection instead of silently acting on only its primary row. Cross-references between selected native effect subtrees can still prevent batch deletion. Exact installed-Mac selection visuals and full operation/UI parity remain incomplete.
- Mac reference for Shift-range deletion: https://support.microsoft.com/en-us/powerpoint/change-remove-or-turn-off-animation-effects-in-powerpoint .

### Multiple-selection animation settings

- Added atomic `setSlideAnimationsSettings` and a single editor history command for changing selected effects, start conditions, duration and delay. Duplicate IDs are deduplicated; missing effects reject before mutation; a cloned timing snapshot restores both the in-memory tree and package data if a later setting fails.
- Ribbon settings now apply to every selected animation. Shared values are displayed; mixed values are blank and mixed effects have no gallery highlight. Duration remains disabled when any selected effect lacks supported fade timing. Selection IDs are captured before asynchronous text completion.
- Validation: 50 animation API tests and all 3 animation browser tests pass. Tests cover selected subsets, unchanged neighbors, persistence, partial-failure rollback, batch start/effect/duration/delay changes and undo/redo. Root/dev typechecks and builds, targeted formatting/lint and whitespace checks pass.
- Installed-Mac comparison retry still fails with `Transport closed`. Multiple selected-effect preview/reordering, the full effect gallery and complex native timings remain outstanding; full operation/UI parity is not complete.

### Previewing multiple selected animation effects

- Play Selected now accepts the full selection, filters out unselected effects and preserves slide order, click groups, With Previous/After Previous sequencing and per-effect delays. Play From starts at the earliest selected effect. Both snapshot selected IDs before text-edit completion.
- Enabled the selected-preview control for multiple selection. Existing stop, Escape, automatic completion and slide navigation restore the artwork without adding history entries.
- Validation: dev typecheck/build and targeted lint/format checks pass. All 4 animation browser tests pass; an additional After Previous assertion passed in a targeted rerun. Deterministic browser checks cover unordered IDs, an omitted predecessor, simultaneous starts, delayed sequential starts, excluded shapes and style restoration. The real editor test now previews a multiple selection and checks Escape, Stop, automatic completion and unchanged history.
- Exact installed-Mac comparison remains unavailable from the preceding Transport closed failure. Full operation/UI parity, including multiple-effect reordering and the full effect gallery, remains incomplete.

### Reordering multiple selected animation effects

- Added atomic `moveSlideAnimations`: selected contiguous blocks move one adjacent position while preserving selected order, including noncontiguous selections and selections at a boundary. Input selection order does not affect movement. Missing IDs reject before mutation; any failed individual move restores the complete timing tree and package data.
- The editor sends the selection as one move/history command, keeps row/marker selection and focus, and enables Earlier/Later only when a selected effect has an unselected neighbor in that direction.
- Validation: 56 animation API tests and all 4 animation browser tests pass. API cases include grouped automatic effects, noncontiguous/whole-list selections, saved metadata and rollback after an earlier successful move. Browser checks cover selection retention, boundary buttons and single-step undo/redo. Root/dev typechecks/builds, targeted lint/format and whitespace checks pass.
- Native-Mac behavior/visual comparison remains pending; broad UI/operation parity, full effect galleries and complex native timing support remain incomplete.

### Animation row context actions

- Added row context actions for Start On Click, Start With Previous, Start After Previous and Remove. Right-clicking within the current selection preserves it; right-clicking another row selects that row. Shared start modes are checked; mixed modes have no check. Shift+F10/ContextMenu opens the same actions with a live row focus anchor.
- Extracted the shared start-change handler so ribbon and context actions use the same selected IDs, async text completion guard, pending state and single history command.
- Validation: all 4 animation browser tests pass, including right-click selection preservation, checked start state, keyboard menu opening, multi-effect start change/undo, unselected-row targeting and Remove/undo. Dev typecheck/build, targeted lint/format and diff checks pass.
- Official start-mode semantics reference: https://support.microsoft.com/en-us/powerpoint/set-the-start-time-and-speed-of-an-animation-effect . This does not establish the exact installed Mac context-menu arrangement. Native CUA retry still returns Transport closed; menu appearance/order and full operation/UI parity remain unverified/incomplete.

### Dragging animation rows to reorder

- Added atomic `reorderSlideAnimations` for placing a selection before a target or at the end. Original slide order within the selection and effect metadata are retained. Missing destinations reject before mutation; dependency failures restore all preceding moves and package timing.
- Animation rows now support native browser drag/drop with a before/after insertion indicator. Dragging a selected row moves the selection; dragging another row selects that effect. The editor checks slide identity, preserves selection/focus, records one history entry, and avoids edits for unchanged/self-selection destinations.
- Validation: 61 animation API tests and all 4 browser animation tests pass. A targeted browser rerun additionally checks that dropping within the selected range does not add history. Coverage includes noncontiguous selection, arbitrary destinations, saved metadata, rollback after a partial move, real multi-row dragging and undo/redo. Root/dev typecheck/build, targeted lint/format and whitespace checks pass.
- Installed Mac drag visuals/behavior remain unverified while native control is unavailable. Exact operation/UI parity remains incomplete.

### Animation pane tail drop and overview audit

Dragging selected animation rows into empty space below the list now moves them
to the end, using the same atomic reorder/history operation as row drops. The
last row displays the insertion marker; leaving the pane clears it. Drops over
pane controls, outside the pane, or from unrelated drags are not accepted.
The overview now distinguishes current implemented operations from historical
gaps that subsequent entries already addressed.

Cross-area verification: 118 editor, table-cell, text-case/search, smart-guide and
picture-crop tests pass. Native CUA was retried and returned `Transport closed`;
no reference document was modified, and native parity remains unverified.

Validation: all four animation browser tests pass, including a real multi-row
drag to pane whitespace, selection retention, marker cleanup and undo restoration.
Dev TypeScript/build, targeted lint/format and whitespace checks pass.

### Group descendant model for in-group editing

The editor model now exposes direct `children` on groups recursively, while
keeping the slide's top-level `shapes` list unchanged. Descendant bounds remain
in parent coordinates; `parentTransform` supplies the ancestor affine transform
into slide coordinates, including nested scale, rotation and reflection through
the existing geometry-context implementation. This avoids flattening groups or
misinterpreting a child's local coordinates as slide coordinates.

Version-1 history fingerprints still use the original shape fields. Regression
coverage verifies nested hierarchy, a rotated parent's coordinate transform,
child text/format edits through PPTX export and history replay, and restoration
when the edit is undone. This is model groundwork: selecting descendants in the
canvas/Selection Pane and transforming interactive child gestures are still
pending, so in-group UI parity is not yet implemented.

Validation: all 99 editor tests pass; the new descendant regression passes again
after the final type correction. Dev TypeScript/build and targeted lint/format
checks pass. A singular parent transform is represented as `null`, so a future
interaction layer can reject a noninvertible group rather than use an incorrect
identity transform.

### Selecting descendants without ungrouping

Selection Pane groups now expand recursively. Choosing a descendant keeps its
ID selected through updates and routes commands to that child; parent/child
combinations are removed from a selection to avoid double transforms. Selecting
a group on the canvas and clicking a child within its frame enters that child,
using inverse ancestor transforms and local rotation for hit testing. Child
selection overlays and text overlays use parent affine transforms. Pointer
move/resize deltas, rotation angles and keyboard nudges convert slide coordinates
to parent coordinates; singular transforms disable child selection. Select All
returns to the slide's root objects.

The existing snapping path remains limited to root selections, so child moves
currently bypass smart snapping. In-group connector attachment, table-cell
interaction, rich text measurement under nonuniform parent scaling, clipboard,
arrange/distribution, child deletion and native group hit-test details need
further focused verification. This increment does not establish full in-group
operation parity or exact native Selection Pane appearance.

Child SVG wrappers now retain `data-pptx-shape-id`, allowing direct text editing
and preview lookup to target a group descendant rather than its parent. Text-fit
measurement lookup also includes descendants.

Validation: the canvas/history browser regression passes with expanded-pane
child selection, slide-direction keyboard movement inside a rotated group,
child text editing and undo, parent reselection and ungrouping. The descendant
model regression and both grouped-shape renderer tests pass. Preview/dev type
checks and builds, targeted lint/format and whitespace checks pass. Canvas
second-click hit testing and pointer resize/drag still need dedicated browser
coverage; no native Mac UI equivalence is claimed from these tests.

### In-group pointer gestures and nested grouping

A dedicated browser regression now exercises entering a child by clicking an
already-selected rotated group, pointer movement in slide coordinates, resizing
along projected parent axes, unchanged parent/sibling geometry, undo and reload.
The test exposed that nested grouping was still rejected by the API and that
post-command selection considered only top-level shapes.

Grouping now accepts siblings within the same parent group, and ungrouping
reinserts children into that same parent. Mixed-parent selections remain invalid.
The editor tracks new descendants after commands, selects only the outermost
new objects, and selects the immediate released children after ungrouping.
The browser test covers nested grouping/ungrouping and the resulting selections.
The API regression covers preservation of local geometry and exact slide XML
after grouping/ungrouping under a rotated, nonuniformly resized parent.

Validation: 12 grouping API tests and the dedicated group browser test pass;
the existing canvas/history browser regression passes after the selection change.
Root/dev type checks and builds, targeted lint/format and whitespace checks pass.
The native PowerPoint connection retry still returns `Transport closed`.
This does not verify ungrouping a group whose own rotation/flip changes after
creation, nor full native hit-testing or nonuniformly transformed text rendering.
All 99 editor persistence/history tests also pass with the updated grouping API.

### Retaining group rotation and reflection on ungroup

Ungroup now projects each child's center through the removed group's scale,
reflection and rotation, and composes the group's rotation/flip into the child
transform. Previously the API applied translation and scale only, visibly
moving/reorienting children when a rotated or reflected group was removed.

Four API regressions compare all child corners before/after ungroup for rotated,
uniformly scaled groups with no reflection, horizontal reflection, vertical
reflection and both reflections, including an already rotated/reflected child.
Export/reload retains the composed rotations and flips. All 16 grouping API tests
pass; root/dev type checks and builds and targeted lint/format checks pass.
Nonuniform parent scale combined with a rotated child can introduce shear that
an ordinary shape transform cannot express; exact native handling of that case,
and preservation of inherited text/stroke scaling, remain unfinished.
The group browser regression also passes after adding rotated-parent ungroup,
child-center/rotation checks and undo restoring the full group model.

### Descendant stacking and deletion

Grouping now orders members by their existing sibling stacking order rather
than the selection argument order. Bring to Front, Send to Back, Bring Forward,
Send Backward, z-index reads/writes and deletion now locate the actual parent
shape container, so they work within a group without moving an object outside
it or silently ignoring the command. A deleted/stale handle cannot be reinserted
by setting its z-index. Parent lookup is shared with grouping/ungrouping.

Validation: 57 grouping, stacking, mutation and media tests pass. The descendant
stacking regression checks every ordering operation, reversed grouping arguments,
parent/external sibling positions, deletion, stale handles and export/reload.
The group browser regression also passes with child Bring to Front, deletion,
and undo restoring the full parent model. Root/dev type checks and builds,
targeted lint/format and whitespace checks pass. Native PowerPoint visual parity,
last-child deletion semantics and multi-selection ordering remain to be verified.

### Stable multi-selection stacking

New editor ordering commands compute the desired sibling sequence before applying
it, preserving selected objects' relative stacking order regardless of selection
order. Front/back partitions retain sibling order; forward/backward move selected
objects across adjacent unselected siblings without reversing objects at the
container boundary. Each immediate parent is handled independently. A command
flag preserves the previous replay behavior for existing journal entries.

Validation: the editor regression passes 12 root/group scenarios covering all four
operations, boundary cases, history replay and undo. The group browser regression
passes with reverse selection order and Bring Forward at the top boundary,
including selection preservation. Dev type checking/build, targeted lint/format
and whitespace checks pass. Native Mac PowerPoint equivalence still requires
comparison; this change does not establish complete UI or operation parity.

### Duplicating descendants within their parent

New duplicate commands retain the original immediate parent, order copies by
source stacking rather than selection order, and convert the 12 pt slide-space
offset back into the parent's coordinates. Previously descendants were copied to
the slide root with their local coordinates interpreted as slide coordinates.
The copy API exposes an opt-in sameParent option, restricted to an attached source
on the same slide. Existing journal commands retain their previous behavior.

Validation: a regression with a rotated, nonuniformly scaled parent checks both
copies' slide-space offsets, unchanged originals/parent, stacking, save/reload,
history replay and undo. The group browser regression passes with Command-D,
selection of the new descendant and undo. Root/dev type checks and builds,
targeted lint/format and whitespace checks pass. Native connection retry still
returns Transport closed. Clipboard extraction of descendants and native
comparison of duplication behavior remain unfinished.

### Keyboard traversal of grouped objects

Canvas Tab/Shift+Tab now resolves the selected descendant's immediate parent and
cycles its eligible children instead of losing the selected ID in a root-only
list. Starting from no selection chooses the first/last object explicitly;
previously reverse traversal skipped the last object because its starting index
was -1. Singular descendant transforms remain excluded.

Microsoft's Mac shortcut reference documents Tab/Shift+Tab for object selection:
https://support.microsoft.com/en-us/accessibility/powerpoint/use-keyboard-shortcuts-to-create-powerpoint-presentations
It does not specify the group traversal boundary, so sibling-only cycling is an
implementation inference awaiting installed Mac PowerPoint comparison.

Validation: the group browser test passes with forward/reverse child traversal,
wraparound, unchanged document revision and first/last root selection from no
selection. Dev type checking/build, targeted formatting/lint and whitespace
checks pass. This does not establish exact native group keyboard behavior.

### Arrange operations across group coordinate spaces

Alignment and distribution now measure transformed shape frames in slide
coordinates, including shape rotation/reflection and ancestor matrices, then
convert slide-space translations back to each shape's parent coordinates.
Previously local child coordinates were compared directly, producing moves along
rotated/scaled parent axes. Local sizes, rotations and parent bounds are retained.
The shared arrangement module excludes singular transforms and rounds final
positions to EMUs. Existing single-selection slide alignment remains supported.

Three geometry regressions cover independently calculated projected extents,
all six alignments (selection and slide), horizontal/vertical distribution,
fixed outer objects, equal gaps, unchanged perpendicular positions, and
nonmutation. Dev type check/build and targeted lint/format/whitespace checks pass.
Exact Mac behavior for irregular shapes, connector extents and alignment reference
modes still needs native comparison; frame-based extents are not a proof of full
native geometry equivalence.
The existing canvas/history browser suite passes. The group browser regression
also passes with Align Top on rotated children, unchanged parent bounds and undo
restoring the full model; the subsequent keyboard step explicitly restores canvas
focus after the toolbar interaction.

### Alignment reference controls

The Align submenu now includes checked Align to Slide and Align Selected Objects
choices. The latter is disabled for fewer than two selected objects; a single
object uses the slide automatically without overwriting the multi-selection
preference. Multiple objects can now be aligned individually to slide edges or
center. The preference applies to the current editor session. Slide distribution
supports two or more objects and includes equal outer margins; selected-object
distribution retains its existing three-object minimum and fixed outer objects.

Microsoft's macOS PowerPoint instructions confirm the two reference choices and
the selection requirement:
https://support.microsoft.com/en-us/office/graphics-visuals/align-or-arrange-objects
The precise slide-distribution margin calculation and preference persistence
still require installed Mac comparison; the documentation does not prove these.

Validation: all five arrangement unit tests pass, including explicit multi-object
slide centering and distribution. The group browser regression passes with mode
switching, checked state, rotated children centered on the slide, undo restoring
the full parent model and switching back to selected objects. Dev type checking,
build and targeted lint/format/whitespace checks pass.

### Smart guides within transformed groups

Dragging group descendants now participates in smart alignment and drawing-guide
snapping. Moving and target frames are measured through the shared slide-space
extent calculation, including ancestor rotation and scale. Candidate collection
traverses ancestors of the moving selection to include sibling objects while
excluding selected subtrees and their containing groups; unrelated groups remain
single targets. Existing inverse-parent delta conversion applies the snapped
slide-space displacement without changing the parent frame. Modifier bypass and
axis constraints remain available.

Validation: all 15 smart-guide and arrangement unit tests pass, including explicit
rotated/scaled parent coordinates and nested candidate filtering. Both browser
suites pass: group dragging displays a guide and aligns the child's projected
center with its sibling, preserves parent/sibling geometry and supports undo;
the existing canvas/history/notes/reload suite also passes. Dev type checking,
build and targeted lint/format checks pass. Native CUA reconnection still returns
`Transport closed`; precise Mac guide candidate scope and visual parity remain
unverified despite authorization for temporary reference-document changes.

### Clipboard extraction of group descendants

New UI paste commands carry `pasteSlideCoordinates`. For snapshot pastes, the
private clipboard presentation releases ancestor groups from outside inward
before importing the selected objects into the destination slide. This applies
ancestor placement/rotation/scale instead of interpreting a child's local bounds
as slide coordinates. Selected groups remain groups unless their descendants are
also explicitly selected. The original document is unchanged. The 12-point
same-slide copy offset is applied after projection; cut/cross-slide paste retains
its no-offset path. Legacy commands without the flag retain replay behavior,
and the flag requires a snapshot so live source objects cannot be ungrouped.

Validation: three snapshot paste tests pass, including nested rotated ancestors,
uniform parent scaling, both offset modes, unchanged source hierarchy, replay,
serialization and undo. The group browser test passes with copied child text,
projected center, inherited rotation, unchanged parent and undo. The existing
object clipboard browser test also passes, including immediate cut, snapshot
contents, reload and pending-paste slide navigation. Dev type checking/build,
targeted lint/format and whitespace checks pass.

This reuses the ungroup transform logic and shares its unresolved cases:
nonuniform parent scale combined with rotated children can require shear, and
inherited text/stroke scaling needs additional handling. Native paste selection
context (root versus entering an existing group) and exact Mac appearance still
need installed-app comparison. This is not complete clipboard/UI parity.

### Command-click in the Selection pane

The Selection pane now extends/toggles selection on Command-click, instead of
replacing it. Shift and Control also extend/toggle. Plain clicks still replace
the selection. Ancestor/descendant exclusion applies before adding the target,
so Command-clicking a parent replaces selected descendants and selecting a child
removes a selected ancestor. This is selection state only, not a document edit.
Microsoft's macOS Selection pane documentation explicitly describes Command-click
for multiple objects:
https://support.microsoft.com/en-us/powerpoint/use-the-selection-pane-to-manage-objects-in-documents

Validation: the group browser regression passes with Command-click addition and
removal, pressed-state feedback, ancestor/descendant switching, unchanged revision
and the subsequent group alignment/editing flow. Dev type checking/build, targeted
format/lint and whitespace checks pass. Exact Mac behavior for mixed-parent
selections and returning from a child selection with Escape remain unverified;
no speculative change was made to Escape handling.

### Animation markers for group descendants

Animation marker rendering and marker-driven selection now resolve targets through
the full shape hierarchy. Marker placement uses the shared projected slide extent,
including ancestor transforms, rather than root-only bounds. The animation list
resolves descendant names and includes those names in its update signature, so a
child target no longer falls back to its numeric ID. Invalid/singular extents are
excluded from marker placement.

Validation: five browser tests pass across the group editing and animation playback
suites. The new group regression authors a child effect, verifies a numbered
marker at the independently projected rotated bounds, selects the child by clicking
the marker after selecting its sibling, checks the child's name in the animation
pane, then undoes the effect and verifies marker removal and unchanged shapes.
Existing playback, timing/selection, preview and show navigation checks also pass.
Dev type checking/build, targeted formatting/lint and whitespace checks pass.
Native CUA reconnect again returns `Transport closed`; exact Mac marker appearance
and the wider missing animation effects/options remain outstanding.

### Inline object names in the Selection pane

Object names can now be edited inline by double-clicking their Selection pane row
or pressing F2 on it. Enter or blur commits the name; Escape cancels. Updates use
the shape name API and the existing revision-checked edit journal, targeting the
captured slide and shape instead of the current selection. Text content is not
changed. Detached fields and navigation to another slide do not submit a rename.
The animation name signature already includes descendant names.

Validation: all 102 editor tests pass. The group browser regression passes with
inline entry by double-click, cancellation, F2/Enter rename, unchanged slide text,
reload persistence and undo restoring the original hierarchy. The first browser
attempt correctly found the Selection pane hidden after opening Animation Pane;
the test now reopens Selection before interacting. Type checking/build and lint
pass. Exact installed Mac activation gestures, focus restoration and empty-name
behavior still need native comparison; this adds the missing name-editing path
without claiming complete Selection pane parity.

### Selection pane visibility and hidden shape rendering

- Added per-object Show/Hide controls to the Selection pane. The command persists `p:cNvPr/@hidden` without removing the shape, and supports history replay and undo.
- The editor model exposes hidden state. Preview rendering omits hidden objects and their grouped descendants; canvas hit overlays, child hit testing, Tab traversal, animation markers, and smart-guide targets exclude hidden objects.
- Validation: preview and dev TypeScript/build passed; existing editor and smart-guide tests passed (112 tests); added hidden-group export/render/replay/undo test passed; group browser test passed including hiding a group, reload persistence, and undo restoration. Targeted lint and diff checks passed.
- Native Mac comparison remains unavailable because the computer-use transport is closed. Exact eye icon appearance, hidden-object selection behavior, bulk visibility controls, and native shortcut semantics are not yet verified. This is a functional increment, not evidence of full PowerPoint parity.

### Bulk visibility in the Selection pane

- Added Show All and Hide All controls above the object list. They include descendants of collapsed groups, leave object contents intact, and record each bulk action as a single update. Already matching objects are omitted; clicking an already satisfied action does not create history.
- Extended the group browser scenario to test individual child Hide/Show, bulk hiding, showing descendants, and two undo steps that restore first the fully hidden state and then the original mixed visibility state.
- Validation: dev TypeScript check and build passed; targeted lint and formatting passed; group browser scenario passed (11.4 seconds); git diff --check passed.
- Microsoft macOS Selection pane documentation supports individual hide/unhide and retained file contents: https://support.microsoft.com/en-us/powerpoint/use-the-selection-pane-to-manage-objects-in-documents . The exact current Mac bulk-control styling and treatment of child flags still require native comparison. CUA getState was retried and again returned Transport closed. Full-operation parity remains incomplete.

### Drag reordering in the Selection pane

- Object-name rows now support drag/drop before or after another sibling, with an insertion indicator. The displayed reverse stacking order is translated into the underlying XML order. Selected sibling objects move together while retaining their internal stacking order.
- Added an explicit relative placement command. It validates that the target is unselected and every moved object has the same parent; it uses the existing in-parent z-order API and records a single replayable history step. No reparenting is performed.
- Validation: dev typecheck/build passed; targeted lint, formatting and diff checks passed. New unit test covers multi-object relative placement, preserved parent, save/replay/undo and rejection of a cross-parent target. Group browser scenario passed (19.2 seconds), including dragging a child row above its sibling, reloading persisted order and undoing to the exact preceding model.
- Mac reference: Microsoft describes dragging an object upward/downward in the Selection pane to change stacking order at https://support.microsoft.com/en-us/powerpoint/use-the-selection-pane-to-manage-objects-in-documents . Exact native drag feedback, scrolling during drag and multi-selection edge cases are still unverified with the installed Mac app; native transport remains unavailable. This does not establish full UI parity.

### Hidden shape action and slideshow verification

- Audited the path from shape visibility through SVG rendering to slideshow hotspots. Hidden content is omitted before object/run links, hover actions and action sounds are emitted; slideshow keyboard navigation enumerates only rendered hotspots. No additional product change was needed for this path.
- Added rendering regression tests for both SVG text and foreignObject text: individually hidden group child, independently visible sibling/image, hidden entire group plus image after export/reload, and restoration after Show. Checks cover run hyperlinks, object click links, hover links and click/hover sound-stop triggers. Generated mask IDs are normalized only when comparing restored rendering.
- Extended the hyperlink browser scenario to hide an action button through the Selection pane, start the slideshow, traverse and wrap Tab focus without that button, then Show it and verify the original hotspot list returns.
- Validation: all 13 preview action tests passed; hyperlink browser scenario passed (16.9 seconds); targeted lint, formatting and diff checks passed. Native Mac visual/behavioral parity still requires the unavailable app comparison; these tests establish behavior in this implementation only.

### Selection pane drag auto-scroll

- Added frame-based scrolling near the top and bottom edges while dragging a Selection pane object. Speed increases toward the edge; leaving the pane, dropping, ending the drag, hiding the pane, or rebuilding its contents stops the scroll loop.
- Added a browser scenario with 45 objects: drag the front object to the offscreen back of the list, verify the resulting order after reload, then undo and compare the full original model.
- Validation: dev TypeScript/build passed; new browser scenario passed (2.9 seconds); existing group editing browser scenario passed (13.9 seconds); targeted lint/format and diff checks passed.
- The first browser runs timed out because the test incorrectly waited for Undo to remain enabled after undoing the only edit. The scenario now waits for Redo and checks the restored model.
- Exact Mac PowerPoint edge thresholds, scroll speed, and insertion feedback remain unverified. Native object lock semantics are also pending direct app verification; no speculative lock control was added.

### Selection pane docking and header

- The Selection pane now reserves 250 px at the right of the editing stage and notes, so it no longer covers the slide. Closing it restores the stage width; the existing ResizeObserver recalculates the slide fit.
- Changed the header to “Selection Pane” with an accessible close icon alongside it. Object names truncate within their available width and expose the full name as a tooltip, keeping visibility controls within the pane.
- Reference: Microsoft Support macOS Selection pane image, https://support.microsoft.com/en-us/powerpoint/media/word-mac-selection-pane-selected-object.png . This supports the header arrangement, not an exact current-version measurement.
- Validation: dev TypeScript/build, targeted lint/format, and diff checks passed. Selection pane browser scenario now verifies non-overlapping stage/pane bounds and restored width on close; it passed along with group editing (2 browser scenarios, 14.6 seconds total).
- Native UI transport was retried and still reports “Transport closed.” Installed Mac version appearance, pane dimensions, and full interaction parity remain unverified.

### Native menu access and Rotate or Flip

- Recovered read access to installed PowerPoint through elevated AppleScript/System Events despite the unavailable CUA transport. Version is 16.113.1. PowerPoint reports the recovered `support [Autosaved]` presentation. AX window enumeration remains empty, but application document-window properties and menu trees are readable.
- Created a new empty `Presentation1` solely for inspection, then closed that exact presentation without saving. No edits were made to the recovered support document.
- Observed native Arrange ordering: Bring to Front, Send to Back, Bring Forward, Send Backward; Group, Ungroup, Regroup; Rotate or Flip, Align or Distribute; Selection Pane.... Rotate or Flip contains Rotate Left 90°, Rotate Right 90°, Flip Horizontal, Flip Vertical, and More Rotation Options.... Alignment submenu matches the existing six alignments, two distributions, and slide/selected-object modes.
- Added an application Arrange menu and a Rotate or Flip submenu shared with the ribbon/context arrangement controls. Native application menu uses “Align or Distribute”; the existing ribbon label remains “Align.” Selection Pane... opens the pane. Empty selections disable geometry commands.
- Added relative horizontal/vertical flip commands so mixed selections toggle each object's stored flag independently. Rotation uses the existing relative-angle command. All four menu actions support journal history; More Rotation Options opens the current shape properties pane.
- Validation: two targeted editor tests (relative rotation/flips), group browser integration including all four menu commands and full-model undo (17.0 seconds), dev TypeScript/build, targeted lint/format and diff checks passed.
- Regroup remains absent, and native visual rendering/selection geometry must still be compared. Menu-tree observations prove labels/order only, not pixel or complete operation parity. Native app reads via AppleScript now provide another avenue for future comparison.

### Regroup session membership and command

- Compared installed Mac PowerPoint 16.113.1 using newly created, disposable presentations. Group → Ungroup → move one child → Regroup produced one group with both members, retained the moved position, and assigned a new group name (`Group 4`, not the former custom name). Closing and reopening the saved ungrouped document left Arrange > Regroup disabled; calling the AppleScript command then did nothing (two shapes before and after).
- Added session-only regroup membership keyed by XML element identity, avoiding accidental membership when deleted shape IDs are reused. A selected former member can restore its available siblings; only the first eligible former group in a selection is restored. Grouping consumes that membership. Detached objects and members now in other parents are excluded.
- Added Arrange > Regroup with eligibility based on the current selection, and selection of the restored group after the action. Editor journal replay reconstructs eligibility across browser reloads and undo; exported PPTX does not carry editing-session membership.
- Microsoft API reference for the first eligible group rule: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.regroup . Exact Mac UI behavior for multiple former groups, partially deleted groups, and regrouping across nested groups still needs comparison.
- Native AppleScript returns some unusable object references and may silently no-op, so observed shape counts and menu enabled state were used instead of command success alone. All temporary presentations were closed without changing the recovered support document. The native saved audit fixture is at `/Users/baseballyama/Library/Containers/com.microsoft.Powerpoint/Data/tmp/pptx-regroup-audit.pptx`; PowerPoint resolved the requested `/tmp` path inside its container.
- Validation: 18 core grouping tests passed; eight targeted editor tests passed, including geometry preservation, export, history replay, and undo. Browser integration and final checks are recorded below once finished. Full PowerPoint operation/UI parity remains incomplete.
- Browser integration initially exposed lost session membership during the server's internal PPTX save/load. Added validated session restoration in both build replay and edit handling, while continuing to render serialized output. The editor test now explicitly covers that internal round trip separately from a fresh PPTX import.
- Final validation: group browser integration passed (27.6 seconds), including Ungroup, reload, single-member selection, Regroup, and full-model Undo. The 18 core and eight targeted editor tests passed after the fix; root/dev typechecks and builds, targeted lint/format, and diff checks passed.

### Native Arrange keyboard equivalents

- Read the installed PowerPoint Arrange menu's `AXMenuItemCmdChar` / `AXMenuItemCmdModifiers`: Group G/2, Ungroup G/3, Regroup J/2, Bring to Front F/1, Send to Back B/1, Bring Forward F/3, Send Backward B/3. These correspond to Option-Command-G, Option-Shift-Command-G, Option-Command-J, Shift-Command-F/B, and Option-Shift-Command-F/B.
- Added Regroup's Option-Command-J handler using physical key codes so Option-produced characters do not prevent recognition. Displayed the native keyboard equivalents alongside the seven Arrange commands; Group/Ungroup now also disable while an edit is pending.
- Arrangement shortcuts remain available with focus in the application menus, ribbon, or Selection pane after dismissing a menu. Text inputs, text editing, dialogs and thumbnail selections continue to use their own handlers. Other canvas-only keyboard handling keeps its original focus restriction.
- Added a browser check for the displayed Regroup shortcut, menu dismissal, keyboard Regroup and exact-model Undo. Dev TypeScript/build, targeted lint/format and diff checks passed; browser result follows.
- Browser integration passed (17.1 seconds), including the new shortcut and undo assertions. Complete Mac PowerPoint UI/operation parity is still not established.

### Native View menu entry points

- Read the installed Mac PowerPoint View menu and its Grid and Guides / Zoom submenus through accessibility. Confirmed Normal Command-1, Slide Sorter Command-2, Presenter View Option-Return, Slide Show Shift-Command-Return, and Ribbon Option-Command-R.
- Added View between Edit and Format in the application menu. Connected the existing Normal, Slide Sorter, Presenter View, Slide Show, Ribbon, Smart Guides, Guides and Zoom operations in native order. The Zoom submenu exposes Fit to Window, Zoom In, Zoom Out and Zoom....
- Centralized Normal/Slide Sorter state updates so ribbon and status controls agree with menu commands and Command-1/2. Added Option-Command-R to toggle the ribbon. Existing presentation shortcut handlers remain in use.
- Validation: dev TypeScript and build, targeted formatting/lint, diff checks and a new browser integration test passed (1.7 seconds). The test covers menu placement, checked state, view shortcuts, ribbon shortcut, Smart Guides toggle and Zoom dialog entry.
- This is partial menu coverage, not complete View parity. Notes Page, Outline View, Reading View, Master, Ruler, Gridlines, Snap to Grid and Grid Options remain absent; native visual equivalence is not established.

### Grid options, document spacing and movement snapping

- Previous goal turn made implementation progress (View menu and browser verification); this turn continued missing View operations.
- Inspected installed Mac PowerPoint's Grid and Guides sheet in a disposable presentation with a blank slide. A presentation without slides did not open the sheet. Captured `/tmp/pptx-grid-options.png`. Native groups are Snap to, Grid Settings, Guide Settings; labels are Snap objects to grid, Spacing:, Display grid on screen, Display drawing guides on screen, Display smart guides when shapes are aligned. Initial checks were off/off/off/on, spacing 5 grids per cm.
- Native spacing choices: 8, 6, 5, 4, 3, 2 grids per cm; 1cm through 5cm; Custom. Custom exposes a centimeters input. Native also has Set as Default, Cancel and OK. Set as Default is not implemented yet.
- Applied grid visibility and snap in the disposable document, verified both checkboxes remained 1 on reopening, and inspected its saved PPTX. `/ppt/viewProps.xml` contained `p:gridSpacing cx="72008" cy="72008"`; the enabled flags were not explicitly emitted in this fixture. Restored both temporary toggles and closed the disposable document without further saving; recovered support was untouched. Saved reference remains in PowerPoint's container `Data/tmp/pptx-grid-audit.pptx`.
- Added getGridSpacing/setGridSpacing with independent EMU axes, related view-part handling, preservation of other view settings and guide metadata, and validation before mutation. Refactored the existing guide visibility writer to share the view-part update path. Editor command/history supports spacing and guide visibility together.
- Added View > Grid and Guides > Gridlines, Snap to Grid and Grid Options..., plus a Grid Options entry in existing guide menus. The dialog implements the observed groups, labels, spacing presets, Custom, Cancel and OK. Spacing is document data; grid visibility, snap and smart guide settings currently persist in browser preferences. Exact native preference scope/default behavior remains to be compared.
- Added a grid overlay and movement snapping of the selection bounding box; Shift constraint and existing Alt/Command bypass remain respected. Existing object/drawing-guide alignment takes priority over grid snapping. Native priority, origin, resize/draw snapping, very fine-grid rendering, and pixel equivalence still need comparison. Overlay pitch is currently clamped to two screen pixels at tiny zoom/spacing values.
- Validation: root/dev TypeScript and builds passed. Five core guide/spacing tests and eleven snapping tests passed. Browser integration passed (3.0 seconds): dialog cancellation, apply, reload, preferences, document spacing/guide visibility, real drag snapping, and Undo for drag and settings. Targeted formatting, lint and diff checks passed. The initial browser failure was a test-only incorrect Undo selector, corrected to the application's existing data-edit selector.
- Complete Mac PowerPoint UI/operation parity remains incomplete, including Set as Default and the broader View/editing gaps recorded above.

### Grid spacing defaults and custom precision

- Compared Set as Default in disposable native presentations. Clicking it leaves the sheet open. Selecting 1cm, clicking Set as Default and Cancel left the next presentation at 5 grids per cm; repeating with OK made the next presentation use 1cm. This establishes the commit/cancel behavior for spacing, not the scope of all guide checkboxes.
- Added Set as Default at the left of the dialog footer. OK persists the spacing default in browser storage; Cancel discards the pending default change. Documents with explicit spacing retain it; documents without stored spacing use the browser default. Invalid stored defaults are ignored, and unavailable storage falls back to the current session.
- Custom spacing now preserves changes smaller than the preset matching tolerance: 0.50001cm stores 180004 EMU instead of being ignored as equivalent to 180000. Preset matching continues to accommodate native 72008 EMU spacing.
- Validation: dev TypeScript/build and targeted lint/format/diff checks passed. Browser integration passed (4.4 seconds), covering the default button remaining open, cancellation, confirmation, undo and reload with fallback defaults, and custom precision. Browser screenshot: `/tmp/pptx-grid-browser.png`; native sheet screenshot: `/tmp/pptx-grid-default.png`.
- Native cleanup completed after the user unlocked the Mac. Restored Presentation12 to snap/grid/drawing/smart = 0/0/0/1 and 5 grids per cm, clicked Set as Default, then OK. Created Presentation13 and independently read the same four values and spacing from its sheet, verifying restoration. Closed disposable Presentation10–13 without saving; PowerPoint reports only support [Autosaved] remaining. The support document was not edited.
- Full default behavior for guide checkboxes, exact native visuals, and the broader UI/operation parity remain unfinished.

### Document grid snapping versus application defaults

- Previous goal turn made implementation progress and completed native cleanup. Continued by comparing preference scope in disposable Presentation14–17. In Presentation14, toggled snap/grid/drawing/smart to 1/1/1/0 and confirmed with OK without Set as Default. New Presentation15 read 0/1/1/0. After enabling snap in Presentation15 and using Set as Default then OK, new Presentation16 read 1/1/1/0. Thus snap has a separate explicit default, whereas the other three checks carried into a new document after ordinary OK in this experiment. Existing-document scope of the other three checks still requires comparison.
- Restored all four baseline checks to 0/0/0/1 using Set as Default and OK. New Presentation17 independently read 0/0/0/1 and 5 grids per cm. Closed only the disposable documents; PowerPoint reports only support [Autosaved] remaining, untouched.
- Saved the snap-off fixture at PowerPoint container `Data/tmp/pptx-snap-off-audit.pptx`. Its slide view has `snapToGrid="0"`; the earlier enabled fixture omitted the attribute. The saved grid spacing was 72010 EMU after native setting round trips, despite the displayed preset remaining 5 grids per cm. Guide XML had showGuides=1 despite the restored UI checkbox, so guide display persistence warrants a separate audit.
- Added getSnapToGrid/setSnapToGrid. No slide-view preferences returns null; stored slide-view preferences follow the native implicit enabled value unless explicitly false. Writes preserve other preferences; newly created slide-view preferences explicitly disable snapping until chosen. Attribute reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.commonslideviewproperties.snaptogrid .
- Moved Snap to Grid from shared browser preferences into document commands, PPTX serialization, editor model and conflict-aware history. Grid Options applies snapping together with spacing and guide visibility. Set as Default now stores snapping alongside spacing; explicit document settings override defaults. Guide edits preserve the current snapping setting when creating view properties.
- Validation: six core guide/spacing/snapping tests and three targeted editor history tests passed; browser View integration passed (8.0 seconds), including reload, drag snapping, undo, default snapping and an explicit document override surviving reload. Root/dev TypeScript and builds, targeted lint/format and diff checks passed. Dev build retains its existing declaration sourcemap warnings.
- Remaining: drawing-guide preference scope/default initialization, grid origin and resize/draw snapping, exact native visuals, and the wider unfinished PowerPoint UI/operation requirements. Full parity is not established.

### Drawing guide visibility across open documents

- Previous goal turn made implementation progress. Created disposable Presentation18 and Presentation19 before changing either. Presentation19's drawing-guide checkbox changed from 0 to 1. After switching to Presentation18 and verifying its window name, its sheet also read 1. Restored 0 in Presentation18; after closing that document, independently read 0 in Presentation19. Closed both without saving. This establishes that native drawing-guide visibility propagates to an already-open document, not only newly created documents.
- Native background UI calls sometimes returned successful element references without completing window switches. AXRaise alone did not switch the active document. Explicit activation, waiting for UI updates and checking the actual sheet/window were necessary; failed reads were not treated as evidence. Screenshot: `/tmp/pptx-guide-scope.png`.
- Drawing-guide visibility now persists with the application view preferences alongside grid visibility and smart guides. A stored preference overrides document guide visibility; document metadata remains the fallback when no application preference exists. Showing/hiding guides no longer creates document edits. Guide positions, additions and deletion still use document commands/history. Adding a guide enables the application display preference.
- Added storage-event synchronization for these display preferences across already-open same-origin tabs. An explicit false value survives reload. Cross-origin development servers do not share browser storage; this is a runtime limitation of the current app.
- A further disposable Presentation20 confirmed the native Guides menu changed display from off to on; the Undo item remained “Undo Last” and enabled before and after. That label comparison alone does not establish full native undo semantics. Restored the checkbox to 0, reopened the sheet to verify 0, and closed Presentation20 without saving.
- Dev TypeScript/build, targeted lint/format and diff checks passed. View browser integration passed (14.3 seconds), including two-tab propagation, unchanged document revision for guide visibility, reload, Grid Options, and existing snapping/default behavior. Initial test failures came from implicit single-page Playwright contexts, queries that did not traverse Shadow DOM, and an old journal-count expectation that included the two removed display edits; these were corrected.
- Full Mac UI/operation parity remains unfinished, including exact native undo semantics for view preferences, native grid origin, resizing/drawing snap and broader View/editor functionality.

- Follow-up validation: the full canvas editing/history/notes/reload browser integration passed (167.1 seconds). Updated its stale Selection Pane close-button accessible name to “Close Selection Pane” and removed temporary progress logging. Increased this long integration timeout from 90 to 180 seconds after the shorter timeout prevented completion on this Mac; individual operation waits remain unchanged. Targeted formatting/lint and repository diff checks passed. This verifies the exercised operations, not full PowerPoint parity.

### Editor full screen from the View menu

- Published the accumulated editor implementation as 71334e5 on origin/feat/mac-powerpoint-parity. The local main branch was behind origin/main; publishing the feature branch preserves both histories. Local .pnpm-store cache was excluded.
- Read the installed Mac View menu and confirmed Enter Full Screen follows Zoom. AX reports command character F and modifier value 24; no browser keyboard mapping is claimed. Native window-level interaction was unavailable during this comparison, so a complete enter/exit visual comparison remains outstanding. No native document or preference was changed.
- Added Enter Full Screen / Exit Full Screen to the editor View menu, using the browser fullscreen API and disabling the command when that API is unavailable. Fullscreen failures are shown through the existing editor error message.
- A slide show started from an already-fullscreen editor now preserves editor fullscreen when End Show is invoked. A show that enters fullscreen itself still exits fullscreen on completion. This prevents the new editor mode from being lost as a side effect of ending a show; exact native window/Spaces behavior and OS shortcuts remain unverified.
- Validation: dev TypeScript/build and targeted lint/format passed. View browser integration passed (10.0 seconds), including fullscreen enter/exit, no document revision change and preserving fullscreen across a show. Existing Mac slideshow shortcut integration passed (6.0 seconds). Full Mac UI/operation parity remains incomplete.
