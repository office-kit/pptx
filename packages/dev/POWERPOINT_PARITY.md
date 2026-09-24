# Mac PowerPoint parity integration

The final review target is PR #287 (`feat/pptx-editor`). No additional PR is required.

The earlier implementation remains preserved on `feat/mac-powerpoint-parity` at `ef67cf0`. It and #287 developed different UI, persistence, animation and rendering implementations after their common base. A trial merge produced 39 conflicted files and was aborted without discarding either branch. Transfer compatible features individually, retaining #287 behavior and testing each integration. The earlier branch is not fully merged.

## Integrated: presentation guides and grid settings

- Public APIs read/write drawing guides, stored guide visibility, grid spacing and snapping. Preserve related view parts, surviving guide metadata and unrelated view settings. Extended guide positions use native master units; the public API uses EMU.
- Registered the four mutations in the existing editor command catalogue, with English/Japanese labels and structured fields for guides and grid spacing. The canvas integration below uses these document settings.
- Six API tests and six capability/localization coverage tests pass. Root TypeScript and targeted lint/format checks pass.

## Outstanding

Reconcile the remaining operations from the earlier branch. Continue native visual/interaction comparison. All-operation Mac PowerPoint UI parity remains incomplete. The older branch contains the detailed comparison history in its version of this file.

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

- Status percentage opens the Zoom dialog. Added the two-segment zoom slider (10–100–400%), keyboard percentage changes, Home/End and 10-point increment/decrement controls from the earlier implementation. Native exact stepping still needs direct comparison.
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

- Integrated the native key equivalents recorded in the earlier branch: Option-Command-G / Option-Shift-Command-G, Shift-Command-F/B and Option-Shift-Command-F/B. Physical key codes avoid Option-generated characters. Existing Control/Command-G compatibility remains.
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
- Picture fill controls now expose image insertion, transparency, tile/stretch offsets, scaling, alignment, mirroring and rotation, with multi-selection edits in one undo transaction. Initial selection opens a file chooser; native default texture selection and visual comparison of this pane remain outstanding. Image insertion also remembers the preceding solid, gradient or pattern fill so switching back restores its settings. Image fills are remembered across type switches, including media, tile/stretch placement, opacity and supported crop values. Mac PowerPoint restored 35% transparency after switching to gradient and back; all four temporary edits were undone and the reference saved. Imported negative/outset crops, explicit DPI overrides and arbitrary image effects are not yet preserved by this cache.
- Texture selection, rotation-independent image rendering, native comparison of source-cropped tiles, additional image formats/EXIF resolution and texture selection remain outstanding.

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
