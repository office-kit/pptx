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
- Native Lock All writes DrawingML `spLocks` attributes (`noGrp`, `noRot`, `noMove`, `noResize`, `noEditPoints`, `noAdjustHandles`, `noChangeArrowheads`, `noChangeShapeType`). Unlock All removes those attributes. The disposable reference was unlocked and saved after comparison. Lock editing/enforcement is not yet implemented.
- Still missing here: native reorder-overlapping view, slide-relative distribution, pane locks, and complete native geometry/styling. Full-operation parity remains incomplete.
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
