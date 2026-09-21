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

- The editor branch incorporates main's development-tool source and current API.
- The generated manifest lists 150 authoring operations. Registry coverage and
  smoke tests verify function discovery and selected API mutations.
- All 150 command labels have English and Japanese versions, enforced by a test.
- Shape text creation and main's bullet-preservation behavior coexist; targeted
  text and registry tests pass.

## Outstanding work

The editor is still a separate site route. Development-preview integration,
persistence, complete workflow coverage, remaining UI translations, and browser
verification are not complete. In particular, current document history coalesces
snapshots while serialization is running and must be audited for rapid edits and
new/open races before relying on it for persistent preview editing.
