---
'@office-kit/pptx-editor': minor
---

Long-running agents and hosts can hand the editor a version of the deck made elsewhere:

- `propose(base, edited, { label, from })` merges `edited` (made from `base`) with the deck as it is now, including the user's unsaved edits. Changes to different slides, shapes, media and links combine into one undo step. When the same item changed on both sides, the title bar names each collision and the user keeps their edits or takes the proposed version (one undo step, so nothing is lost); the promise resolves with `{ status: 'applied' | 'kept-mine' | 'took-theirs' }` and the conflicts. `from: 'source'` is for a newer version of the file itself: the prompt speaks of the source, `'change'` reports `'source'`, and a result equal to what the host saved is not an unsaved change.
- `@office-kit/pptx-editor/merge` exports the same three-way merge (`mergeDecks`, `describeConflict`) for Node and the browser.
- New options: `fileName`, `autoSave` (the AutoSave switch, saving through `onSave` after each edit), `compact` (a slimmer title bar), `status` (your own element in the title bar) and `isolate: false` (render without a shadow root, for a page that is the editor's alone). `onSave` may resolve `false` to leave the deck unsaved without a message.
- New on the handle: `open(pptx, { fileName, unsaved })`, `save()`, `dirty`, `locale`, and the `'dirtychange'` and `'localechange'` events.
