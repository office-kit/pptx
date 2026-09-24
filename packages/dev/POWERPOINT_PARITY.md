# Mac PowerPoint parity integration

The final review target is PR #287 (`feat/pptx-editor`). No additional PR is required.

The earlier implementation remains preserved on `feat/mac-powerpoint-parity` at `ef67cf0`. It and #287 developed different UI, persistence, animation and rendering implementations after their common base. A trial merge produced 39 conflicted files and was aborted without discarding either branch. Transfer compatible features individually, retaining #287 behavior and testing each integration. The earlier branch is not fully merged.

## Integrated: presentation guides and grid settings

- Public APIs read/write drawing guides, stored guide visibility, grid spacing and snapping. Preserve related view parts, surviving guide metadata and unrelated view settings. Extended guide positions use native master units; the public API uses EMU.
- Registered the four mutations in the existing editor command catalogue, with English/Japanese labels and structured fields for guides and grid spacing. These edit document settings; this does not yet add the old branch's rendered guide overlay, drag interactions, application visibility preferences or native Grid Options dialog to the #287 canvas.
- Six API tests and six capability/localization coverage tests pass. Root TypeScript and targeted lint/format checks pass.

## Outstanding

Port the guide UI and appropriate view preferences to the current editor, then reconcile the remaining operations from the earlier branch. Continue native visual/interaction comparison. All-operation Mac PowerPoint UI parity remains incomplete. The older branch contains the detailed comparison history in its version of this file.
