# @office-kit/pptx-dev

## 0.2.2

### Patch Changes

- 8278aa4: Initialize slide projects with separate slide files, a shared theme and a small deck entry. Guide AI revisions toward focused source patches and the running preview instead of regenerating the deck.
- Updated dependencies [983eb10]
  - @office-kit/pptx-dsl@0.4.0

## 0.2.1

### Patch Changes

- ad3cbd2: Keep slide previews steady while editing TSX: update only changed thumbnails and slides, preserve zoom and scroll position, and retain the previous frame until its replacement is ready. Transfer SVG changes incrementally, preload build workers between saves, and cancel obsolete evaluations so rapid edits and accidental infinite loops do not delay the next revision.
- Updated dependencies [71deec5]
- Updated dependencies [71deec5]
  - @office-kit/pptx@0.17.0
  - @office-kit/pptx-preview@0.9.6
  - @office-kit/pptx-dsl@0.3.0

## 0.2.0

### Minor Changes

- 4fe8fde: Add typed TSX authoring without a React or Vue runtime, including native text,
  shapes, images, tables and charts, template slide reuse and explicit editing,
  and a Raw callback escape hatch. Add local project initialization, PPTX export,
  template inspection and watch preview with error recovery, plus VSCode tasks
  and Claude Code authoring guidance.
- a9a2598: Add a view-only slide workspace with vertical thumbnails, keyboard navigation,
  fit/zoom and full-screen presentation. Live TSX updates retain the selected slide;
  all content changes continue to happen in source rather than on the canvas.

### Patch Changes

- Updated dependencies [4fe8fde]
- Updated dependencies [b8adc13]
- Updated dependencies [a13afc2]
  - @office-kit/pptx-dsl@0.2.0
  - @office-kit/pptx@0.16.1
  - @office-kit/pptx-preview@0.9.5
