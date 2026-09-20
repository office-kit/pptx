# @office-kit/pptx-dev

## 0.4.1

### Patch Changes

- 1ca6ad7: Resize the preview chat panel by dragging its left edge or using the keyboard, with the chosen width remembered across reloads.

## 0.4.0

### Minor Changes

- 7982fc9: Embed interactive Claude Code in the preview, including its model selection, skills, settings and permission prompts. Attach focused-slide context with a prompt hook, preserve the session across browser reloads, simplify the chat layout, and render Markdown in Codex responses.

### Patch Changes

- Updated dependencies [bb5a337]
- Updated dependencies [3c9850f]
- Updated dependencies [55fc76b]
  - @office-kit/pptx-preview@0.9.9

## 0.3.0

### Minor Changes

- 9be0737: Add a preview chat panel backed by local Claude Code or Codex, with focused-slide context, automatic TSX edit previews, conversation recovery and cancellation.

## 0.2.4

### Patch Changes

- c47afdd: feat: slide text in the preview can be selected and copied

  The viewer rendered each slide inside a sandboxed `<iframe>` with pointer events disabled, so nothing on a slide could be selected. The slide now lives in a shadow root on the stage: the deck's SVG stays isolated from the viewer's DOM and CSS, but its text (XHTML inside `<foreignObject>`) is ordinary selectable content, and arrow-key navigation keeps working after clicking into a slide. In presentation mode a click still advances the deck unless it made a selection.

## 0.2.3

### Patch Changes

- Updated dependencies [7e89e62]
- Updated dependencies [36c4987]
- Updated dependencies [36c4987]
- Updated dependencies [36c4987]
- Updated dependencies [64fa3c7]
- Updated dependencies [7e89e62]
  - @office-kit/pptx@0.18.0
  - @office-kit/pptx-dsl@0.5.0

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
