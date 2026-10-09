# @office-kit/pptx-editor

## 0.1.0

### Minor Changes

- 1c7d8ee: Agents in the browser can now read and edit the embedded editor's presentation through the `mountEditor` handle:

  - `selection()` returns the shapes the user selected as `ShapeRef`s (`{ slideIndex, slide, shapeId, name }`), and `resolveShape(presentation, ref)` finds such a shape again, throwing once it has been deleted.
  - `apply(label, edit)` runs `edit(presentation)` with the `@office-kit/pptx` API as one undo step named "Agent: label" (「エージェント: label」 in Japanese). An edit that throws is rolled back completely and `apply` rejects with its error.
  - `on('change', listener)` reports every kept edit with its source (`'user'` or `'agent'`), including Undo and Redo; `on('selectionchange', listener)` reports the user's selection.

- 52ff93e: A language model can now edit the embedded editor's presentation without per-function glue:

  - `tools()` returns every mutating `@office-kit/pptx` export as a tool definition (`{ name, description, input_schema }`, the shape the Anthropic Messages API takes; OpenAI takes `input_schema` as `parameters`). Names are the export names, descriptions their TSDoc, and input schemas JSON Schema (draft 2020-12) generated from their TypeScript signatures, with slides, shapes, cells and layouts passed as refs. `listSlides`, `listShapes`, `listLayouts` and `listComments` let the model find those refs. The definitions load on first use, so embeddings that do not call them do not download them.
  - `run(name, input)` checks the model's input against the tool's schema without coercing it, rejecting with every problem by its JSON Pointer for the model to correct, then runs the tool as one undo step named "Agent: name" and resolves with its result as JSON.

- c8512fe: Animations tab:

  - Effect Options lists the emphasis effects' own options. Spin has Direction and Amount (Quarter Spin to Two Spins), Grow/Shrink has Direction and Amount (Tiny, Smaller, Larger, Huge), and Transparency has Amount (25–100%). The colour effects show the theme and standard colour palette. The animation pane offers the same choices.
  - Exit Effects opens a gallery of tiles under the reference desktop app's group headings (Basic, Subtle, Moderate, Exciting), like Entrance and Emphasis. A narrow ribbon's Emphasis Effects button opens the same kind of gallery.
  - Picking a different effect gives it that effect's own default duration instead of keeping the old one.
  - Preview plays the spin angle, grow/shrink size and transparency amount that are set. Multi-part presets such as Bounce, Boomerang, Center Revolve, Rise Up, Float, Drop, Flip, Whip, Curve Up, Teeter, Wave, Pulse and Blink now follow each part's own timing.

- 167bdc1: The editor no longer links to this project's issue tracker. Help ▸ Feedback is gone from the menu bar of every editor mounted with `mountEditor`, so an application that embeds the editor shows no feedback entry point that leads its users outside the application.
- 88105a3: Long-running agents and hosts can hand the editor a version of the deck made elsewhere:

  - `propose(base, edited, { label, from })` merges `edited` (made from `base`) with the deck as it is now, including the user's unsaved edits. Changes to different slides, shapes, media and links combine into one undo step. When the same item changed on both sides, the title bar names each collision and the user keeps their edits or takes the proposed version (one undo step, so nothing is lost); the promise resolves with `{ status: 'applied' | 'kept-mine' | 'took-theirs' }` and the conflicts. `from: 'source'` is for a newer version of the file itself: the prompt speaks of the source, `'change'` reports `'source'`, and a result equal to what the host saved is not an unsaved change.
  - `@office-kit/pptx-editor/merge` exports the same three-way merge (`mergeDecks`, `describeConflict`) for Node and the browser.
  - New options: `fileName`, `autoSave` (the AutoSave switch, saving through `onSave` after each edit), `compact` (a slimmer title bar), `status` (your own element in the title bar) and `isolate: false` (render without a shadow root, for a page that is the editor's alone). `onSave` may resolve `false` to leave the deck unsaved without a message.
  - New on the handle: `open(pptx, { fileName, unsaved })`, `save()`, `dirty`, `locale`, and the `'dirtychange'` and `'localechange'` events.

- c535941: feat: copy slides and objects to the system clipboard. Copying slides or objects in the editor now puts them on the clipboard for other presentation apps and documents: slides and drawn objects paste there as pictures at their size on the slide, and text boxes and tables copied on their own paste as text and tables. Another editor — in another tab or browser — pastes them back as editable slides and objects, with their pictures, layouts and themes. Pictures and text copied in other apps paste as a picture or a text box.

  feat: the slide thumbnail menu has **Download Selected Slides...**, which saves just the selected slides as a `.pptx` to import into another presentation app.

  feat: `renderSlideToSvg(pres, slide, { background: false })` draws only the slide's own shapes on a transparent surface, without the background or master and layout graphics.

- d92bbd7: Outline View now edits across slide boundaries as the reference desktop app's (Mac) outline does:

  - **Select across slides** by dragging with the mouse or Shift-clicking, as well as with Shift+arrow keys.
  - **Delete, Cut, typing, pasting and Enter** over a selection that crosses a slide title remove the slides whose titles are selected: the remaining text after the selection joins the paragraph where it starts, and the last slide's remaining body moves up. This now works from any title or body paragraph, not only from a title. Slides with other objects ask for confirmation first.
  - **Backspace** at the start of a slide title (or **Delete** at the end of the text before it) merges that slide into the previous one.
  - **Drag a bullet** to move the paragraph and its sub-points to any position, including another slide; dragging sideways changes their level, and dragging a top-level bullet left turns it into a new slide.
  - **Drag a slide icon to the right** to demote the slide into the previous slide's body.
  - **Drag selected text** to move it, or hold Option (Control elsewhere) to copy it, including into another slide.

  Each of these is one Undo step.

- fc1acfa: New package: embed the desktop-app-style presentation editor in your own web application.

  - `mountEditor(target, { source, locale, onSave })` opens a `.pptx` (or a new presentation) in `target` and returns a handle with `ready`, `snapshot()` and `destroy()`. `onSave` receives the `.pptx` bytes when the user saves; without it, saving downloads the file.
  - The editor renders in a shadow root: the page's styles do not reach it and its styles do not reach the page. Its keyboard shortcuts leave the page's own inputs alone, and several editors can share a page.
  - The package ships compiled JavaScript with the Svelte runtime bundled, so the host application needs no framework.

### Patch Changes

- 5e6bbc5: Built-in theme names are now neutral.

  - `createPresentation()` writes its theme as "Default Theme", with color, font and format schemes named "Default". Decks created earlier keep the names they were saved with.
  - The editor's Design tab names the first three themes Standard Theme, Classic Theme and Legacy Theme (標準テーマ, クラシック テーマ, レガシー テーマ), and the matching color sets and font pairs Standard, Classic and Legacy (標準, クラシック, レガシー). Picking one writes the new name into the deck. When an opened deck's color scheme uses the name another presentation app gives one of these sets, Theme Colors shows the editor's name for it.

- 74e4e7d: The editor and the dev tool no longer show third-party product names or our own branding in their UI.

  - **Editor.** The title bar starts with the File menu; the "◈ @office-kit/pptx Editor" mark is gone, so an embedded editor carries no brand. The Help menu's product-named help item is now "Editor Help" (エディター ヘルプ), the Tools menu's product-named add-ins item is "Add-ins..." (アドイン), and Share ▸ Send a Copy offers "PPTX Presentation". Tooltips for unavailable features describe what is missing (for example "Translation needs an online translation service.", "Macros (VBA) do not run in this editor.", "Soft edges are not available for text.") instead of naming a third-party product or service.
  - **Dev tool.** The preview, presenter and editor pages are titled "Presentation preview", "Presenter view" and "Presentation editor", and the scaffolded project instructions and agent prompts say "presentation" rather than naming a product.

- 3f4aefe: The READMEs and npm package descriptions describe compatibility in terms of presentation apps in general instead of naming a third-party product. The root README gains a Trademarks section. No code behaviour changes.
- 9d0754c: API documentation (the TSDoc shipped in the type declarations and in the editor's tool descriptions) no longer uses third-party product names. Behaviour that was checked against a specific desktop presentation app is now described as "the reference desktop app". The `@office-kit/pptx-dsl` package description and the `@office-kit/pptx-preview` npm keywords were reworded the same way. No API or output changes.
- af5b2b4: The chart API documentation (the TSDoc shipped in the type declarations and in the editor's tool descriptions) now says "spreadsheet" number formats and date system instead of naming a spreadsheet product. No API or output changes. The embedded chart workbook keeps its part name, `/ppt/embeddings/Microsoft_Excel_Worksheet{N}.xlsx`.
- ec775df: The editor's styled-text gallery is now called "Text Art" (テキスト アート) instead of using a third-party feature name: Insert ▸ Text Art, Shape Format and Table Design ▸ Text Art Styles (テキスト アートのスタイル), Text Art Quick Styles (テキスト アートのクイック スタイル) and Clear Text Art (テキスト アートのクリア). The presets and what they write are unchanged.
- ec775df: `TextFormat.bold` and `TextFormat.spc` accept `null`, which removes the `b` / `spc` attribute so the run inherits its weight and character spacing again. `false` and `0` still write `b="0"` / `spc="0"`, which override an inherited bold or spacing.

  The editor's Text Art Quick Styles now remove bold and character spacing that a preset does not set, as the reference desktop app does, instead of writing `b="0"` and `spc="0"`. Text whose list style or placeholder makes it bold or spaced keeps that look after applying such a preset.

- 9cad210: Fixed: Video Format ▸ Poster Frame ▸ Reset could occasionally set a blank, fully transparent poster instead of the video's first frame (seen in Chromium on Linux). Reset now waits until the first frame has actually painted before saving it as the poster.
- Updated dependencies [c8512fe]
- Updated dependencies [4410413]
- Updated dependencies [d8e455d]
- Updated dependencies [0d26527]
- Updated dependencies [c535941]
- Updated dependencies [550b55c]
- Updated dependencies [afe80eb]
- Updated dependencies [5e6bbc5]
- Updated dependencies [3f4aefe]
- Updated dependencies [946dac4]
- Updated dependencies [9d0754c]
- Updated dependencies [af5b2b4]
- Updated dependencies [1d447f8]
- Updated dependencies [acb95fc]
- Updated dependencies [21d59f7]
- Updated dependencies [1d447f8]
- Updated dependencies [2912bdc]
- Updated dependencies [ec775df]
- Updated dependencies [6ac3ead]
- Updated dependencies [197b737]
- Updated dependencies [24a6ae0]
  - @office-kit/pptx@0.24.0
  - @office-kit/pptx-preview@0.15.0
