---
'@office-kit/pptx-dev': minor
---

The editor's title bar now carries Mac PowerPoint's menu bar — File, Edit, View, Insert, Format, Arrange, Tools, Slide Show, Window and Help — in place of the Edit and View buttons:

- **Menus match PowerPoint** in English and Japanese: order, separators, submenus and shortcut glyphs, with items enabled and checked as PowerPoint's are with nothing, a shape or text selected (Undo names the last edit; Apply To Defaults becomes Apply Object Style when something is selected). Items run the editor's existing commands; what the editor cannot do yet is disabled with the reason as a tooltip. While text is being edited, the menus leave the caret and selection in the text.
- **Keyboard shortcuts follow PowerPoint's**, from the same table the menus show. Changed keys: ⌘G is Find Next, so Group is ⌥⌘G and Ungroup ⌥⇧⌘G; ⌘K inserts a hyperlink and the command search moved to ⌘? (Help ▸ PowerPoint Help); Pick Up / Apply Object Style are ⇧⌘C / ⇧⌘V (⌥⌘C to pick up in Japanese) instead of ⌥⌘C / ⌥⌘V; Paste and Match Formatting is ⌥⇧⌘V (was ⇧⌘V); Replace is ⌃H (was ⌘H). While editing slide text, ⌘L/⌘E/⌘R align it, and ⌘T, ⌘K, ⌥⌘M and the object-style keys act on the selected text. Control still stands in for Command where PowerPoint has no Control shortcut of its own.
- The command search now closes with Escape.
