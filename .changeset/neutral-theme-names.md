---
'@office-kit/pptx': patch
'@office-kit/pptx-editor': patch
'@office-kit/pptx-dev': patch
---

Built-in theme names are now neutral.

- `createPresentation()` writes its theme as "Default Theme", with color, font and format schemes named "Default". Decks created earlier keep the names they were saved with.
- The editor's Design tab names the first three themes Standard Theme, Classic Theme and Legacy Theme (標準テーマ, クラシック テーマ, レガシー テーマ), and the matching color sets and font pairs Standard, Classic and Legacy (標準, クラシック, レガシー). Picking one writes the new name into the deck. When an opened deck's color scheme uses the name another presentation app gives one of these sets, Theme Colors shows the editor's name for it.
