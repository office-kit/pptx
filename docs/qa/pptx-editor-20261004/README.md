# WordArt font color — editor screenshots (2026-10-04)

Real Chromium captures of the development editor choosing a font color for text
that carries a WordArt pattern fill (`textFill`), on the main canvas and in the
speaker notes. All four images come from one run against one deck.

| File                | Step                                                                                      |
| ------------------- | ----------------------------------------------------------------------------------------- |
| `canvas-before.png` | Deck loaded; text box "Pattern" has the pattern fill.                                     |
| `canvas-after.png`  | Caret after "Pattern", Text color → Red, typed "X"; the reopened menu shows Red selected. |
| `notes-before.png`  | After the canvas edit was committed; Notes pane opened on notes "Pattern" (pattern fill). |
| `notes-after.png`   | End → Bold → Text color → Red → typed "X" → Tab.                                          |

## Data

Built with the public API (`createPresentation`, `addBlankSlide`,
`addSlideTextBox`, `setShapeTextFormat`, `setSlideNotes`, `setSlideNotesFormat`)
and opened through `<Presentation source={bytes} />`:

- Slide 1: text box at (1 in, 1 in), 8 × 1.5 in, text `Pattern`, 60 pt bold,
  `textFill: { kind: 'pattern', preset: 'dkUpDiag', foreground: 'accent1', background: '#FFFFFF' }`.
- Speaker notes: `Pattern` with the same `textFill`.

## Environment

- Viewport 1400 × 900, Playwright 1.63.0 with headless Chromium (cached
  headless shell 1243), Node 24.16.0, `node packages/dev/dist/cli.mjs dev` on a
  temporary deck.
- Source: HEAD `01461812` plus the uncommitted changes in this PR's tree,
  rebuilt at 2026-10-04 11:13 UTC. Bundle sha256: `packages/dev/dist/editor.js`
  `de1aecbb128c2c72…`, `packages/dev/dist/cli.mjs` `57bd61eae224999d…`,
  `dist/index.js` `4ea663478d167651…`.
- Captured 2026-10-04 11:17 UTC. Page errors: none. Console errors: none.

## What the run checked (read back from the saved deck)

- Canvas: runs are `Pattern` (pattern fill kept) and `X` (`color: '#FF0000'`,
  no `textFill`).
- Notes: text `PatternX`; the inserted `X` is bold, `#FF0000`, no `textFill`;
  the original text keeps the pattern fill.

## Limits

- The editor does not paint gradient or pattern text fills yet
  (`packages/dev/NATIVE_PARITY.md`, "Font color replacement on WordArt").
  The pattern-filled "Pattern" therefore shows as plain black in every image;
  the pattern is proven only by the saved-deck checks above.
- Not compared with the reference desktop app, Keynote or another native app.

## sha256

```
e54ba4ebf3bb5c5e12eab0f9753e33faaa616f4d4dda55a981411e9a977599b3  canvas-before.png
20307964895c558fb6409201aa87542b1269f4f2a1aeaa93eb4ae1b565b6b3d4  canvas-after.png
09d67df684e0c746c5b141fd0f043ab8915ad93173069eee031cfe1bfb316a0b  notes-before.png
1fc6b593ecb66217752900f70665760e388b4909db2e1657abdb9652aee0b738  notes-after.png
```
