# Mac PowerPoint WordArt captures

Captured on 2026-10-03 from the Shape Format ribbon in Mac desktop PowerPoint.
The selected object was the title placeholder in
`/tmp/pptx-outline-audit/reference.pptx`, containing `Outline title`.
Each style was applied independently, saved, copied, undone, and saved again.
The final slide XML was byte-identical to the pre-capture slide XML.

| Fixture | Native gallery label |
| --- | --- |
| `wordart-black-shadow-shape.xml` | Fill: Black, Text color 1; Shadow |
| `wordart-accent1-shadow-shape.xml` | Fill: Blue, Accent color 1; Shadow |
| `wordart-accent2-outline-shape.xml` | Fill: Red, Accent color 2; Outline: Red, Accent color 2 |

These are native shape XML extracts, with namespace declarations added to make
each extract self-contained. They are evidence for implementing the WordArt
style gallery, not evidence that the editor already supports that gallery.

Observed payloads:

- Black shadow: no explicit text fill in this title; zero-width outline;
  outer shadow blur 38100 EMU, distance 19050 EMU, direction 45 degrees,
  alignment `tl`, no rotation with shape, `dk1` at 40% opacity.
- Accent 1 shadow: explicit `accent1` fill; zero-width outline; outer shadow
  blur 38100 EMU, distance 25400 EMU, direction 90 degrees, alignment `ctr`,
  no rotation with shape, RGB `6E747A` at 43% opacity.
- Accent 2 outline: sets bold, outline width 22225 EMU with solid accent 2
  line, and accent 2 text fill with `lumMod=40000`, `lumOff=60000`.
- All three apply the character properties to both the run and end-paragraph
  properties. Shape geometry and body properties remain unchanged.

Before implementing all presets, capture the remaining gallery entries and
verify how these presets replace pre-existing fill and effect properties.
The black preset's omitted fill on this inherited black title does not prove
that it should preserve an arbitrary prior text color.
