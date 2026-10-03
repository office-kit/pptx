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
| `wordart-white-accent5-shadow-shape.xml` | Fill: White; Outline: Aqua, Accent color 5; Shadow |
| `wordart-gray-gradient-shape.xml` | Gradient Fill, Gray |
| `wordart-accent4-soft-bevel-shape.xml` | Fill: Purple, Accent color 4; Soft Bevel |
| `wordart-accent5-gradient-reflection-shape.xml` | Gradient Fill: Aqua, Accent color 5; Reflection |
| `wordart-accent4-gradient-outline-shape.xml` | Gradient Fill: Purple, Accent color 4; Outline: Purple, Accent color 4 |
| `wordart-white-accent1-glow-shape.xml` | Fill: White; Outline: Blue, Accent color 1; Glow: Blue, Accent color 1 |
| `wordart-accent3-sharp-bevel-shape.xml` | Fill: Olive Green, Accent color 3; Sharp Bevel |
| `wordart-black-white-hard-shadow-shape.xml` | Fill: Black, Text color 1; Outline: White, Background color 1; Hard Shadow: White, Background color 1 |
| `wordart-black-accent5-hard-shadow-shape.xml` | Fill: Black, Text color 1; Outline: White, Background color 1; Hard Shadow: Aqua, Accent color 5 |
| `wordart-accent5-hard-shadow-shape.xml` | Fill: Aqua, Accent color 5; Outline: White, Background color 1; Hard Shadow: Aqua, Accent color 5 |
| `wordart-white-accent2-hard-shadow-shape.xml` | Fill: White; Outline: Red, Accent color 2; Hard Shadow: Red, Accent color 2 |
| `wordart-background2-inner-shadow-shape.xml` | Fill: Tan, Background color 2; Inner Shadow |
| `wordart-white-pattern-shadow-shape.xml` | Pattern Fill: White; Dark Upward Diagonal Stripe; Shadow |
| `wordart-accent3-pattern-inner-shadow-shape.xml` | Pattern Fill: Olive Green, Accent color 3, Narrow Horizontal Stripe; Inner Shadow |
| `wordart-accent1-pattern-hard-shadow-shape.xml` | Pattern Fill: Blue, Accent color 1, 50%; Hard Shadow: Blue, Accent color 1 |
| `wordart-accent5-pattern-outline-shape.xml` | Pattern Fill: Aqua, Accent color 5, Light Downward Diagonal Stripe; Outline: Aqua, Accent color 5 |
| `wordart-dark-blue-pattern-hard-shadow-shape.xml` | Pattern Fill: Dark Blue, Dark Upward Diagonal Stripe; Hard Shadow |

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
- White/accent 5 shadow: sets bold, white RGB fill, accent 5 outline width
  10160 EMU; black shadow at 30% opacity, blur 38100 EMU, distance 22860 EMU,
  direction 90 degrees, alignment `tl`, no rotation with shape.
- Gray gradient: stops at 21% (`53575C`) and 88% (`C5C7CA`), linear angle
  90 degrees, zero-width outline, and no explicit effects.
- These first five presets apply the character properties to both the run and end-paragraph
  properties. Shape geometry and body properties remain unchanged.

All 20 entries in the native gallery now have captures. Additional observations:

- The two bevel presets write `scene3d` and `sp3d` under `bodyPr`. Shape-level
  `spPr` stays empty; applying these through a shape-effects API would target
  the wrong object.
- The aqua reflection uses start opacity 53%, end opacity 0.3%, end position
  35.5%, vertical scale -90%, blur 6350 EMU, bottom-left alignment, and no
  rotation with the shape.
- Pattern presets use `pattFill` on character properties. Inner shadows use
  `innerShdw`, not the outer-shadow element. Neither can be represented by a
  plain character color plus an outer shadow.
- The white/blue glow preset writes `spc="50"` and the unexpected color
  `srgbClr=70AD47` with `tint=1000`, despite the gallery label saying white.
  Preserve this evidence; verify in another native deck/theme before choosing
  general preset constants.
- Applying the white/accent5 shadow preset and then the black shadow preset
  removes the earlier explicit white fill and bold flag. The resulting title
  properties match the independently captured black preset. Presets replace
  these visual properties rather than merging them. This was followed by two
  Undo operations and Save; saved slide XML matches the pre-comparison copy.

Still verify how presets treat manually applied fonts, font size, italics,
paragraph formatting, selected text ranges, inherited non-black text colors,
and existing 3D properties. A whole-format reset may clear more than native
WordArt does. Captures alone do not establish gallery or rendering parity.
