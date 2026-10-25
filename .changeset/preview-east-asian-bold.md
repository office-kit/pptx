---
'@office-kit/pptx-preview': minor
---

Fixed bold Japanese, Chinese and Korean text drawn thin by `renderSlideToImage` / `renderSlideToRgba`. A run's East Asian characters have no glyphs in the bundled Latin fonts, and resvg's fallback for them picks the first loaded face that has the glyph whatever its weight, so with regular and bold CJK faces in `fontFiles` a bold heading painted with the regular face. The Latin letters and full-width punctuation (`、`, `・`, `（）`) on the same line painted with that face too. The `/node` entry now names the family of the first `fontFiles` face that covers kana for East Asian characters, so resvg picks the face closest to the run's weight, and Latin text on the line keeps its own font, as presentation apps draw a run's East Asian characters with its East Asian font.

Added `eastAsianFamily` to `RenderSlideOptions` for `textLayout: 'svg'`: the family written for East Asian characters (kana, kanji, Hangul, and CJK and full-width punctuation), measured with that family too. Unset, the SVG is unchanged. A line that mixes it with another family is written as separate text chunks, each with its own `x`, because resvg otherwise paints the whole line with one face.
