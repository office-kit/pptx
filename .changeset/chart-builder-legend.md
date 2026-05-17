---
'pptx-kit': minor
---

feat: chart builder writes back `<c:legend>` and `<c:dispBlanksAs>`.
The chart-root previously emitted only the default legend / blanks
behavior; the builder now:

- emits `<c:legend>` with `legendPos`, `overlay`, and one
  `<c:legendEntry><c:idx><c:delete val="1"/></c:legendEntry>` per
  hidden series index — or skips the element when
  `spec.legend.position === null` (author wants no legend)
- threads `spec.dispBlanksAs` (`'gap' | 'zero' | 'span'`) into
  `<c:dispBlanksAs>`

Round-trip test added.
