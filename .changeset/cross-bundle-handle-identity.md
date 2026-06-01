---
'pptx-kit': patch
---

fix: presentation handles now interoperate across the `pptx-kit` and
`pptx-kit/node` entry points. The two entries ship as separate bundles,
and the opaque handles (`PresentationData`, `SlideData`, …) were keyed by
plain `Symbol`s minted per bundle. Loading a deck with
`loadPresentationFile` (from `pptx-kit/node`) and then reading it with,
say, `getSlides` (from `pptx-kit`) crashed with
`Cannot read properties of undefined`. The handle keys now use the global
symbol registry (`Symbol.for`), so a handle from either entry is readable
by the other — and by companion packages that bundle their own reader copy.
