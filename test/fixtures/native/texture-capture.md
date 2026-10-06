# Mac PowerPoint default texture captures

Captured on 2026-10-07 from Mac desktop PowerPoint, driven through System
Events UI scripting. The deck was created by PowerPoint itself (blank layout,
slide 1 with one plain rectangle, slide 2 empty) and saved in the app
container. Microsoft's texture bitmaps are not stored here; only XML and
metadata are kept.

| Fixture | Action |
| --- | --- |
| `texture-default-shape.xml` | Format Shape ▸ Fill ▸ Picture or texture fill on the rectangle |
| `texture-default-background.xml` | Format Background ▸ Fill ▸ Picture or texture fill on slide 2 |

Observations:

- Both actions immediately insert a tiled picture fill; no file chooser opens.
  The shape pane title changes from Format Shape to Format Picture and gains a
  Picture category. Format Background keeps its title and gains Effects and
  Picture categories.
- The inserted texture is **Papyrus**, the first gallery entry. Choosing Papyrus
  from the Texture gallery afterwards saved a byte-identical `image1.png`
  (SHA-1 `7dd4b009…`); choosing Canvas saved different bytes. With a fresh
  PowerPoint process and the background acted on first, the background also
  received Papyrus, and a following shape reused the same media part.
- Saved XML for both targets is
  `<a:blipFill><a:blip r:embed="…"/><a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/></a:blipFill>`,
  with no `rotWithShape` and no `dpi` attribute. Rotate with shape shows checked
  for the shape and checked but disabled for the background. The background
  also writes an empty `a:effectLst`.
- The media is a 128 × 128 RGB PNG with `sRGB`, `eXIf` and `pHYs` chunks. The
  density differed between PowerPoint processes: 5669 px/m (≈144 DPI, 64 pt
  tiles) in the first session and 2835 px/m (≈72 DPI, 128 pt tiles) after a
  relaunch, for both the default and an explicit Papyrus pick. The pixel data
  was identical; only `pHYs` and the EXIF resolution differed.
- The Texture ▾ gallery marks no swatch as selected, either after the default
  insertion or after picking Papyrus explicitly (every swatch reports
  `AXValue 0` and none is drawn highlighted). It lays out the 24 textures five
  per row as 49 × 49 pt radio buttons on a 47 pt pitch inside a 235 × 235 pt
  group; each visible tile is about 36 pt with a thin grey border. The Format
  pane's gallery has no More Textures... item.
- Within one session PowerPoint can reuse a recently applied texture instead of
  Papyrus: after Canvas had been picked on a shape, Picture or texture fill on a
  background inserted Canvas, and a newly drawn oval also received Canvas even
  after Denim had been picked for that background. The exact rule was not
  isolated.
