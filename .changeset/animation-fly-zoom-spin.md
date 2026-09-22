---
"@office-kit/pptx": minor
---

`setShapeAnimation` and `updateSlideAnimation` take five more effects, and
`getSlideAnimations` reads them back: `flyIn` / `flyOut` travel in from, or out
through, one edge of the slide, `zoomIn` / `zoomOut` grow from nothing and
shrink back to it about the shape's centre, and `spin` turns the shape one
clockwise turn.

`spin` is the first emphasis effect. It animates a shape that is already on the
slide and leaves it exactly where it was, so unlike an entrance or an exit it
never decides whether the shape is shown — a shape no entrance has revealed
stays unrevealed while it turns.

A fly takes a `direction`: `'top'`, `'right'`, `'bottom'` (the default) or
`'left'`, naming the edge it comes from or leaves by. Steps report it as
`direction`, `null` for every effect that does not fly, and passing it for one
of those is an error rather than a silently dropped field.

An effect is only named when the tree states what it does. A preset is matched
on all three of `presetClass`, `presetID` and `presetSubtype`, so an imported
entrance that shrinks in from four times its size is no longer read as one that
grows from nothing; and a rotation is only reported as `spin` when it really is
a single full clockwise turn, so a half turn or one the other way is listed and
saved as it arrived rather than rewritten as a full one.

Fixes: an exit fade hid its shape the moment it started instead of fading it
out. The `<p:set>` that takes the shape off the slide now trails the motion the
way PowerPoint writes it, and moves when the effect's duration changes.
