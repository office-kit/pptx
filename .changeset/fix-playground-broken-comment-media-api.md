---
'pptx-kit': patch
---

fix(site): playground stopped rendering after `SlideCommentData` was made
opaque and `getSlideMediaPartNames` lost its `(pres, slide)` two-arg
form. The playground was still doing `comment.text` and
`getSlideMediaPartNames(pres, slide)`, both of which threw at runtime.
Switched to the public `getCommentText(comment)` accessor and the
single-arg `getSlideMediaPartNames(slide)` signature.
