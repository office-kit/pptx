---
'@office-kit/pptx': patch
---

`removeSlide` now drops slide relationships in other slides that pointed at the removed slide, along with the `<a:hlinkClick>` / `<a:hlinkHover>` elements that carried them. A slide-jump click action stores its relationship on the *referring* slide, so removing the target used to leave a dangling relationship: PowerPoint rejects the package and a later `duplicateSlide` of the referring slide throws `Cannot duplicate missing dependency`.
