---
'@office-kit/pptx': minor
'@office-kit/pptx-editor': patch
---

`compactPackage` now removes every part that cannot be reached from the package root (`/_rels/.rels`) through internal relationships, not only `/ppt/media/` parts. This drops the charts, embedded workbooks, notes slides and tags that `removeSlide` leaves behind, so the saved file no longer carries a removed slide's chart data. Any other part nothing points at, such as custom data added without a relationship, is removed too. If the root relationships do not reach `/ppt/presentation.xml`, nothing is removed. `getOrphanMediaPartNames` follows the same rule, so media used only by an unreachable part now counts as orphaned. In the editor, copied slides and shapes no longer carry the other slides' notes, charts and workbooks.
