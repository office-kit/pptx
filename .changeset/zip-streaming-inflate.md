---
'@office-kit/pptx': patch
---

Fixed: `loadPresentation` kept decompressing a ZIP entry to the end even after its output passed the size the archive's central directory declared for it, so a small `.pptx` whose entries understate their size could block the thread for seconds or minutes. Entries are now inflated in 64 KB steps and the load throws as soon as one outgrows its declared size, so decompression work is bounded by the declared sizes, which a caller can check up front from the central directory. A ZIP64 end-of-central-directory locator is still followed, and the sizes of the directory it points at are enforced the same way. The load also now throws on archives that are malformed in ways that were previously read silently: an entry that inflates to less than its declared size, a stored entry whose two sizes differ, two entries with the same name, and offsets or sizes that point outside the file.
