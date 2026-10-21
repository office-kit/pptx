---
'@office-kit/pptx': patch
---

Fixed: a small `.pptx` whose ZIP entries understate their decompressed size could block `loadPresentation` for seconds or minutes. The load now throws as soon as an entry outgrows its declared size, so decompression work is bounded by the declared sizes (plus at most about 17 MB of overrun), which a caller can check up front.

`loadPresentation` also now throws on malformed archives it used to read silently: an entry whose data is shorter than declared, a stored entry whose two sizes differ, two entries with the same name, and offsets or sizes that point outside the file. Errors for corrupt compressed data now start with `invalid ZIP:` and name the entry.
