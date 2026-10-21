---
'@office-kit/pptx': patch
---

Fixed: a small `.pptx` whose ZIP entries understate their decompressed size could block `loadPresentation` for seconds or minutes. The load now throws as soon as an entry outgrows its declared size, so decompression work is bounded by the uncompressed sizes recorded in the archive's central directory (plus at most about 17 MB per load).

`loadPresentation` also now throws on malformed archives it used to read silently: an entry whose data is shorter than declared, an entry declaring more data than its compressed bytes could possibly hold, a stored entry whose two sizes differ, two entries with the same name, and offsets or sizes that point outside the file. Errors for corrupt compressed data now start with `invalid ZIP:` and name the entry.
