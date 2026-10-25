---
'@office-kit/pptx': minor
---

Added: `loadPresentation` and `loadPresentationFile` take an optional `limits` option for opening untrusted decks, such as uploads on a server: `{ maxEntries, maxTotalUncompressedBytes, maxXmlPartBytes, maxTotalXmlBytes }`. The limits are checked against the sizes the archive declares before anything is decompressed, and an entry that decompresses past its declared size still stops mid-way. With `limits` set, archives that use ZIP64 records are rejected too. A deck that breaks a limit throws an `Error` whose message starts with `invalid ZIP:` and names the limit. Without `limits`, loading is unchanged.
