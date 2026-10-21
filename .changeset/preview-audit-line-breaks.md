---
'@office-kit/pptx-preview': minor
---

`auditTextLayout` can now report three more text problems. With `reportLineBreaks: true` it reports `word-break` when a wrap splits a word (Latin, Greek or Cyrillic letters and digits, or a katakana run; `text` holds the word with `|` at the break) and `short-last-line` when a wrapped paragraph ends on a line of only one or two characters (泣き別れ). With `minFontSizePt` it reports `small-font` for a shape whose text, after autofit shrinking, is smaller than that size. Both are off unless you pass them, so existing calls return the same issues; code that switches over `TextAuditIssue['kind']` exhaustively has three new kinds to handle.
