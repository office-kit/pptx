---
"@office-kit/pptx": minor
---

Comments written by PowerPoint 2021 and Microsoft 365 are now read and edited,
not merely carried along. Those decks keep their comments in a different format
from the one ECMA-376 defines ([MS-PPTX] §2.16.1): a thread owns its replies
instead of the replies pointing back at a parent, the text is a DrawingML body
rather than a string, authors are identified by GUID in their own
`/ppt/authors.xml`, and — the part that has no equivalent at all in ECMA-376 — a
thread can be **resolved**.

`getSlideComments` returns both kinds. `getCommentFormat` says which one a
comment came from, `getCommentStatus` reports `'active'`, `'resolved'` or
`'closed'`, and `setCommentStatus` resolves a thread or reopens it. Asking for
a status on an ECMA-376 comment gives `null`, and setting one throws rather than
quietly doing nothing: that file has nowhere to keep it.

`setCommentText`, `removeSlideComment` and `clearSlideComments` work on either
format. `addSlideComment` follows the file rather than the caller — a reply
joins its own thread, and a new thread goes where the slide already keeps its
comments, so a deck that has never had one still starts an ECMA-376 list that
every reader understands. Everything the library does not model is left exactly
where it was: the slide or shape anchor, extension lists, reactions, and the
pin.

Fixes: a comment added without a pin produced invalid XML. `<p:pos>` is
required by `CT_Comment`, and nothing in the API made a caller pass one — so
every comment the editor added was rejected by the schema. Such a comment is now
pinned to the slide's origin, and `getCommentPosition` reports that rather than
`null`. Comments read from a file that omits the element still report no
position, and are still saved the way they arrived.
