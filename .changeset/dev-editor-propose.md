---
'@office-kit/pptx-dev': patch
---

TSX rebuilds now reach the open editor as one undo step ("Source changed") merged with unsaved canvas edits, instead of reloading the deck and starting a new Undo history. A collision is offered as before (**Keep my edits** / **Use source**), and **Use source** can be undone in the editor.
