---
'@office-kit/pptx': patch
'@office-kit/pptx-preview': patch
'@office-kit/pptx-dsl': patch
'@office-kit/pptx-editor': patch
---

API documentation (the TSDoc shipped in the type declarations and in the editor's tool descriptions) no longer uses third-party product names. Behaviour that was checked against a specific desktop presentation app is now described as "the reference desktop app". The `@office-kit/pptx-dsl` package description and the `@office-kit/pptx-preview` npm keywords were reworded the same way. No API or output changes.
