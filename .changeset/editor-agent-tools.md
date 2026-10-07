---
'@office-kit/pptx-editor': minor
---

A language model can now edit the embedded editor's presentation without per-function glue:

- `tools()` returns every mutating `@office-kit/pptx` export as a tool definition (`{ name, description, input_schema }`, the shape the Anthropic Messages API takes; OpenAI takes `input_schema` as `parameters`). Names are the export names, descriptions their TSDoc, and input schemas JSON Schema (draft 2020-12) generated from their TypeScript signatures, with slides, shapes, cells and layouts passed as refs. `listSlides`, `listShapes`, `listLayouts` and `listComments` let the model find those refs. The definitions load on first use, so embeddings that do not call them do not download them.
- `run(name, input)` checks the model's input against the tool's schema without coercing it, rejecting with every problem by its JSON Pointer for the model to correct, then runs the tool as one undo step named "Agent: name" and resolves with its result as JSON.
