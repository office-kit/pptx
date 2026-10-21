---
'@office-kit/pptx-editor': minor
---

New entry `@office-kit/pptx-editor/headless` lets an agent revise a presentation without mounting the editor, in Node as well as the browser. `agentEditSchema` describes nine edits as one tool (`setShapeText`, `setTableCellText`, `setChartData`, `setShapeTextFormat`, `setShapeFill`, `setShapeBounds`, `removeShape`, `moveSlide`, `removeSlide`) in one compact JSON Schema, with text edits always keeping the formatting of unchanged text. `readForAgent` returns each slide's title, notes and shapes, `decodeAgentEdits` checks a batch and names every problem by its JSON Pointer from the tool input (`/edits/0/…`), and `applyAgentEdits` applies it, reporting which edit failed as an `AgentEditError`.

The `listShapes` tool now also reports a chart's kind, title, categories and series values, and `tools()` includes `setChartData`.
