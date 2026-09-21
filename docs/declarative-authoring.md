# Declarative PowerPoint authoring

## End state

Claude Code can author typed TSX, preview the generated presentation locally,
revise the source with live preview feedback, and export an editable PPTX.
The DSL is a separate package; preview remains a separate package in this repo.
No React or Vue runtime is required. Application-specific UI components are
outside the initial scope. Raw is an escape hatch, not a substitute for standard
OOXML support.

## Completion requirements

- Typed JSX runtime, composition, data-driven iteration, meaningful diagnostics.
- Native text, rich paragraphs, shapes, images, tables, charts and their styles.
- Source presentations, layout selection, slide reuse and explicit editing.
- Preserve untouched source information, including unknown extension parts.
- Dependency-preserving slide copying/importing; no silent feature removal.
- Master/layout/theme access and authoring without flattening inheritance.
- Typed Raw contexts using public core APIs.
- CLI project initialization, TSX build/export, watch preview and error recovery.
- VSCode editing/preview integration and template reference discovery.
- Claude Code authoring instructions and runnable examples.
- Source-to-output diagnostics, representative integration tests and visual QA.
- Feature coverage against OOXML/PptxGenJS recorded honestly; Raw-only support
  must not be counted as typed declarative coverage.

## Architecture

`pptx-dsl` produces presentations through `pptx`. `pptx-preview` renders those
presentations independently of the DSL. Development tooling composes these
packages and owns file watching and the local preview server. Node-only code
must not enter the browser-safe DSL runtime.

Positions and dimensions in the DSL are inches; text sizes are points. Existing
source styles are retained unless explicitly overridden. Raw runs after normal
construction, before validation/export. Source inputs are loaded fresh for each
compilation so watch rebuilds cannot accumulate mutations.

## Implementation status

The first authoring workflow is implemented in `packages/dsl` and `packages/dev`.
It can initialize a typed TSX project, compile and export native editable objects,
inspect template references, and update a local preview after source changes.
VSCode uses the generated TypeScript configuration, tasks and Simple Browser.
The generated `CLAUDE.md` documents the edit/preview/export loop.

| Capability                                                     | Status                                                                                   |
| -------------------------------------------------------------- | ---------------------------------------------------------------------------------------- |
| Typed TSX, functions, arrays and fragments                     | Implemented without React/Vue                                                            |
| Text, rich paragraphs, shapes, images, charts and basic tables | Typed native elements; see package README for properties                                 |
| Existing decks, shape targets and layout selection             | Implemented; source decks default to preserving the original sequence                    |
| Unknown source parts and untouched slide bodies                | Retained in round-trip tests; no reconstruction through DSL types                        |
| Same-deck slide reuse                                          | Owned dependency copying includes charts, workbooks, notes and unknown parts             |
| Typed Raw                                                      | Presentation, slide and shape scopes; deferred until construction ends                   |
| Build/watch/export and template inspection                     | CLI and integration tests, including recovery after errors                               |
| Preview                                                        | Separate package; renders serialized output; visually checked on the three-slide example |
| Diagnostics                                                    | TypeScript in VSCode/check command, TSX element source locations for evaluation errors   |
| Cross-deck imports                                             | Not exposed in DSL; core import still needs dependency preservation                      |
| Master/layout creation and targeted theme authoring            | Remain to implement; existing template hierarchy is retained                             |
| Full OOXML/PptxGenJS expression coverage                       | Not complete; Raw-only features are not counted as typed support                         |
| Preview-to-source navigation                                   | Not implemented                                                                          |

These remaining design requirements are deliberately recorded separately from
the working authoring loop. High-level presentation UI components are deferred.
