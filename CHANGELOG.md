# pptx-kit

## 1.0.0

### Major Changes

- f47b78b: **1.0.0** — first stable release. The public API is now frozen under SemVer.

  **What works at 1.0:**

  - **Read** any `.pptx` produced by PowerPoint, Keynote, Google Slides, or
    LibreOffice Impress, and save it back without corruption. Unknown
    extensions are preserved verbatim on round-trip.
  - **Template editing**: token / text replace across slides and speaker
    notes, image swap with geometry preserved, slide CRUD with placeholder
    inheritance from layout / master.
  - **Authoring on top of an existing master**: 180+ preset shapes, custom
    text formatting, tables, embedded charts (column / line / bar / pie /
    doughnut / area) with auto-generated xlsx, solid / gradient / pattern /
    image fills, shadows and glows, rotation / flip / z-order, hyperlinks
    and click actions, notes and comments, slide transitions, simple
    entrance / exit animations.
  - **Diagnostics**: `validatePresentation` returns invariant violations;
    every XML part is validated against the ECMA-376 XSDs in CI.
  - **Bundling**: one ESM build runs in both Node ≥ 20 and modern browsers.
    Tree-shaking is enforced by a CI test — minimal `load → save` bundle
    is < 75 KB unminified, full fn-API bundle is ~120 KB.

  **Deferred to post-1.0** (read pass-through preserved on round-trip):

  - Constructing new themes / masters / layouts from scratch.
  - SmartArt authoring.
  - Complex animation timing-tree authoring.
  - OLE / ActiveX authoring.
  - Document encryption (read + write).

  **Performance (M-series Node 20):** 100-slide synthetic deck saves in
  ~25 ms, loads in ~20 ms. 100 MB templates fit comfortably under the 2 s
  load/save targets.

  **Migration:** if you were on the pre-1.0 class API
  (`Presentation` / `Slide` / `SlideShape` / `SlideLayout`), see the
  preceding changeset for the rename table. There is no class API at 1.0.

- 665c979: **BREAKING**: the class-based API (`Presentation`, `Slide`, `SlideShape`,
  `SlideLayout`) has been removed. Use the free-function API for every
  capability — one canonical path per operation.

  | Was                              | Now                                      |
  | -------------------------------- | ---------------------------------------- |
  | `Presentation.load(bytes)`       | `loadPresentation(bytes)`                |
  | `Presentation.create()`          | `createPresentation()`                   |
  | `pres.save()`                    | `savePresentation(pres)`                 |
  | `pres.slides`                    | `getSlides(pres)`                        |
  | `pres.slideLayouts`              | `getSlideLayouts(pres)`                  |
  | `pres.addSlide({ layout })`      | `addSlide(pres, { layout })`             |
  | `pres.removeSlide(slide)`        | `removeSlide(pres, slide)`               |
  | `pres.moveSlide(slide, i)`       | `moveSlide(pres, slide, i)`              |
  | `pres.duplicateSlide(slide)`     | `duplicateSlide(pres, slide)`            |
  | `pres.replaceTokens(map)`        | `replaceTokensInPresentation(pres, map)` |
  | `slide.shapes`                   | `getSlideShapes(slide)`                  |
  | `slide.findPlaceholder('title')` | `findSlidePlaceholder(slide, 'title')`   |
  | `slide.addTextBox(opts)`         | `addSlideTextBox(slide, opts)`           |
  | `slide.addShape(opts)`           | `addSlideShape(slide, opts)`             |
  | `slide.addImage(bytes, opts)`    | `addSlideImage(slide, bytes, opts)`      |
  | `slide.addTable(opts)`           | `addSlideTable(slide, opts)`             |
  | `slide.addLine(opts)`            | `addSlideLine(slide, opts)`              |
  | `slide.setBackground(color)`     | `setSlideBackground(slide, color)`       |
  | `slide.setTransition(opts)`      | `setSlideTransition(slide, opts)`        |
  | `slide.setNotes(text)`           | `setSlideNotes(slide, text)`             |
  | `slide.layout`                   | `getSlideLayout(slide)`                  |
  | `slide.notes`                    | `getSlideNotes(slide)`                   |
  | `slide.text`                     | `getSlideText(slide)`                    |
  | `shape.text`                     | `getShapeText(shape)`                    |
  | `shape.setText(value)`           | `setShapeText(shape, value)`             |
  | `shape.position`                 | `getShapePosition(shape)`                |
  | `shape.setPosition(x, y)`        | `setShapePosition(shape, x, y)`          |
  | `shape.setFill(color)`           | `setShapeFill(shape, color)`             |
  | `shape.setStroke(opts)`          | `setShapeStroke(shape, opts)`            |
  | `shape.setRotation(deg)`         | `setShapeRotation(shape, deg)`           |
  | `shape.setHyperlink(url)`        | `setShapeHyperlink(shape, url)`          |
  | `layout.name`                    | `getSlideLayoutName(layout)`             |

  Node entry (`pptx-kit/node`) drops the `Presentation` subclass; use
  `loadPresentationFile` / `savePresentationToFile` instead.

  **Why**: every capability used to have two paths through the public API
  — a class method and a free function. The duplication hurt
  discoverability (which one should you use?), made the bundle larger
  (class consumers dragged the whole prototype in), and forced every
  breaking change to land in two places. The free-function API is the
  canonical surface from now on.

### Minor Changes

- b41c502: Comprehensive feature surface for PPTX authoring + editing. This is the
  first release that covers every L1–L4 capability in the foundation
  plan. Highlights:

  **Round-trip + template editing (L1 / L2)**

  - `loadPresentation` / `savePresentation` (`Uint8Array` / `ArrayBuffer` / `Blob`).
  - Node convenience: `loadPresentationFile`, `savePresentationToFile`.
  - Token replace: `replaceTokensInPresentation`, `replaceTokensInSlide`.
  - Free-text replace: `replaceTextInPresentation`, `replaceTextInSlide`.
  - Slide CRUD: `addSlide`, `removeSlide`, `moveSlide`, `duplicateSlide`,
    `getSlideAt`, `getSlideIndex`, `clearSlideShapes`, `sortSlides`.
  - Cross-deck: `importSlide` (with image-media propagation).
  - Cross-slide: `copyShape`.
  - Diagnostics: `validatePresentation`, `getPresentationSummary`,
    `listPackageParts`, `readPackagePart`, `getMediaParts`,
    `setMediaPartBytes`, `compactPackage`.

  **Authoring (L3)**

  - Shapes: `addSlideTextBox`, `addSlideShape` (180+ presets),
    `addSlideLine`, `addSlideTable`, `addSlideImage`, `addSlideChart`.
  - Charts: `bar` / `column` / `line` / `pie` / `doughnut` / `area` with
    embedded xlsx; read/update via `getSlideCharts` / `setChartSpec`.
  - Tables: per-cell access (`getTableCells`, `setTableCellText`,
    `setTableCellFill`, `setTableCellTextFormat`,
    `setTableCellAlignment`); row + column insert/remove.
  - Slide layout swap: `setSlideLayout`, `findSlideLayout`.

  **Text**

  - Per-shape: `setShapeText`, `setShapeBullets`, `setShapeAlignment`,
    `setShapeTextFormat`, `setShapeHyperlink`, `setShapeTextAnchor`,
    `setShapeTextMargins`, `setShapeTextWrap`, `setShapeTextAutoFit`.
  - Per-paragraph: `setParagraphAlignment`, `setParagraphBullet`,
    `setParagraphLevel`, `setParagraphSpacing` + read-back pairs.
  - Per-run: `setShapeRunFormat`, `setShapeRunText`,
    `getShapeRunFormat`, `getShapeParagraphCount`, `getShapeRunCount`,
    `getShapeRunText`.

  **Geometry**

  - Position / size / rotation / flip + combined `setShapeBounds` /
    `getShapeBounds`. Z-order: `bringShapeToFront`, `sendShapeToBack`,
    `bringShapeForward`, `sendShapeBackward`.

  **Fill / stroke / effects**

  - Fill kinds: solid, gradient, pattern, image, none + `getShapeFill`
    read-back.
  - Stroke: color + width + dash + arrowheads + `getShapeStroke` /
    `getShapeStrokeDash` / `getShapeStrokeArrow` read-back.
  - Effects: `setShapeShadow`, `setShapeGlow`, `clearShapeEffects` +
    `getShapeEffect` read-back.

  **Pictures**

  - Crop, opacity, brightness (`lumOff`), contrast (`lumMod`),
    image replacement, image-as-fill. Read-back pairs for every setter.

  **Slide-level (L4)**

  - Notes (`getSlideNotes` / `setSlideNotes`).
  - Transitions (every effect + read-back).
  - Animations (`fadeIn` / `fadeOut` / `appear` / `disappear`) +
    read-back.
  - Comments (legacy schema, author dedup, optional position + date).
  - Backgrounds: solid color or embedded picture; read-back.
  - Visibility: `setSlideHidden` / `isSlideHidden`.
  - Slide sections (p14:sectionLst).
  - Slide size + presets (`SLIDE_SIZE_4_3` / `16_9` / `16_10`).
  - Slide title shortcut (`getSlideTitle` / `setSlideTitle`).
  - Click actions: URL / slide jump / preset nav + read-back.

  **Theme + package**

  - `getPresentationTheme` — color scheme (`accent1`–`accent6`, `dark1`,
    `light1`, `hyperlink`, ...).
  - `getMediaParts`, `listPackageParts`, `readPackagePart` for audit /
    export workflows.

  **Tree-shake**

  - The minimal `load`+`save` import is ~60 KB; the full fn-API
    bundle ~123 KB. CI guard via `test/tree-shake.test.ts`.

  All emitted XML validates against the ECMA-376 strict schemas
  (pml.xsd, dml-chart.xsd, opc-relationships.xsd, opc-contentTypes.xsd)
  via Layer-1 tests.

  **Additional helpers** (all tree-shakeable free functions)

  - Properties: `getCoreProperties` / `setCoreProperties`,
    `getExtendedProperties` / `setExtendedProperties`, plus convenience
    `getPresentationCreated`, `getPresentationModified`,
    `incrementRevision`, `touchModified`.
  - Thumbnail: `getThumbnail` / `setThumbnail` / `removeThumbnail`.
  - Theme: `getPresentationTheme`, `getPresentationFonts`.
  - Slide queries: `getSlideCount`, `getSlideLayoutCount`,
    `getVisibleSlides`, `getHiddenSlides`, `getSlidesWithNotes`,
    `getSlidesWithComments`, `getSlidesWithImages`,
    `getSlidesWithCharts`, `getSlidesWithTables`,
    `getSlidesByLayout`, `findSlideByTitle`, `findSlideByText`,
    `findSlidesByText`, `findSlideByPartName`,
    `findSlideLayoutByType`, `findSlideLayoutByPartName`.
  - Bulk inventories: `getAllNotes`, `getAllComments`, `getAllCharts`,
    `getAllTables`, `getAllImages`, `getPresentationText`,
    `getSlideOutline`.
  - Shape introspection: `getShapeAt`, `getShapeIndex`,
    `getShapeSlide`, `getShapeXmlString`, `getShapeChartKind`,
    `getShapeChartSpec`, `getShapeImageFillBytes`,
    `getShapeImageFormat`, `getShapeImagePartName`,
    `getShapeAltTitle` / `setShapeAltTitle`,
    `getShapeDescription` / `setShapeDescription`.
  - Shape predicates: `isChartShape`, `isTableShape`,
    `isShapeHidden` / `setShapeHidden`, `isShapePlaceholder`,
    `hasShapeImage`, `hasShapeText`.
  - Shape search: `findShapeByText`, `findShapesByText`,
    `findShapesByKind`, `findChartByKind`,
    `findChartsBySeriesName`, `findCommentsByAuthor`,
    `findSlidePlaceholders`, `findSlidePlaceholderByIdx`.
  - Mutation: `setShapeRunHyperlink`, `getShapeRunHyperlink`,
    `getSlideBody`, `appendShapeText`,
    `appendSlideNotes`, `removeSlideNotes`,
    `swapSlides`, `mergePresentations`, `slidesUsingMediaPart`,
    `setTableColumnWidth`, `setTableRowHeight`, `getTableColumnWidths`,
    `getTableRowHeights`, `getTableCellAlignment`, `getTableCellFill`.
  - Diagnostics: `getSlideXmlString`, `getSlidePartName`,
    `getSlideLayoutPartName`, `getSlidesByLayout`.
