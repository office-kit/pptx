---
'pptx-kit': minor
---

Comprehensive feature surface for PPTX authoring + editing. This is the
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
- Shape introspection: `getShapeCount`, `getTotalShapeCount`,
  `getShapeAt`, `getShapeIndex`, `getShapeSlide`,
  `getShapeXmlString`, `getShapeChartKind`, `getShapeChartSpec`,
  `getShapeImageFillBytes`, `getShapeImageFormat`,
  `getShapeImagePartName`, `getShapeAltTitle` /
  `setShapeAltTitle`, `getShapeDescription` / `setShapeDescription`.
- Shape predicates: `isChartShape`, `isTableShape`,
  `isShapeHidden` / `setShapeHidden`, `isShapePlaceholder`,
  `hasShapeImage`, `hasShapeText`.
- Shape search: `findShapeByText`, `findShapesByText`,
  `findShapesWithImages`, `findChartByKind`,
  `findChartsBySeriesName`, `findCommentsByAuthor`,
  `findSlidePlaceholders`, `findSlidePlaceholderByIdx`.
- Mutation: `setShapeRunHyperlink`, `getShapeRunHyperlink`,
  `setShapeBody`, `getSlideBody`, `appendShapeText`,
  `appendSlideNotes`, `removeSlideNotes`,
  `swapSlides`, `mergePresentations`, `slidesUsingMediaPart`,
  `setTableColumnWidth`, `setTableRowHeight`, `getTableColumnWidths`,
  `getTableRowHeights`, `getTableCellAlignment`, `getTableCellFill`.
- Diagnostics: `getSlideXmlString`, `getSlidePartName`,
  `getSlideLayoutPartName`, `getSlidesByLayout`.
