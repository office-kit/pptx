<script lang="ts">
  import { base } from '$app/paths';

  type Group = {
    title: string;
    description: string;
    exports: ReadonlyArray<{ name: string; signature?: string }>;
  };

  const REPO = 'https://github.com/office-kit/pptx/blob/main/src/api/fn.ts';

  const groups: ReadonlyArray<Group> = [
    {
      title: 'Load / save',
      description: 'Open and persist a .pptx, plus the Node fs convenience subpath.',
      exports: [
        { name: 'loadPresentation', signature: '(input: Uint8Array | ArrayBuffer | Blob) => Promise<PresentationData>' },
        { name: 'savePresentation', signature: '(pres: PresentationData) => Promise<Uint8Array>' },
        { name: 'createPresentation', signature: '() => PresentationData' },
        { name: 'loadPresentationFile', signature: '(path: string) => Promise<PresentationData> — @office-kit/pptx/node' },
        { name: 'savePresentationToFile', signature: '(pres: PresentationData, path: string) => Promise<void> — @office-kit/pptx/node' },
      ],
    },
    {
      title: 'Units',
      description: 'Branded Emu constructors. Internal code only sees Emu — every numeric position is type-checked at the boundary.',
      exports: [
        { name: 'inches', signature: '(n: number) => Emu' },
        { name: 'cm', signature: '(n: number) => Emu' },
        { name: 'mm', signature: '(n: number) => Emu' },
        { name: 'pt', signature: '(n: number) => Emu' },
        { name: 'emu', signature: '(n: number) => Emu — escape hatch' },
      ],
    },
    {
      title: 'Slides',
      description: 'Enumerate, add, remove, reorder. Layouts come from the loaded template.',
      exports: [
        { name: 'getSlides' },
        { name: 'getSlideAt' },
        { name: 'getSlideCount' },
        { name: 'getSlideIndex' },
        { name: 'addSlide' },
        { name: 'addSlideAt' },
        { name: 'addBlankSlide' },
        { name: 'addTitleSlide' },
        { name: 'addContentSlide' },
        { name: 'addSectionHeaderSlide' },
        { name: 'removeSlide' },
        { name: 'duplicateSlide' },
        { name: 'duplicateSlideAt' },
        { name: 'moveSlide' },
        { name: 'swapSlides' },
        { name: 'reverseSlides' },
        { name: 'sortSlides' },
        { name: 'importSlide' },
        { name: 'mergePresentations' },
      ],
    },
    {
      title: 'Slide layouts',
      description: 'Inspect and switch slides between the layouts shipped by the deck\'s masters.',
      exports: [
        { name: 'getSlideLayouts' },
        { name: 'findSlideLayout' },
        { name: 'findSlideLayoutByPartName' },
        { name: 'findSlideLayoutByType' },
        { name: 'findLayoutsWithPlaceholderType' },
        { name: 'getSlideLayout' },
        { name: 'setSlideLayout' },
        { name: 'getSlideLayoutName' },
        { name: 'getSlideLayoutType' },
        { name: 'getSlideLayoutPlaceholders' },
      ],
    },
    {
      title: 'Slide metadata',
      description: 'Title, size, visibility, sections, background, transitions, speaker notes.',
      exports: [
        { name: 'getSlideTitle' },
        { name: 'setSlideTitle' },
        { name: 'getSlideSize' },
        { name: 'setSlideSize' },
        { name: 'isSlideHidden' },
        { name: 'setSlideHidden' },
        { name: 'getSlideSections' },
        { name: 'setSlideSections' },
        { name: 'getSlideBackground' },
        { name: 'setSlideBackground' },
        { name: 'setSlideBackgroundImage' },
        { name: 'clearSlideBackground' },
        { name: 'getSlideNotes' },
        { name: 'setSlideNotes' },
        { name: 'appendSlideNotes' },
        { name: 'removeSlideNotes' },
        { name: 'getSlideTransition' },
        { name: 'setSlideTransition' },
        { name: 'clearSlideTransition' },
      ],
    },
    {
      title: 'Placeholders & text',
      description: 'Find placeholders by type, replace tokens, set per-run / per-paragraph formatting.',
      exports: [
        { name: 'findSlidePlaceholder' },
        { name: 'findSlidePlaceholderByIdx' },
        { name: 'findSlidePlaceholders' },
        { name: 'findEmptyPlaceholders' },
        { name: 'setSlidePlaceholders' },
        { name: 'setSlideBody' },
        { name: 'getSlideBody' },
        { name: 'replaceTokensInPresentation' },
        { name: 'replaceTokensInSlide' },
        { name: 'replaceTextInPresentation' },
        { name: 'replaceTextInSlide' },
        { name: 'replaceTextInNotes' },
        { name: 'replaceTextInSlideNotes' },
        { name: 'setShapeText' },
        { name: 'appendShapeText' },
        { name: 'setShapeRunText' },
        { name: 'setShapeRunFormat' },
        { name: 'setShapeTextFormat' },
        { name: 'setParagraphAlignment' },
        { name: 'setParagraphLevel' },
        { name: 'setParagraphBullet' },
        { name: 'setParagraphSpacing' },
        { name: 'setShapeBullets' },
        { name: 'setShapeAlignment' },
        { name: 'setShapeTextAnchor' },
        { name: 'setShapeTextMargins' },
        { name: 'setShapeTextWrap' },
        { name: 'setShapeTextAutoFit' },
      ],
    },
    {
      title: 'Shapes',
      description: 'Add, find, mutate, remove shapes. 180+ preset geometries.',
      exports: [
        { name: 'addSlideShape' },
        { name: 'addSlideTextBox' },
        { name: 'addSlideLine' },
        { name: 'addSlideImage' },
        { name: 'addSlideTable' },
        { name: 'addSlideChart' },
        { name: 'addSlideComment' },
        { name: 'getSlideShapes' },
        { name: 'getShapeAt' },
        { name: 'getAllShapes' },
        { name: 'findShapeById' },
        { name: 'findShapeByName' },
        { name: 'findShapeByText' },
        { name: 'findShapesByKind' },
        { name: 'findShapesByName' },
        { name: 'findShapesByText' },
        { name: 'findShapesAtPoint' },
        { name: 'findShapesOutsideCanvas' },
        { name: 'findFlippedShapes' },
        { name: 'findOverlappingShapePairs' },
        { name: 'findShapeInPresentation' },
        { name: 'copyShape' },
        { name: 'removeShape' },
        { name: 'renameShape' },
        { name: 'clearSlideShapes' },
      ],
    },
    {
      title: 'Shape geometry',
      description: 'Position, size, rotation, flip, z-order.',
      exports: [
        { name: 'getShapePosition' },
        { name: 'setShapePosition' },
        { name: 'getShapeSize' },
        { name: 'setShapeSize' },
        { name: 'getShapeBounds' },
        { name: 'setShapeBounds' },
        { name: 'getShapesBounds' },
        { name: 'translateShapes' },
        { name: 'centerShapeOnSlide' },
        { name: 'getShapeCenter' },
        { name: 'getShapeRotation' },
        { name: 'setShapeRotation' },
        { name: 'getShapeFlip' },
        { name: 'setShapeFlip' },
        { name: 'getShapeZIndex' },
        { name: 'setShapeZIndex' },
        { name: 'bringShapeToFront' },
        { name: 'sendShapeToBack' },
        { name: 'bringShapeForward' },
        { name: 'sendShapeBackward' },
      ],
    },
    {
      title: 'Fill, stroke, effects',
      description: 'Solid / gradient / pattern / image fill + stroke + shadow / glow.',
      exports: [
        { name: 'getShapeFill' },
        { name: 'setShapeFill' },
        { name: 'setShapeGradientFill' },
        { name: 'setShapePatternFill' },
        { name: 'setShapeImageFill' },
        { name: 'setShapeNoFill' },
        { name: 'clearShapeFill' },
        { name: 'getShapeFillColor' },
        { name: 'getShapeStroke' },
        { name: 'setShapeStroke' },
        { name: 'setShapeStrokeDash' },
        { name: 'setShapeStrokeArrow' },
        { name: 'setShapeNoStroke' },
        { name: 'clearShapeStroke' },
        { name: 'getShapeStrokeColor' },
        { name: 'getShapeStrokeWidth' },
        { name: 'getShapeStrokeDash' },
        { name: 'getShapeStrokeArrow' },
        { name: 'getShapeEffect' },
        { name: 'setShapeShadow' },
        { name: 'setShapeGlow' },
        { name: 'clearShapeEffects' },
      ],
    },
    {
      title: 'Images',
      description: 'Embedded picture shapes — bytes, crop, opacity, brightness, contrast.',
      exports: [
        { name: 'hasShapeImage' },
        { name: 'setShapeImage' },
        { name: 'getShapeImageBytes' },
        { name: 'getShapeImageFormat' },
        { name: 'getShapeImagePartName' },
        { name: 'setShapeImageCrop' },
        { name: 'getShapeImageCrop' },
        { name: 'setShapeImageOpacity' },
        { name: 'getShapeImageOpacity' },
        { name: 'setShapeImageBrightness' },
        { name: 'getShapeImageBrightness' },
        { name: 'setShapeImageContrast' },
        { name: 'getShapeImageContrast' },
        { name: 'getShapeImageFillBytes' },
      ],
    },
    {
      title: 'Tables',
      description: 'Per-cell text / fill / alignment, structural insert / remove, sizing.',
      exports: [
        { name: 'isTableShape' },
        { name: 'getTableSize' },
        { name: 'getTableDimensions' },
        { name: 'getTableRowHeights' },
        { name: 'getTableColumnWidths' },
        { name: 'setTableRowHeight' },
        { name: 'setTableColumnWidth' },
        { name: 'getTableCells' },
        { name: 'getTableCell' },
        { name: 'getTableCellText' },
        { name: 'setTableCellText' },
        { name: 'getTableCellPosition' },
        { name: 'getTableCellFill' },
        { name: 'setTableCellFill' },
        { name: 'clearTableCellFill' },
        { name: 'getTableCellAlignment' },
        { name: 'setTableCellAlignment' },
        { name: 'setTableCellTextFormat' },
        { name: 'insertTableRow' },
        { name: 'insertTableColumn' },
        { name: 'removeTableRow' },
        { name: 'removeTableColumn' },
      ],
    },
    {
      title: 'Charts',
      description: 'Embedded charts with auto-generated xlsx. Kinds: bar, column, line, pie, doughnut, area.',
      exports: [
        { name: 'isChartShape' },
        { name: 'getAllCharts' },
        { name: 'getSlideCharts' },
        { name: 'getShapeChartSpec' },
        { name: 'setChartSpec' },
        { name: 'getShapeChartKind' },
        { name: 'getShapeChartCategories' },
        { name: 'getShapeChartSeriesNames' },
        { name: 'getShapeChartSeriesValues' },
        { name: 'findChartByKind' },
        { name: 'findChartsBySeriesName' },
      ],
    },
    {
      title: 'Hyperlinks & click actions',
      description: 'External URL hyperlinks plus in-deck click navigation.',
      exports: [
        { name: 'getShapeHyperlink' },
        { name: 'setShapeHyperlink' },
        { name: 'getShapeRunHyperlink' },
        { name: 'setShapeRunHyperlink' },
        { name: 'getShapeClickAction' },
        { name: 'setShapeClickAction' },
        { name: 'getAllHyperlinks' },
        { name: 'getDistinctHyperlinkUrls' },
        { name: 'getSlidesWithHyperlinks' },
        { name: 'findSlidesByHyperlink' },
        { name: 'clearSlideHyperlinks' },
        { name: 'clearAllHyperlinks' },
        { name: 'replaceHyperlink' },
      ],
    },
    {
      title: 'Animations',
      description: 'Entrance / exit / emphasis presets. Complex timing-tree authoring is post-1.0.',
      exports: [
        { name: 'getShapeAnimation' },
        { name: 'setShapeAnimation' },
        { name: 'clearSlideAnimations' },
      ],
    },
    {
      title: 'Comments',
      description: 'Per-slide review comments with author metadata and EMU position.',
      exports: [
        { name: 'addSlideComment' },
        { name: 'removeSlideComment' },
        { name: 'clearSlideComments' },
        { name: 'clearAllSlideComments' },
        { name: 'getSlideComments' },
        { name: 'getAllComments' },
        { name: 'getCommentAuthor' },
        { name: 'getCommentAuthors' },
        { name: 'getCommentDate' },
        { name: 'getCommentPosition' },
        { name: 'getCommentSlide' },
        { name: 'getCommentText' },
        { name: 'getCommentsSortedByDate' },
        { name: 'findCommentAuthorByName' },
        { name: 'findCommentsAfter' },
        { name: 'findCommentsBefore' },
        { name: 'findCommentsByAuthor' },
        { name: 'findCommentsByText' },
        { name: 'getSlideCommentAuthors' },
        { name: 'getPresentationCommenters' },
        { name: 'getSlidesWithComments' },
        { name: 'findSlidesWithCommentsByAuthor' },
      ],
    },
    {
      title: 'Document properties',
      description: 'Title, Author, Subject, Keywords, App, Company, thumbnail.',
      exports: [
        { name: 'getCoreProperties' },
        { name: 'setCoreProperties' },
        { name: 'getExtendedProperties' },
        { name: 'setExtendedProperties' },
        { name: 'getPresentationCreated' },
        { name: 'getPresentationModified' },
        { name: 'touchModified' },
        { name: 'incrementRevision' },
        { name: 'getPresentationTheme' },
        { name: 'getPresentationFonts' },
        { name: 'getPresentationSummary' },
        { name: 'getThumbnail' },
        { name: 'setThumbnail' },
        { name: 'removeThumbnail' },
      ],
    },
    {
      title: 'Diagnostics',
      description: 'Validation, raw OPC introspection, media accounting.',
      exports: [
        { name: 'validatePresentation' },
        { name: 'listPackageParts' },
        { name: 'readPackagePart' },
        { name: 'getPackageSize' },
        { name: 'getMediaParts' },
        { name: 'setMediaPartBytes' },
        { name: 'getOrphanMediaPartNames' },
        { name: 'getSlideMediaPartNames' },
        { name: 'slidesUsingMediaPart' },
        { name: 'compactPackage' },
        { name: 'getPresentationText' },
        { name: 'getPresentationTextLength' },
        { name: 'getPresentationNotesText' },
        { name: 'getPresentationNotesLength' },
      ],
    },
  ];

  const total = groups.reduce((n, g) => n + g.exports.length, 0);
</script>

<svelte:head>
  <title>API reference — @office-kit/pptx</title>
</svelte:head>

<div class="content">
  <p class="eyebrow">Reference</p>
  <h1>API reference</h1>
  <p class="lede">
    Every public export of <code>@office-kit/pptx</code> and <code>@office-kit/pptx/node</code>, organized
    by category. The library is fn-only — no classes — so this list is the entire
    callable surface. {total} exports across {groups.length} groups.
  </p>
  <p class="lede">
    For the conceptual map and migration notes see the
    <a href="{base}/docs/api">API overview</a>. For copy-pasteable code see the
    <a href="{base}/docs/cheatsheet">Cheatsheet</a> and <a href="{base}/docs/recipes">Recipes</a>.
  </p>

  {#each groups as group, i (group.title)}
    <section class="group" id={group.title.toLowerCase().replace(/[^a-z0-9]+/g, '-')}>
      <header class="g-head">
        <span class="g-num">{String(i + 1).padStart(2, '0')}</span>
        <div>
          <h2>{group.title}</h2>
          <p>{group.description}</p>
        </div>
        <span class="g-count">{group.exports.length}</span>
      </header>
      <ul class="exports">
        {#each group.exports as ex (ex.name)}
          <li>
            <a href={REPO} rel="noopener" target="_blank">
              <code class="ex-name">{ex.name}</code>
            </a>
            {#if ex.signature}<span class="sig">{ex.signature}</span>{/if}
          </li>
        {/each}
      </ul>
    </section>
  {/each}
</div>

<style>
  .content {
    max-width: var(--max-wide);
    margin: 0 auto;
    padding: 2.25rem 1.5rem 5rem;
  }

  .eyebrow {
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--fg-muted);
    text-transform: uppercase;
    letter-spacing: 0.14em;
    margin: 0 0 0.85rem;
  }

  h1 {
    margin: 0 0 0.6rem;
    font-family: var(--display);
    font-weight: 460;
    font-size: clamp(2rem, 4.6vw, 2.95rem);
    line-height: 1.05;
    letter-spacing: -0.026em;
    font-variation-settings: 'opsz' 144, 'SOFT' 30;
    border: none;
    padding: 0;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.05rem;
    margin: 0 0 1rem;
    max-width: 72ch;
  }

  .group {
    margin-top: 2.75rem;
    scroll-margin-top: calc(var(--header-h) + 1rem);
  }

  .g-head {
    display: grid;
    grid-template-columns: 4ch 1fr auto;
    gap: 1rem;
    align-items: baseline;
    border-bottom: 1px solid var(--border);
    padding-bottom: 0.65rem;
    margin-bottom: 1rem;
  }

  .g-num {
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--accent);
    font-weight: 500;
    letter-spacing: 0.06em;
  }

  .g-head h2 {
    font-family: var(--display);
    font-weight: 500;
    font-size: 1.45rem;
    letter-spacing: -0.015em;
    margin: 0 0 0.25rem;
    border: none;
    padding: 0;
    font-variation-settings: 'opsz' 96, 'SOFT' 25;
  }

  .g-head p {
    color: var(--fg-muted);
    font-size: 0.9rem;
    margin: 0;
  }

  .g-count {
    font-family: var(--mono);
    font-size: 10.5px;
    color: var(--fg-muted);
    text-transform: uppercase;
    letter-spacing: 0.12em;
  }

  .exports {
    list-style: none;
    margin: 0;
    padding: 0;
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(240px, 1fr));
    gap: 0.45rem 1.2rem;
  }

  .exports li {
    padding: 0.25rem 0;
    line-height: 1.45;
  }

  .ex-name {
    font-family: var(--mono);
    font-size: 0.85rem;
    color: var(--fg);
    background: transparent;
    border: none;
    padding: 0;
  }

  .exports a:hover .ex-name {
    color: var(--accent);
    text-decoration: underline;
  }

  .sig {
    display: block;
    font-family: var(--mono);
    font-size: 11px;
    color: var(--fg-muted);
    margin-top: 0.15rem;
    word-break: break-word;
  }
</style>
