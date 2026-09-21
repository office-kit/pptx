<script lang="ts">
  import { base } from '$app/paths';

  type Group = {
    title: string;
    description: string;
    exports: ReadonlyArray<{ name: string; signature?: string }>;
  };

  // eslint-disable-next-line prefer-const -- reassigned by `bind:value` in template
  let filter = $state('');

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
        { name: 'getShapeFillOpacity' },
        { name: 'getShapeStroke' },
        { name: 'setShapeStroke' },
        { name: 'setShapeStrokeDash' },
        { name: 'setShapeStrokeArrow' },
        { name: 'setShapeNoStroke' },
        { name: 'clearShapeStroke' },
        { name: 'getShapeStrokeColor' },
        { name: 'getShapeStrokeOpacity' },
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
      title: 'Media',
      description: 'Embedded video / audio and online (URL) video — add and read back.',
      exports: [
        { name: 'addSlideMedia' },
        { name: 'getShapeMedia' },
        { name: 'findShapesWithMedia' },
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

  const slug = (title: string): string => title.toLowerCase().replace(/[^a-z0-9]+/g, '-');

  const visible = $derived.by(() => {
    const q = filter.trim().toLowerCase();
    if (!q) return groups;
    return groups
      .map((g) => ({ ...g, exports: g.exports.filter((e) => e.name.toLowerCase().includes(q)) }))
      .filter((g) => g.exports.length > 0);
  });

  const visibleCount = $derived(visible.reduce((n, g) => n + g.exports.length, 0));
</script>

<svelte:head>
  <title>API reference · @office-kit/pptx</title>
</svelte:head>

<div class="api frame">
  <aside class="toc" data-pagefind-ignore>
    <nav aria-label="API groups">
      <h2>Groups</h2>
      <ul>
        {#each groups as group (group.title)}
          <li><a href="#{slug(group.title)}">{group.title}</a></li>
        {/each}
      </ul>
    </nav>
  </aside>

  <div class="content">
    <h1>API reference</h1>
    <p class="lede">
      Every public export of <code>@office-kit/pptx</code> and <code>@office-kit/pptx/node</code>,
      by category. The library is functions only, with no classes. This page covers the {total}
      exports people reach for most, in {groups.length} groups; the package's type declarations
      list every one.
    </p>
    <p class="lede">
      For the conceptual map see the <a href="{base}/docs/api">API overview</a>. For code you can
      paste, see the <a href="{base}/docs/cheatsheet">cheatsheet</a> and
      <a href="{base}/docs/recipes">recipes</a>.
    </p>

    <div class="filter" data-pagefind-ignore>
      <label for="api-filter">Filter by function name</label>
      <input
        id="api-filter"
        type="search"
        bind:value={filter}
        placeholder="setShape, Table, Notes…"
        autocomplete="off"
        spellcheck="false"
      />
      <p aria-live="polite">
        {#if filter.trim()}
          {visibleCount === 0
            ? `No function name contains “${filter.trim()}”.`
            : `${visibleCount} of ${total} exports`}
        {/if}
      </p>
    </div>

    {#each visible as group (group.title)}
      <section class="group" id={slug(group.title)}>
        <h2>{group.title}</h2>
        <p class="g-desc">{group.description}</p>
        <ul class="exports">
          {#each group.exports as ex (ex.name)}
            <li>
              <a href={REPO} rel="noopener" target="_blank"><code>{ex.name}</code></a>
              {#if ex.signature}<span class="sig">{ex.signature}</span>{/if}
            </li>
          {/each}
        </ul>
      </section>
    {/each}
  </div>
</div>

<style>
  .api {
    display: flex;
    align-items: stretch;
  }

  .toc {
    flex: 0 0 var(--sidebar-w);
    width: var(--sidebar-w);
    border-right: 1px solid var(--line);
  }

  .toc nav {
    position: sticky;
    top: var(--header-h);
    max-height: calc(100vh - var(--header-h));
    overflow-y: auto;
    padding: 2.25rem 1rem 3rem var(--gutter);
  }

  .toc h2 {
    margin: 0 0 0.5rem;
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 600;
    letter-spacing: 0;
    color: var(--ink-3);
  }

  .toc ul {
    list-style: none;
    margin: 0;
    padding: 0;
    border-left: 1px solid var(--line);
  }

  .toc li {
    margin: 0;
  }

  .toc a {
    display: block;
    padding: 0.28rem 0 0.28rem 0.95rem;
    color: var(--ink-2);
    font-size: 0.9rem;
  }

  .toc a:hover {
    color: var(--accent-ink);
    text-decoration: none;
  }

  .content {
    flex: 1;
    min-width: 0;
    padding: 2.75rem clamp(1.25rem, 4vw, 3.5rem) 5rem;
  }

  .lede {
    max-width: 68ch;
    color: var(--ink-2);
    font-size: 1.05rem;
  }

  .filter {
    position: sticky;
    top: var(--header-h);
    z-index: 5;
    display: grid;
    grid-template-columns: minmax(0, 24rem) 1fr;
    gap: 0.35rem 1rem;
    align-items: center;
    margin: 2rem 0 0;
    padding: 0.85rem 0;
    background: var(--paper);
    border-bottom: 1px solid var(--line);
  }

  .filter label {
    grid-column: 1 / -1;
    font-size: 0.85rem;
    font-weight: 600;
    color: var(--ink-2);
  }

  .filter input {
    width: 100%;
    height: 42px;
    padding: 0 0.85rem;
    border: 1px solid var(--line-strong);
    border-radius: var(--radius);
    background: var(--paper);
    color: var(--ink);
    font-family: var(--mono);
    /* 16px keeps iOS Safari from zooming the page when the field takes focus. */
    font-size: 1rem;
  }

  .filter input:focus-visible {
    outline-offset: 0;
    border-color: var(--accent);
  }

  .filter p {
    margin: 0;
    color: var(--ink-2);
    font-size: 0.9rem;
  }

  .group {
    margin-top: 3rem;
    scroll-margin-top: calc(var(--header-h) + 7rem);
  }

  .group h2 {
    margin: 0;
    font-size: 1.4rem;
  }

  .g-desc {
    max-width: 68ch;
    margin: 0.35rem 0 1rem;
    color: var(--ink-2);
    font-size: 0.97rem;
  }

  .exports {
    list-style: none;
    margin: 0;
    padding: 0;
    border-top: 1px solid var(--line);
    columns: 2;
    column-gap: 2.5rem;
  }

  .exports li {
    display: flex;
    flex-direction: column;
    gap: 0.15rem;
    margin: 0;
    padding: 0.5rem 0;
    border-bottom: 1px solid var(--line);
    break-inside: avoid;
  }

  .exports code {
    padding: 0;
    border: none;
    background: none;
    color: var(--accent-ink);
    font-size: 0.9rem;
    font-weight: 500;
  }

  .sig {
    color: var(--ink-2);
    font-family: var(--mono);
    font-size: 0.78rem;
    line-height: 1.5;
    overflow-wrap: anywhere;
  }

  @media (max-width: 1000px) {
    .toc {
      display: none;
    }
  }

  @media (max-width: 640px) {
    .exports {
      columns: 1;
    }

    .filter {
      grid-template-columns: 1fr;
    }
  }
</style>
