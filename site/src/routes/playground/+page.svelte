<script lang="ts">
  import { base } from '$app/paths';

  // @office-kit/pptx is consumed via the source-tree path alias set in
  // svelte.config.js (`@office-kit/pptx` → `../src/index.ts`). The playground
  // exercises the real surface the same way the rest of the docs site
  // does — type errors here break the build.
  import {
    getCommentText,
    getCoreProperties,
    getPresentationChartKindCounts,
    getPresentationCommentCountsByAuthor,
    getPresentationSummary,
    getSlideLayoutUsageCountsByType,
    getShapeHyperlink,
    getSlideCharts,
    getSlideMasterCount,
    getSlideMediaPartNames,
    getSlideTables,
    type ValidationIssue,
    validatePresentation,
    getShapeKind,
    getSlideComments,
    getSlideLayout,
    getSlideIndex,
    getSlideLayoutName,
    getSlideLayoutType,
    getSlideNotes,
    getSlideSections,
    getSlideShapes,
    getSlideTitle,
    getSlideTransition,
    getSlides,
    getSlideTextLength,
    isSlideHidden,
    listPackageParts,
    loadPresentation,
    savePresentation,
    slideHasAnimations,
  } from '@office-kit/pptx';
  import { renderSlideToSvg } from '@office-kit/pptx-preview';

  type SlideSnapshot = {
    index: number;
    title: string;
    textLength: number;
    shapeKinds: string[];
    svg: string;
    notes: string | null;
    hasTransition: boolean;
    hasAnimations: boolean;
    hidden: boolean;
    commentCount: number;
    commentTexts: string;
    layoutType: string | null;
    layoutName: string | null;
    chartCount: number;
    tableCount: number;
    hyperlinkCount: number;
    mediaCount: number;
  };

  type PackagePart = { name: string; contentType: string; byteLength: number };

  let fileName = $state<string>('');
  let status = $state<string>('');
  let busy = $state<boolean>(false);
  let dropping = $state<boolean>(false);
  let slideCount = $state<number>(0);
  let coreTitle = $state<string>('');
  let coreCreator = $state<string>('');
  let summary = $state<ReturnType<typeof getPresentationSummary> | null>(null);
  let chartKindCounts = $state<ReturnType<typeof getPresentationChartKindCounts> | null>(null);
  let layoutTypeCounts = $state<ReturnType<typeof getSlideLayoutUsageCountsByType> | null>(null);
  let commentAuthorCounts = $state<ReturnType<typeof getPresentationCommentCountsByAuthor> | null>(
    null,
  );
  let masterCount = $state<number>(0);
  let issues = $state<ReadonlyArray<ValidationIssue>>([]);
  let slides = $state<SlideSnapshot[]>([]);
  let sectionStartByIndex = $state<Record<number, string>>({});
  let parts = $state<PackagePart[]>([]);
  let lastBytes = $state<Uint8Array | null>(null);

  async function inspect(bytes: Uint8Array, source: string) {
    busy = true;
    status = 'Parsing…';
    try {
      const pres = await loadPresentation(bytes);
      const core = getCoreProperties(pres);
      coreTitle = core?.title ?? '';
      coreCreator = core?.creator ?? '';
      summary = getPresentationSummary(pres);
      chartKindCounts = getPresentationChartKindCounts(pres);
      layoutTypeCounts = getSlideLayoutUsageCountsByType(pres);
      commentAuthorCounts = getPresentationCommentCountsByAuthor(pres);
      masterCount = getSlideMasterCount(pres);
      issues = validatePresentation(pres);
      // Map section name → 1-based slide index of its first slide. We
      // surface them so the playground can render a divider above
      // each section's first slide.
      const sectionMap: Record<number, string> = {};
      for (const sec of getSlideSections(pres)) {
        const first = sec.slides[0];
        if (!first) continue;
        const idx0 = getSlideIndex(pres, first);
        if (idx0 >= 0) sectionMap[idx0 + 1] = sec.name;
      }
      sectionStartByIndex = sectionMap;

      const list = getSlides(pres);
      slideCount = list.length;
      slides = list.map((slide, i) => {
        const layout = getSlideLayout(slide);
        const layoutType = layout ? getSlideLayoutType(layout) : null;
        const layoutName = layout ? getSlideLayoutName(layout) : null;
        return {
          index: i + 1,
          title: getSlideTitle(slide) ?? '',
          textLength: getSlideTextLength(slide),
          shapeKinds: getSlideShapes(slide).map((sh) => getShapeKind(sh)),
          svg: renderSlideToSvg(pres, slide),
          notes: getSlideNotes(slide),
          hasTransition: getSlideTransition(slide) !== null,
          hasAnimations: slideHasAnimations(slide),
          hidden: isSlideHidden(slide),
          commentCount: getSlideComments(slide).length,
          commentTexts: getSlideComments(slide)
            .map((c) => getCommentText(c))
            .filter((t) => t.length > 0)
            .join('\n'),
          layoutType,
          layoutName,
          chartCount: getSlideCharts(slide).length,
          tableCount: getSlideTables(slide).length,
          hyperlinkCount: getSlideShapes(slide).filter((sh) => getShapeHyperlink(sh) !== null)
            .length,
          mediaCount: getSlideMediaPartNames(slide).length,
        };
      });

      parts = listPackageParts(pres)
        .slice()
        .sort((a, b) => a.name.localeCompare(b.name));

      // Re-save so the round-trip button has something to offer.
      lastBytes = await savePresentation(pres);
      fileName = source;
      status = `Parsed ${list.length} ${list.length === 1 ? 'slide' : 'slides'} and ${parts.length} package parts.`;
    } catch (err) {
      status = `This file could not be parsed: ${err instanceof Error ? err.message : String(err)}`;
      slides = [];
      parts = [];
    } finally {
      busy = false;
    }
  }

  const deckFlags = $derived.by(() => {
    if (!summary) return [];
    const flags: string[] = [];
    if (summary.hiddenSlideCount > 0) flags.push(`${summary.hiddenSlideCount} hidden`);
    if (summary.hasCharts) flags.push('charts');
    if (summary.hasComments) flags.push('comments');
    if (summary.hasAnimations) flags.push('animations');
    return flags;
  });

  // Gives a visitor with no .pptx at hand something real to inspect: the deck
  // from the landing page, built in this tab.
  async function loadDemoDeck() {
    const { buildHeroDeck } = await import('$lib/examples/hero-deck');
    await inspect(await savePresentation(buildHeroDeck()), 'office-kit-demo.pptx');
  }

  async function onFileChosen(file: File) {
    const buf = await file.arrayBuffer();
    await inspect(new Uint8Array(buf), file.name);
  }

  async function onDrop(ev: DragEvent) {
    ev.preventDefault();
    dropping = false;
    const file = ev.dataTransfer?.files[0];
    if (file) await onFileChosen(file);
  }

  function downloadRoundtrip() {
    if (!lastBytes) return;
    const blob = new Blob([lastBytes as BlobPart], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName.replace(/\.pptx$/, '') + '.roundtrip.pptx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
</script>

<svelte:head>
  <title>Playground · @office-kit/pptx</title>
</svelte:head>

<section class="content">
  <h1>Inspect a .pptx in your browser</h1>
  <p class="lede">
    Drop a file and this page parses it with the real <code>@office-kit/pptx</code> source, draws
    every slide with <code>@office-kit/pptx-preview</code>, validates the package, and lists its
    parts. Nothing is uploaded: the whole pipeline runs in this tab.
  </p>

  <div
    class="drop"
    class:dropping
    role="group"
    aria-label="Choose a .pptx file"
    ondragover={(e) => {
      e.preventDefault();
      dropping = true;
    }}
    ondragleave={() => (dropping = false)}
    ondrop={onDrop}
  >
    <p class="drop-text">{fileName || 'Drop a .pptx file here'}</p>
    <div class="drop-actions">
      <label class="btn primary drop-pick">
        <input
          type="file"
          accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
          onchange={(e) => {
            const f = (e.currentTarget as HTMLInputElement).files?.[0];
            if (f) onFileChosen(f);
          }}
        />
        Choose a file
      </label>
      <button type="button" class="btn" onclick={loadDemoDeck} disabled={busy}>
        Load the demo deck
      </button>
      {#if lastBytes}
        <button type="button" class="btn" onclick={downloadRoundtrip}>
          Download the re-saved file
        </button>
      {/if}
    </div>
  </div>

  <p class="caveat">
    The preview renders preset and custom geometry, theme and placeholder inheritance, gradient,
    pattern, and picture fills, effects, charts, and tables. SmartArt, animations, and 3D show as
    labelled placeholders. PowerPoint and LibreOffice remain the pixel-exact renderers.
  </p>

  <p class="status" class:busy aria-live="polite">{status}</p>

  {#if slideCount > 0}
    <div class="meta-clip"><div class="meta">
      <div class="cell">
        <span class="label">Slides</span>
        <span class="value">{slideCount}</span>
      </div>
      <div class="cell">
        <span class="label">Package parts</span>
        <span class="value">{parts.length}</span>
      </div>
      <div class="cell">
        <span class="label">Title</span>
        <span class="value">{coreTitle || 'Not set'}</span>
      </div>
      <div class="cell">
        <span class="label">Creator</span>
        <span class="value">{coreCreator || 'Not set'}</span>
      </div>
      {#if summary}
        <div class="cell">
          <span class="label">Theme</span>
          <span class="value">{summary.themeName ?? 'Not set'}</span>
        </div>
        <div class="cell">
          <span class="label">Masters, layouts, sections</span>
          <span class="value">{masterCount}, {summary.layoutCount}, {summary.sectionCount}</span>
        </div>
        <div class="cell">
          <span class="label">Shapes</span>
          <span class="value">{summary.totalShapes}</span>
        </div>
        <div class="cell">
          <span class="label">Also contains</span>
          <span class="value">{deckFlags.length > 0 ? deckFlags.join(', ') : 'Nothing else'}</span>
        </div>
        {#if chartKindCounts && Object.values(chartKindCounts).some((n) => n > 0)}
          <div class="cell">
            <span class="label">Chart kinds</span>
            <span class="value">
              {Object.entries(chartKindCounts)
                .filter(([, n]) => n > 0)
                .map(([k, n]) => `${n} ${k}`)
                .join(', ')}
            </span>
          </div>
        {/if}
        {#if layoutTypeCounts && Object.keys(layoutTypeCounts).length > 0}
          <div class="cell">
            <span class="label">Layout types in use</span>
            <span class="value">
              {Object.entries(layoutTypeCounts)
                .map(([k, n]) => `${n} ${k}`)
                .join(', ')}
            </span>
          </div>
        {/if}
        {#if commentAuthorCounts && Object.keys(commentAuthorCounts).length > 0}
          <div class="cell">
            <span class="label">Comment authors</span>
            <span class="value">
              {Object.entries(commentAuthorCounts)
                .sort(([, a], [, b]) => b - a)
                .map(([k, n]) => `${k} (${n})`)
                .join(', ')}
            </span>
          </div>
        {/if}
      {/if}
    </div></div>

    {#if issues.length > 0}
      <h2>Validation</h2>
      <ul class="issues">
        {#each issues as iss (iss.message)}
          <li class={`issue issue-${iss.severity}`}>
            <span class="issue-sev">{iss.severity}</span>
            <span class="issue-msg">{iss.message}</span>
            {#if iss.partName}<code class="issue-part">{iss.partName}</code>{/if}
          </li>
        {/each}
      </ul>
    {/if}

    <h2>Slides</h2>
    <ol class="slides">
      {#each slides as s (s.index)}
        {#if sectionStartByIndex[s.index] !== undefined}
          <li class="section-divider" aria-label="section">
            <span class="section-name">{sectionStartByIndex[s.index]}</span>
          </li>
        {/if}
        <li id={`slide-${s.index}`}>
          <div class="s-head">
            <a class="s-num" href={`#slide-${s.index}`} title="Link to this slide">Slide {s.index}</a>
            <span class="s-title">{s.title || '(untitled)'}</span>
            {#if s.layoutType}<span class="s-badge" title={s.layoutName ? `layout: ${s.layoutName} (type: ${s.layoutType})` : `slide layout type: ${s.layoutType}`}>{s.layoutType}</span>{/if}
            {#if s.hidden}<span class="s-badge s-badge-hidden" title='show="0" — hidden from slideshow'>hidden</span>{/if}
            {#if s.hasTransition}<span class="s-badge" title="slide carries <p:transition>">trans</span>{/if}
            {#if s.hasAnimations}<span class="s-badge" title="slide carries <p:timing>">anim</span>{/if}
            {#if s.commentCount > 0}<span class="s-badge" title={s.commentTexts || 'slide has authored review comments'}>{s.commentCount} cmt</span>{/if}
            {#if s.chartCount > 0}<span class="s-badge" title="number of <p:graphicFrame> chart shapes on the slide">{s.chartCount} chart</span>{/if}
            {#if s.tableCount > 0}<span class="s-badge" title="number of <p:graphicFrame> table shapes on the slide">{s.tableCount} table</span>{/if}
            {#if s.hyperlinkCount > 0}<span class="s-badge" title="shapes whose text body carries an <a:hlinkClick>">{s.hyperlinkCount} link</span>{/if}
            {#if s.mediaCount > 0}<span class="s-badge" title="number of media parts (images / audio / video) the slide references">{s.mediaCount} media</span>{/if}
            {#if s.notes && s.notes.length > 0}<span class="s-badge" title="speaker notes character count">{s.notes.length} notes</span>{/if}
            <span class="s-len">{s.textLength} characters, {s.shapeKinds.length} shapes</span>
          </div>
          <div class="s-canvas">
            {@html s.svg}
          </div>
          {#if s.shapeKinds.length > 0}
            <p class="s-kinds">
              {#each Array.from(new Set(s.shapeKinds)) as k (k)}
                <code>{k}</code>
              {/each}
            </p>
          {/if}
          {#if s.notes}
            <details class="s-notes">
              <summary>speaker notes ({s.notes.length} chars)</summary>
              <pre>{s.notes}</pre>
            </details>
          {/if}
        </li>
      {/each}
    </ol>

    <h2>Package parts</h2>
    <div class="table-scroll">
    <table class="parts">
      <thead>
        <tr><th>Part</th><th>Content type</th><th>Bytes</th></tr>
      </thead>
      <tbody>
        {#each parts as p (p.name)}
          <tr>
            <td><code>{p.name}</code></td>
            <td class="ct">{p.contentType}</td>
            <td class="num">{p.byteLength.toLocaleString()}</td>
          </tr>
        {/each}
      </tbody>
    </table>
    </div>
  {/if}
</section>

<style>
  .content {
    max-width: 1000px;
    margin: 0 auto;
    padding: 2.75rem var(--gutter) 5rem;
  }

  .lede {
    max-width: 66ch;
    color: var(--ink-2);
    font-size: 1.08rem;
  }

  .drop {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 1.1rem;
    margin: 2rem 0 0;
    padding: 2.25rem 1.25rem;
    border: 1.5px dashed var(--line-strong);
    border-radius: 12px;
    background: var(--wash);
    text-align: center;
    transition:
      border-color 120ms ease,
      background 120ms ease;
  }

  .drop.dropping {
    border-color: var(--accent);
    background: var(--accent-wash);
  }

  .drop-text {
    margin: 0;
    font-family: var(--display);
    font-size: 1.25rem;
    font-weight: 600;
    letter-spacing: -0.015em;
    overflow-wrap: anywhere;
  }

  .drop-actions {
    display: flex;
    flex-wrap: wrap;
    justify-content: center;
    gap: 0.6rem;
  }

  .drop-pick input {
    position: absolute;
    width: 1px;
    height: 1px;
    opacity: 0;
  }

  .drop-pick:focus-within {
    outline: 2px solid var(--accent);
    outline-offset: 2px;
  }

  .btn:disabled {
    opacity: 0.6;
    cursor: progress;
  }

  .caveat {
    max-width: 72ch;
    margin: 1rem 0 0;
    color: var(--ink-3);
    font-size: 0.88rem;
  }

  .status {
    min-height: 1.6em;
    margin: 1.5rem 0 0;
    color: var(--ink-2);
    font-size: 0.95rem;
  }

  .status.busy {
    color: var(--accent-ink);
  }

  h2 {
    margin: 3rem 0 1rem;
    padding-bottom: 0.6rem;
    border-bottom: 1px solid var(--line);
    font-size: 1.4rem;
  }

  /* Each cell draws its own right and bottom rule and the grid is pulled 1px
   * past the clipping box, so the outer edge never doubles up and a short last
   * row leaves plain paper rather than a filled gap. */
  .meta-clip {
    margin-top: 1rem;
    border: 1px solid var(--line);
    border-radius: var(--radius);
    overflow: hidden;
  }

  .meta {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    margin: 0 -1px -1px 0;
  }

  .cell {
    display: flex;
    flex-direction: column;
    gap: 0.15rem;
    padding: 0.8rem 1rem;
    border-right: 1px solid var(--line);
    border-bottom: 1px solid var(--line);
  }

  .label {
    color: var(--ink-3);
    font-size: 0.8rem;
  }

  .value {
    font-weight: 600;
    font-size: 0.97rem;
    overflow-wrap: anywhere;
  }

  .issues {
    list-style: none;
    margin: 0;
    padding: 0;
  }

  .issue {
    display: flex;
    flex-wrap: wrap;
    align-items: baseline;
    gap: 0.4rem 0.75rem;
    margin: 0;
    padding: 0.7rem 0;
    border-bottom: 1px solid var(--line);
    font-size: 0.93rem;
  }

  .issue-sev {
    flex: none;
    padding: 0.05rem 0.5rem;
    border-radius: 999px;
    background: var(--wash);
    border: 1px solid var(--line-strong);
    font-size: 0.78rem;
    font-weight: 600;
  }

  .issue-error .issue-sev {
    background: var(--accent-wash);
    border-color: var(--accent);
    color: var(--accent-ink);
  }

  .issue-msg {
    flex: 1 1 16rem;
  }

  .slides {
    list-style: none;
    margin: 0;
    padding: 0;
  }

  .slides > li {
    margin: 0 0 2.5rem;
    scroll-margin-top: calc(var(--header-h) + 1rem);
  }

  .section-divider {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    margin: 2.5rem 0 1.25rem !important;
    color: var(--ink-2);
    font-weight: 600;
  }

  .section-divider::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--line);
  }

  .s-head {
    display: flex;
    flex-wrap: wrap;
    align-items: baseline;
    gap: 0.35rem 0.6rem;
    margin-bottom: 0.6rem;
  }

  .s-num {
    font-family: var(--mono);
    font-size: 0.85rem;
    font-weight: 550;
  }

  .s-title {
    font-weight: 600;
  }

  .s-badge {
    padding: 0.02rem 0.45rem;
    border: 1px solid var(--line-strong);
    border-radius: 999px;
    color: var(--ink-2);
    font-size: 0.76rem;
    white-space: nowrap;
  }

  .s-badge-hidden {
    border-color: var(--accent);
    color: var(--accent-ink);
  }

  .s-len {
    margin-left: auto;
    color: var(--ink-3);
    font-size: 0.82rem;
  }

  /* Slides are always drawn on white: that is the page colour of the deck,
   * not of this site, so it must not follow the dark theme. */
  .s-canvas {
    border: 1px solid var(--line-strong);
    border-radius: 4px;
    background: #fff;
    overflow: hidden;
    box-shadow: var(--shadow-pop);
  }

  .s-canvas :global(svg) {
    display: block;
    width: 100%;
    height: auto;
  }

  .s-kinds {
    display: flex;
    flex-wrap: wrap;
    gap: 0.3rem;
    margin: 0.6rem 0 0;
  }

  .s-kinds code {
    font-size: 0.76rem;
  }

  .s-notes {
    margin-top: 0.6rem;
    font-size: 0.9rem;
  }

  .s-notes summary {
    cursor: pointer;
    color: var(--ink-2);
  }

  .s-notes pre {
    white-space: pre-wrap;
  }

  .parts {
    min-width: 560px;
    margin: 0;
    font-size: 0.85rem;
  }

  .parts code {
    padding: 0;
    border: none;
    background: none;
    overflow-wrap: anywhere;
  }

  .ct {
    color: var(--ink-2);
    overflow-wrap: anywhere;
  }

  .num {
    text-align: right;
    font-family: var(--mono);
    white-space: nowrap;
  }

  .parts th:last-child {
    text-align: right;
  }

  @media (max-width: 560px) {
    .drop-actions .btn {
      flex: 1 1 100%;
    }

    .s-len {
      margin-left: 0;
      flex-basis: 100%;
    }
  }
</style>
