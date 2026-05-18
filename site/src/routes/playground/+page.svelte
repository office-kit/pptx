<script lang="ts">
  import { base } from '$app/paths';

  // pptx-kit is consumed via the source-tree path alias set in
  // svelte.config.js (`pptx-kit` → `../src/index.ts`). The playground
  // exercises the real surface the same way the rest of the docs site
  // does — type errors here break the build.
  import {
    getCommentText,
    getCoreProperties,
    getPresentationChartKindCounts,
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
  } from 'pptx-kit';
  import { renderSlideSvg } from '$lib/playground/render-slide';

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

  let fileName = $state<string>('(drop a .pptx here, or pick one)');
  let status = $state<string>('Ready.');
  let busy = $state<boolean>(false);
  let dropping = $state<boolean>(false);
  let slideCount = $state<number>(0);
  let coreTitle = $state<string>('');
  let coreCreator = $state<string>('');
  let summary = $state<ReturnType<typeof getPresentationSummary> | null>(null);
  let chartKindCounts = $state<ReturnType<typeof getPresentationChartKindCounts> | null>(null);
  let layoutTypeCounts = $state<ReturnType<typeof getSlideLayoutUsageCountsByType> | null>(null);
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
          svg: renderSlideSvg(pres, slide),
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
      status = `Parsed ${list.length} slide(s) · ${parts.length} OPC parts.`;
    } catch (err) {
      status = `Failed: ${err instanceof Error ? err.message : String(err)}`;
      slides = [];
      parts = [];
    } finally {
      busy = false;
    }
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
  <title>Playground · pptx-kit</title>
</svelte:head>

<section class="content">
  <p class="eyebrow">§ 03 · Playground</p>
  <h1>Inspect a <code>.pptx</code> in the browser.</h1>
  <p class="lede">
    Drop a file below. The page parses it with the real <code>pptx-kit</code> source from this
    repo, renders each slide's shapes as approximate SVG (preset geometry, fills, strokes,
    rotation, embedded images), and dumps the OPC parts list. No bytes leave your machine — the
    whole pipeline runs in this tab. For loading via fetch / fs see
    <a href="{base}/docs/getting-started">Getting started</a>.
  </p>
  <p class="caveat">
    <strong>Approximate.</strong> Shape geometry is reconstructed from the preset name; theme /
    placeholder inheritance, gradient / pattern / picture fills, custom geometry, charts, tables
    and SmartArt show as labelled fallbacks. PowerPoint or LibreOffice will render the file
    correctly.
  </p>

  <div
    class="drop"
    class:dropping
    role="button"
    tabindex="0"
    aria-label="Drop a .pptx here"
    ondragover={(e) => {
      e.preventDefault();
      dropping = true;
    }}
    ondragleave={() => (dropping = false)}
    ondrop={onDrop}
  >
    <span class="drop-coord">▸</span>
    <span class="drop-text">{fileName}</span>
    <label class="drop-pick">
      <input
        type="file"
        accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        onchange={(e) => {
          const f = (e.currentTarget as HTMLInputElement).files?.[0];
          if (f) onFileChosen(f);
        }}
      />
      <span>Pick a file</span>
    </label>
    {#if lastBytes}
      <button class="drop-roundtrip" onclick={downloadRoundtrip}>
        Download round-trip
      </button>
    {/if}
  </div>

  <p class="status" class:busy aria-live="polite">{status}</p>

  {#if slideCount > 0}
    <div class="meta">
      <div class="cell">
        <span class="label">slides</span>
        <span class="value">{slideCount}</span>
      </div>
      <div class="cell">
        <span class="label">parts</span>
        <span class="value">{parts.length}</span>
      </div>
      <div class="cell">
        <span class="label">core / title</span>
        <span class="value">{coreTitle || '—'}</span>
      </div>
      <div class="cell">
        <span class="label">core / creator</span>
        <span class="value">{coreCreator || '—'}</span>
      </div>
      {#if summary}
        <div class="cell">
          <span class="label">theme</span>
          <span class="value">{summary.themeName ?? '—'}</span>
        </div>
        <div class="cell">
          <span class="label">masters · layouts · sections</span>
          <span class="value">{masterCount} · {summary.layoutCount} · {summary.sectionCount}</span>
        </div>
        <div class="cell">
          <span class="label">shapes (total)</span>
          <span class="value">{summary.totalShapes}</span>
        </div>
        <div class="cell">
          <span class="label">deck flags</span>
          <span class="value">
            {summary.hiddenSlideCount > 0 ? `${summary.hiddenSlideCount} hidden · ` : ''}{summary.hasCharts
              ? 'charts · '
              : ''}{summary.hasComments ? 'comments · ' : ''}{summary.hasAnimations
              ? 'animations'
              : ''}{!summary.hasCharts && !summary.hasComments && !summary.hasAnimations && summary.hiddenSlideCount === 0
              ? '—'
              : ''}
          </span>
        </div>
        {#if chartKindCounts && Object.values(chartKindCounts).some((n) => n > 0)}
          <div class="cell">
            <span class="label">chart kinds</span>
            <span class="value">
              {Object.entries(chartKindCounts)
                .filter(([, n]) => n > 0)
                .map(([k, n]) => `${n} ${k}`)
                .join(' · ')}
            </span>
          </div>
        {/if}
        {#if layoutTypeCounts && Object.keys(layoutTypeCounts).length > 0}
          <div class="cell">
            <span class="label">layout types in use</span>
            <span class="value">
              {Object.entries(layoutTypeCounts)
                .map(([k, n]) => `${n} ${k}`)
                .join(' · ')}
            </span>
          </div>
        {/if}
      {/if}
    </div>

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
            <a class="s-num" href={`#slide-${s.index}`} title="copy link to this slide">{String(s.index).padStart(2, '0')}</a>
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
            <span class="s-len">{s.textLength} chars · {s.shapeKinds.length} shapes</span>
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

    <h2>OPC parts</h2>
    <table class="parts">
      <thead>
        <tr><th>part</th><th>content-type</th><th>bytes</th></tr>
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
  {/if}
</section>

<style>
  .content {
    max-width: var(--max-content);
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
    font-family: var(--display);
    font-weight: 460;
    font-size: clamp(2rem, 4.6vw, 2.95rem);
    line-height: 1.05;
    letter-spacing: -0.026em;
    margin: 0 0 1rem;
    border: none;
    padding: 0;
    font-variation-settings: 'opsz' 144, 'SOFT' 30;
    max-width: 22ch;
  }

  h2 {
    font-family: var(--display);
    font-weight: 500;
    font-size: 1.45rem;
    letter-spacing: -0.015em;
    margin: 2.25rem 0 0.75rem;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid var(--border);
    font-variation-settings: 'opsz' 96, 'SOFT' 25;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.05rem;
    line-height: 1.55;
    max-width: 64ch;
    margin: 0 0 0.6rem;
  }

  .caveat {
    color: var(--fg-muted);
    font-size: 0.88rem;
    max-width: 64ch;
    margin: 0 0 1.75rem;
    line-height: 1.55;
  }

  .caveat strong {
    color: var(--fg-soft);
  }

  .drop {
    display: flex;
    align-items: center;
    gap: 1rem;
    padding: 1.1rem 1.25rem;
    border: 1px dashed var(--border-strong);
    border-radius: var(--radius);
    background: var(--bg-elev);
    transition:
      background 150ms ease,
      border-color 150ms ease;
    flex-wrap: wrap;
  }

  .drop.dropping {
    background: var(--accent-soft);
    border-color: var(--accent);
  }

  .drop-coord {
    color: var(--accent);
    font-family: var(--mono);
    font-weight: 600;
  }

  .drop-text {
    font-family: var(--mono);
    font-size: 0.92rem;
    color: var(--fg);
    flex: 1;
    min-width: 12ch;
  }

  .drop-pick {
    display: inline-flex;
    cursor: pointer;
  }

  .drop-pick input {
    display: none;
  }

  .drop-pick span {
    padding: 0.55rem 0.95rem;
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    background: var(--bg);
    color: var(--fg);
    font-family: var(--sans);
    font-size: 0.92rem;
  }

  .drop-pick span:hover {
    border-color: var(--border-strong);
  }

  .drop-roundtrip {
    padding: 0.55rem 0.95rem;
    border: 1px solid var(--accent);
    background: var(--accent);
    color: var(--bg);
    border-radius: var(--radius-sm);
    font-family: var(--sans);
    font-size: 0.92rem;
    font-weight: 540;
    cursor: pointer;
  }

  .drop-roundtrip:hover {
    background: var(--accent-hot);
    border-color: var(--accent-hot);
  }

  .status {
    margin: 1rem 0 0;
    font-family: var(--mono);
    font-size: 0.85rem;
    color: var(--fg-muted);
  }

  .status.busy {
    color: var(--accent);
  }

  .meta {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
    margin: 1.5rem 0 0;
    background: var(--bg-paper);
  }

  .cell {
    display: flex;
    flex-direction: column;
    gap: 0.2rem;
    padding: 0.9rem 1rem;
    border-right: 1px solid var(--border);
  }

  .cell:last-child {
    border-right: none;
  }

  .label {
    font-family: var(--mono);
    font-size: 10.5px;
    font-weight: 500;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--fg-muted);
  }

  .value {
    font-family: var(--display);
    font-size: 1.05rem;
    color: var(--fg);
    font-variation-settings: 'opsz' 32, 'SOFT' 25;
  }

  .slides {
    list-style: none;
    margin: 0;
    padding: 0;
  }

  .slides li {
    padding: 1.1rem 0 1.6rem;
    border-bottom: 1px solid var(--rule);
  }

  .s-head {
    display: flex;
    align-items: baseline;
    gap: 0.8rem;
    flex-wrap: wrap;
    margin-bottom: 0.55rem;
  }

  .s-num {
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--accent);
    font-weight: 500;
    text-decoration: none;
  }

  .s-num:hover {
    text-decoration: underline;
  }

  .s-title {
    font-family: var(--display);
    font-weight: 540;
    font-size: 1.05rem;
    flex: 1;
    min-width: 18ch;
  }

  .s-len {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--fg-muted);
  }

  .s-canvas {
    aspect-ratio: 16 / 9;
    background: #ffffff;
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    overflow: hidden;
    box-shadow: 0 4px 18px -10px rgba(0, 0, 0, 0.45);
  }

  .s-canvas :global(svg) {
    display: block;
    width: 100%;
    height: 100%;
  }

  .s-kinds {
    margin: 0.55rem 0 0;
    display: flex;
    flex-wrap: wrap;
    gap: 0.4rem;
  }

  .s-kinds code {
    font-size: 11px;
    padding: 0.1em 0.45em;
  }

  .section-divider {
    list-style: none;
    margin: 1.5rem 0 0.5rem;
    padding: 0.5rem 0;
    border-bottom: 1px dashed var(--border, #cbd5e1);
    font-family: var(--mono, monospace);
    font-size: 0.85rem;
    color: var(--muted, #4b5563);
  }

  .section-divider .section-name {
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }

  .s-badge {
    display: inline-block;
    padding: 0.05em 0.45em;
    font-size: 10px;
    font-family: var(--mono, monospace);
    color: var(--muted, #4b5563);
    background: var(--panel, #f1f5f9);
    border-radius: 3px;
    margin-left: 0.25em;
  }

  .s-badge-hidden {
    color: #92400e;
    background: #fef3c7;
  }

  .issues {
    list-style: none;
    margin: 0.5rem 0 1.5rem;
    padding: 0;
    display: flex;
    flex-direction: column;
    gap: 0.35rem;
    font-size: 0.88rem;
  }

  .issue {
    display: flex;
    gap: 0.5rem;
    padding: 0.35rem 0.55rem;
    border-radius: 4px;
    background: var(--panel, #f8fafc);
  }

  .issue-error {
    color: #b91c1c;
    background: #fee2e2;
  }

  .issue-warning {
    color: #92400e;
    background: #fef3c7;
  }

  .issue-sev {
    font-family: var(--mono, monospace);
    font-size: 11px;
    text-transform: uppercase;
    align-self: center;
  }

  .issue-msg {
    flex: 1;
  }

  .issue-part {
    font-family: var(--mono, monospace);
    font-size: 11px;
    opacity: 0.7;
  }

  .s-notes {
    margin: 0.55rem 0 0;
    font-size: 0.85rem;
    color: var(--muted, #4b5563);
  }

  .s-notes summary {
    cursor: pointer;
    user-select: none;
  }

  .s-notes pre {
    margin: 0.4rem 0 0;
    padding: 0.5rem 0.7rem;
    background: var(--panel, #f8fafc);
    border-radius: 4px;
    white-space: pre-wrap;
    word-break: break-word;
    font-family: inherit;
    font-size: inherit;
  }

  .parts {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.88rem;
    font-family: var(--mono);
  }

  .parts th,
  .parts td {
    text-align: left;
    border-bottom: 1px solid var(--rule);
    padding: 0.4rem 0.6rem;
    vertical-align: top;
  }

  .parts th {
    color: var(--fg-muted);
    font-weight: 500;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.1em;
  }

  .parts .ct {
    color: var(--fg-muted);
    font-size: 11.5px;
  }

  .parts .num {
    text-align: right;
    color: var(--fg-soft);
  }
</style>
