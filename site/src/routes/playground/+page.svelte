<script lang="ts">
  import { base } from '$app/paths';

  // pptx-kit is consumed via the source-tree path alias set in
  // svelte.config.js (`pptx-kit` → `../src/index.ts`). The playground
  // exercises the real surface the same way the rest of the docs site
  // does — type errors here break the build.
  import {
    getCoreProperties,
    getShapeKind,
    getSlideShapes,
    getSlideTitle,
    getSlides,
    getSlideTextLength,
    listPackageParts,
    loadPresentation,
    savePresentation,
  } from 'pptx-kit';
  import { renderSlideSvg } from '$lib/playground/render-slide';

  type SlideSnapshot = {
    index: number;
    title: string;
    textLength: number;
    shapeKinds: string[];
    svg: string;
  };

  type PackagePart = { name: string; contentType: string; byteLength: number };

  let fileName = $state<string>('(drop a .pptx here, or pick one)');
  let status = $state<string>('Ready.');
  let busy = $state<boolean>(false);
  let dropping = $state<boolean>(false);
  let slideCount = $state<number>(0);
  let coreTitle = $state<string>('');
  let coreCreator = $state<string>('');
  let slides = $state<SlideSnapshot[]>([]);
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

      const list = getSlides(pres);
      slideCount = list.length;
      slides = list.map((slide, i) => ({
        index: i + 1,
        title: getSlideTitle(slide) ?? '',
        textLength: getSlideTextLength(slide),
        shapeKinds: getSlideShapes(slide).map((sh) => getShapeKind(sh)),
        svg: renderSlideSvg(pres, slide),
      }));

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
    </div>

    <h2>Slides</h2>
    <ol class="slides">
      {#each slides as s (s.index)}
        <li>
          <div class="s-head">
            <span class="s-num">{String(s.index).padStart(2, '0')}</span>
            <span class="s-title">{s.title || '(untitled)'}</span>
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
