<script lang="ts">
  import { base } from '$app/paths';
  import { onMount } from 'svelte';
  import * as kit from 'pptx-kit';
  import { renderSlideToSvg } from '@pptx-kit/preview';
  import { EditorState } from '@codemirror/state';
  import { EditorView, basicSetup } from 'codemirror';
  import { javascript } from '@codemirror/lang-javascript';
  import { oneDark } from '@codemirror/theme-one-dark';

  // Default snippet — touches the most-used corners of the API so a
  // brand-new visitor sees a non-trivial deck immediately.
  const DEFAULT_CODE = `// pptx-kit is exposed as global functions — no imports needed.
// Edit anything; the preview updates as you type.
// \`pres\` is a fresh PresentationData loaded from a blank template.

const titleLayout = findSlideLayout(pres, 'Title Slide');
const contentLayout = findSlideLayout(pres, 'Title and Content');

// Slide 1 — title
const cover = addSlide(pres, { layout: titleLayout });
const t = findSlidePlaceholder(cover, 'ctrTitle') ?? findSlidePlaceholder(cover, 'title');
if (t) setShapeText(t, 'pptx-kit REPL');
const sub = findSlidePlaceholder(cover, 'subTitle');
if (sub) setShapeText(sub, 'Edit the code on the left.');

// Slide 2 — shapes + chart
const s2 = addSlide(pres, { layout: contentLayout });
const titleSlot = findSlidePlaceholder(s2, 'title');
if (titleSlot) setShapeText(titleSlot, 'Quarterly numbers');

addSlideChart(s2, {
  x: inches(0.7), y: inches(1.5),
  w: inches(8.5), h: inches(4.5),
  spec: {
    kind: 'column',
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Revenue', values: [120, 180, 240, 300] },
      { name: 'Cost',    values: [80,  90,  130, 160] },
    ],
    title: 'FY26',
  },
});

// Slide 3 — table + shape
const s3 = addSlide(pres, { layout: contentLayout });
const s3Title = findSlidePlaceholder(s3, 'title');
if (s3Title) setShapeText(s3Title, 'Action items');

addSlideTable(s3, {
  x: inches(0.7), y: inches(1.5),
  w: inches(5), h: inches(2.5),
  rows: [
    ['Owner', 'Task'],
    ['A',     'Draft spec'],
    ['B',     'Review SLA'],
  ],
  firstRow: true,
});

const star = addSlideShape(s3, {
  preset: 'star5',
  x: inches(6.5), y: inches(1.5),
  w: inches(2.5), h: inches(2.5),
  text: 'GO',
});
setShapeFill(star, '#FFD966');
`;

  let code = $state<string>(DEFAULT_CODE);
  let error = $state<string>('');
  let slides = $state<Array<{ index: number; svg: string; title: string }>>([]);
  let bytes = $state<Uint8Array | null>(null);
  let busy = $state<boolean>(false);
  let blankBytes: Uint8Array | null = null;

  // CodeMirror instance — created on mount, replaces the
  // <textarea> from the previous version. We mirror its document
  // into the `code` state on every change so the existing $effect
  // / debounce / runner logic stays untouched.
  let editorContainer: HTMLDivElement | undefined = $state();
  let view: EditorView | null = null;

  // onMount must return its cleanup synchronously, but loading the
  // blank template + first compile are async — kick those off in a
  // detached IIFE and return the editor cleanup synchronously.
  onMount(() => {
    if (editorContainer) {
      view = new EditorView({
        state: EditorState.create({
          doc: code,
          extensions: [
            basicSetup,
            javascript({ typescript: true }),
            oneDark,
            EditorView.theme({
              '&': { height: '100%', fontSize: '13px' },
              '.cm-scroller': { fontFamily: "var(--mono)", overflow: 'auto' },
            }),
            EditorView.updateListener.of((update) => {
              if (update.docChanged) {
                code = update.state.doc.toString();
              }
            }),
          ],
        }),
        parent: editorContainer,
      });
    }
    void (async () => {
      try {
        const res = await fetch(`${base}/blank.pptx`);
        blankBytes = new Uint8Array(await res.arrayBuffer());
        await run();
      } catch (err) {
        error = `Failed to load blank template: ${err instanceof Error ? err.message : String(err)}`;
      }
    })();
    return () => {
      view?.destroy();
      view = null;
    };
  });

  // External setters (Reset button) push a fresh document into
  // CodeMirror via a transaction. Without this the in-state `code`
  // would update but the visible editor would be stale.
  function setEditorText(next: string) {
    code = next;
    if (view) {
      view.dispatch({
        changes: { from: 0, to: view.state.doc.length, insert: next },
      });
    }
  }

  let runTimer: ReturnType<typeof setTimeout> | null = null;
  $effect(() => {
    // Subscribe to `code` so each keystroke triggers a re-run, with a
    // 250 ms debounce so we don't recompile on every character.
    void code;
    if (!blankBytes) return;
    if (runTimer) clearTimeout(runTimer);
    runTimer = setTimeout(() => void run(), 250);
  });

  async function run() {
    if (!blankBytes) return;
    busy = true;
    error = '';
    try {
      const pres = await kit.loadPresentation(blankBytes);
      // Filter out the underscore-prefixed escape hatch and the
      // `VERSION` constant; expose everything else as a free function
      // parameter so the user can write `addSlide(...)` etc. directly.
      const entries = Object.entries(kit).filter(
        ([k]) => !k.startsWith('_') && k !== 'VERSION',
      );
      const names = entries.map((e) => e[0]);
      const values = entries.map((e) => e[1]);
      // Wrap user code in an async function so `await` is allowed.
      const fn = new Function(
        ...names,
        'pres',
        `'use strict';\nreturn (async () => {\n${code}\n})();`,
      );
      // biome-ignore lint/suspicious/noExplicitAny: dynamic eval surface
      await (fn as any)(...values, pres);

      const list = kit.getSlides(pres);
      slides = list.map((slide, i) => ({
        index: i + 1,
        title: kit.getSlideTitle(slide) ?? '',
        svg: renderSlideToSvg(pres, slide),
      }));
      bytes = await kit.savePresentation(pres);
    } catch (err) {
      error = err instanceof Error ? err.stack ?? err.message : String(err);
      slides = [];
      bytes = null;
    } finally {
      busy = false;
    }
  }

  function download() {
    if (!bytes) return;
    const blob = new Blob([bytes as BlobPart], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'pptx-kit-repl.pptx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function copyCode() {
    if (typeof navigator !== 'undefined' && navigator.clipboard) {
      void navigator.clipboard.writeText(code);
    }
  }

  function resetCode() {
    setEditorText(DEFAULT_CODE);
  }
</script>

<svelte:head>
  <title>REPL · pptx-kit</title>
</svelte:head>

<section class="content">
  <p class="eyebrow">§ 04 · REPL</p>
  <h1>Write code, see the deck.</h1>
  <p class="lede">
    A live editor for <code>pptx-kit</code>. Every public free-function
    export is in scope (no imports needed), <code>pres</code> is a
    fresh <code>PresentationData</code> loaded from a blank template,
    and the preview re-renders on every keystroke. Hit <kbd>Download</kbd>
    to get the actual <code>.pptx</code> bytes — the same path
    <code>savePresentation</code> writes in production.
  </p>

  <div class="repl-grid">
    <div class="editor-pane">
      <div class="pane-head">
        <span class="pane-label">code</span>
        <div class="pane-actions">
          <button type="button" onclick={resetCode}>Reset</button>
          <button type="button" onclick={copyCode}>Copy</button>
        </div>
      </div>
      <div class="editor" bind:this={editorContainer}></div>
      {#if error}
        <pre class="error">{error}</pre>
      {/if}
    </div>

    <div class="preview-pane">
      <div class="pane-head">
        <span class="pane-label">preview</span>
        <div class="pane-actions">
          <span class="busy" class:visible={busy}>compiling…</span>
          <button type="button" onclick={download} disabled={!bytes}>Download .pptx</button>
        </div>
      </div>
      <div class="slides">
        {#each slides as s (s.index)}
          <article class="slide-card">
            <header class="slide-card-head">
              <span class="slide-num">{String(s.index).padStart(2, '0')}</span>
              <span class="slide-title">{s.title || '(untitled)'}</span>
            </header>
            <div class="slide-canvas">{@html s.svg}</div>
          </article>
        {/each}
        {#if slides.length === 0 && !error}
          <p class="empty">no slides yet — try calling <code>addSlide(pres, &#123; layout &#125;)</code></p>
        {/if}
      </div>
    </div>
  </div>
</section>

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
    font-family: var(--display);
    font-weight: 460;
    font-size: clamp(2rem, 4.6vw, 2.95rem);
    line-height: 1.05;
    letter-spacing: -0.026em;
    margin: 0 0 1rem;
    border: none;
    padding: 0;
    font-variation-settings: 'opsz' 144, 'SOFT' 30;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.05rem;
    line-height: 1.55;
    max-width: 64ch;
    margin: 0 0 1.75rem;
  }

  .lede code,
  .lede kbd {
    font-family: var(--mono);
    font-size: 0.9em;
  }

  .lede kbd {
    background: var(--bg-elev);
    border: 1px solid var(--border);
    border-radius: 3px;
    padding: 0.05em 0.4em;
    color: var(--fg);
  }

  .repl-grid {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
    gap: 1.25rem;
  }

  @media (max-width: 1000px) {
    .repl-grid {
      grid-template-columns: 1fr;
    }
  }

  .editor-pane,
  .preview-pane {
    display: flex;
    flex-direction: column;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    background: var(--bg-elev);
    overflow: hidden;
    min-height: 60vh;
  }

  .pane-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0.55rem 0.85rem;
    border-bottom: 1px solid var(--border);
    background: var(--bg-soft);
  }

  .pane-label {
    font-family: var(--mono);
    font-size: 10.5px;
    font-weight: 500;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--fg-muted);
  }

  .pane-actions {
    display: flex;
    gap: 0.5rem;
    align-items: center;
  }

  .pane-actions button {
    padding: 0.3rem 0.7rem;
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    background: var(--bg);
    color: var(--fg);
    font-family: var(--sans);
    font-size: 0.82rem;
    cursor: pointer;
  }

  .pane-actions button:hover:not(:disabled) {
    border-color: var(--border-strong);
  }

  .pane-actions button:disabled {
    opacity: 0.4;
    cursor: not-allowed;
  }

  .busy {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--accent);
    opacity: 0;
    transition: opacity 80ms ease;
  }

  .busy.visible {
    opacity: 1;
  }

  .editor {
    flex: 1;
    min-height: 50vh;
    width: 100%;
    overflow: hidden;
    background: #282c34; /* matches CodeMirror one-dark base */
  }

  /* CodeMirror lives inside .editor; let it claim full height and
   * stretch the gutter to the panel's background colour. */
  .editor :global(.cm-editor) {
    height: 100%;
  }

  .editor :global(.cm-editor.cm-focused) {
    outline: none;
  }

  .editor :global(.cm-gutters) {
    background: #21252b;
    border-right-color: rgba(255, 255, 255, 0.06);
  }

  .error {
    margin: 0;
    padding: 0.75rem 1rem;
    border-top: 1px solid var(--accent);
    background: rgba(232, 80, 28, 0.08);
    color: #fca5a5;
    font-family: var(--mono);
    font-size: 11.5px;
    white-space: pre-wrap;
    max-height: 14rem;
    overflow: auto;
  }

  .slides {
    flex: 1;
    overflow: auto;
    padding: 0.85rem;
    display: flex;
    flex-direction: column;
    gap: 0.85rem;
    background: var(--bg);
  }

  .slide-card {
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    overflow: hidden;
    background: var(--bg-elev);
  }

  .slide-card-head {
    display: flex;
    align-items: baseline;
    gap: 0.6rem;
    padding: 0.4rem 0.7rem;
    border-bottom: 1px solid var(--border);
    background: var(--bg-soft);
  }

  .slide-num {
    font-family: var(--mono);
    font-size: 10.5px;
    color: var(--accent);
    font-weight: 500;
  }

  .slide-title {
    font-family: var(--display);
    font-weight: 540;
    font-size: 0.95rem;
    color: var(--fg);
  }

  .slide-canvas {
    aspect-ratio: 16 / 9;
    background: #ffffff;
  }

  .slide-canvas :global(svg) {
    display: block;
    width: 100%;
    height: 100%;
  }

  .empty {
    color: var(--fg-muted);
    font-family: var(--mono);
    font-size: 12px;
    padding: 1rem;
  }
</style>
