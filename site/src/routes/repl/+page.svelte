<script lang="ts">
  import { base } from '$app/paths';
  import { onMount } from 'svelte';
  import * as kit from '@office-kit/pptx';
  import { renderSlideToSvg } from '@office-kit/pptx-preview';
  import { EditorState } from '@codemirror/state';
  import { EditorView, basicSetup } from 'codemirror';
  import { javascript } from '@codemirror/lang-javascript';
  import { oneDark } from '@codemirror/theme-one-dark';

  // Default snippet — touches the most-used corners of the API so a
  // brand-new visitor sees a non-trivial deck immediately.
  const DEFAULT_CODE = `// @office-kit/pptx is exposed as global functions — no imports needed.
// Edit anything; the preview updates as you type.
// \`pres\` is a fresh PresentationData loaded from a blank template.

const titleLayout = findSlideLayout(pres, 'Title Slide');
const contentLayout = findSlideLayout(pres, 'Title and Content');

// Slide 1 — title
const cover = addSlide(pres, { layout: titleLayout });
const t = findSlidePlaceholder(cover, 'ctrTitle') ?? findSlidePlaceholder(cover, 'title');
if (t) setShapeText(t, '@office-kit/pptx REPL');
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
            // One Dark supplies the syntax colours; the surfaces are ours so
            // the editor matches every other code panel on the site.
            EditorView.theme(
              {
                '&': { height: '100%', fontSize: '13px', backgroundColor: 'var(--night)' },
                '.cm-gutters': { backgroundColor: 'var(--night)', border: 'none' },
                '.cm-activeLine, .cm-activeLineGutter': { backgroundColor: 'var(--night-2)' },
                '.cm-scroller': { fontFamily: 'var(--mono)', overflow: 'auto' },
              },
              { dark: true },
            ),
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
    a.download = 'office-kit-pptx-repl.pptx';
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
  <title>REPL · @office-kit/pptx</title>
</svelte:head>

<section class="content">
  <header class="intro">
    <h1>REPL</h1>
    <p class="lede">
      Write code and the deck redraws as you type. Every public function is already in
      scope, and <code>pres</code> is a fresh presentation loaded from a blank template. The
      download is the same bytes <code>savePresentation</code> writes in production.
    </p>
  </header>

  <div class="repl-grid">
    <div class="pane editor-pane">
      <div class="pane-head">
        <h2>Code</h2>
        <div class="pane-actions">
          <button type="button" onclick={resetCode}>Reset</button>
          <button type="button" onclick={copyCode}>Copy</button>
        </div>
      </div>
      <div class="editor" bind:this={editorContainer}></div>
      {#if error}
        <pre class="error" role="alert">{error}</pre>
      {/if}
    </div>

    <div class="pane preview-pane">
      <div class="pane-head">
        <h2>Preview</h2>
        <div class="pane-actions">
          <span class="busy" class:visible={busy} aria-live="polite">{busy ? 'Building…' : ''}</span>
          <button type="button" class="strong" onclick={download} disabled={!bytes}>
            Download .pptx
          </button>
        </div>
      </div>
      <div class="slides">
        {#each slides as s (s.index)}
          <article class="slide-card">
            <h3>Slide {s.index}{s.title ? `: ${s.title}` : ''}</h3>
            <div class="slide-canvas">{@html s.svg}</div>
          </article>
        {/each}
        {#if slides.length === 0 && !error}
          <p class="empty">
            The deck has no slides yet. Add one with <code>addSlide(pres, &#123; layout &#125;)</code>.
          </p>
        {/if}
      </div>
    </div>
  </div>
</section>

<style>
  .content {
    max-width: 1440px;
    margin: 0 auto;
    padding: 2.25rem var(--gutter) 3rem;
  }

  .intro h1 {
    margin-bottom: 0.6rem;
  }

  .lede {
    max-width: 78ch;
    margin: 0 0 1.5rem;
    color: var(--ink-2);
  }

  .repl-grid {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
    gap: 1rem;
    align-items: start;
  }

  .pane {
    display: flex;
    flex-direction: column;
    min-width: 0;
    height: calc(100vh - var(--header-h) - 2rem);
    min-height: 480px;
    border: 1px solid var(--line-strong);
    border-radius: var(--radius);
    overflow: hidden;
  }

  .editor-pane {
    position: sticky;
    top: calc(var(--header-h) + 1rem);
    background: var(--night);
    border-color: var(--night-line);
  }

  .pane-head {
    flex: none;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.75rem;
    min-height: 48px;
    padding: 0 0.6rem 0 1rem;
    border-bottom: 1px solid var(--line);
    background: var(--wash);
  }

  .editor-pane .pane-head {
    background: var(--night);
    border-bottom-color: var(--night-line);
    color: var(--night-ink);
  }

  .pane-head h2 {
    margin: 0;
    font-family: var(--sans);
    font-size: 0.9rem;
    font-weight: 600;
    letter-spacing: 0;
    color: inherit;
  }

  .pane-actions {
    display: flex;
    align-items: center;
    gap: 0.4rem;
  }

  .pane-actions button {
    height: 32px;
    padding: 0 0.75rem;
    border: 1px solid var(--line-strong);
    border-radius: var(--radius-sm);
    background: var(--paper);
    color: var(--ink);
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 550;
    cursor: pointer;
  }

  .pane-actions button:hover:not(:disabled) {
    border-color: var(--ink-3);
  }

  .pane-actions button:disabled {
    opacity: 0.5;
    cursor: not-allowed;
  }

  .editor-pane .pane-actions button {
    background: var(--night-2);
    border-color: var(--night-line);
    color: var(--night-ink);
  }

  .pane-actions button.strong {
    background: var(--ink);
    border-color: var(--ink);
    color: var(--paper);
  }

  .pane-actions button.strong:hover:not(:disabled) {
    background: var(--accent);
    border-color: var(--accent);
    color: var(--on-accent);
  }

  .busy {
    color: var(--ink-3);
    font-size: 0.82rem;
  }

  .editor {
    flex: 1;
    min-height: 0;
    overflow: hidden;
  }

  .editor :global(.cm-editor) {
    height: 100%;
    font-size: 13px;
  }

  .editor :global(.cm-scroller) {
    font-family: var(--mono);
    line-height: 1.6;
  }

  .error {
    flex: none;
    max-height: 35%;
    margin: 0;
    border: none;
    border-top: 1px solid #7a2a1a;
    border-radius: 0;
    background: #2a1410;
    color: #ffb4a1;
    font-size: 0.8rem;
    white-space: pre-wrap;
    overflow: auto;
  }

  .slides {
    flex: 1;
    min-height: 0;
    overflow-y: auto;
    padding: 1rem;
    background: var(--wash);
  }

  .slide-card + .slide-card {
    margin-top: 1.25rem;
  }

  .slide-card h3 {
    margin: 0 0 0.4rem;
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 550;
    letter-spacing: 0;
    color: var(--ink-2);
  }

  /* A slide's page colour belongs to the deck, so it stays white in dark mode. */
  .slide-canvas {
    border: 1px solid var(--line-strong);
    border-radius: 3px;
    background: #fff;
    overflow: hidden;
  }

  .slide-canvas :global(svg) {
    display: block;
    width: 100%;
    height: auto;
  }

  .empty {
    margin: 2rem 0;
    color: var(--ink-2);
    text-align: center;
  }

  @media (max-width: 900px) {
    .repl-grid {
      grid-template-columns: 1fr;
    }

    .pane {
      height: auto;
      min-height: 0;
    }

    .editor-pane {
      position: static;
    }

    .editor {
      height: 46vh;
      min-height: 280px;
      flex: none;
    }

    .slides {
      overflow: visible;
    }
  }
</style>
