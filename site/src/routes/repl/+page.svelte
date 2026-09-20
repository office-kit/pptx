<script lang="ts">
  import { onMount } from 'svelte';
  import * as kit from '@office-kit/pptx';
  import { renderSlideToSvg } from '@office-kit/pptx-preview';
  import { EditorState } from '@codemirror/state';
  import { EditorView, basicSetup } from 'codemirror';
  import { javascript } from '@codemirror/lang-javascript';
  import { oneDark } from '@codemirror/theme-one-dark';

  // The starter deck. It has to survive being typed into: nothing here can
  // return null, so a half-edited string never takes the whole deck down with
  // a null-dereference. (`findSlideLayout(pres, 'Title Slid')` did exactly that.)
  const DEFAULT_CODE = `// Every @office-kit/pptx function is in scope. No imports needed.
// \`pres\` is a new 16:9 deck from createPresentation().

const INK = '#15171C';
const MUTED = '#5B616E';
const ACCENT = '#E5481F';
const WASH = '#F3F4F7';

// Adds a text box and formats it in one go.
function text(slide, str, x, y, w, h, format) {
  const box = addSlideTextBox(slide, {
    x: inches(x), y: inches(y), w: inches(w), h: inches(h), text: str,
  });
  setShapeTextFormat(box, format);
  return box;
}

// A filled rectangle with no outline.
function block(slide, preset, x, y, w, h, color) {
  const shape = addSlideShape(slide, {
    preset, x: inches(x), y: inches(y), w: inches(w), h: inches(h),
  });
  setShapeFill(shape, color);
  setShapeNoStroke(shape);
  return shape;
}

// Slide 1: cover
const cover = addBlankSlide(pres);
setSlideBackground(cover, INK);
block(cover, 'rect', 0.9, 2.55, 0.14, 1.75, ACCENT);
text(cover, 'Q3 business review', 1.25, 2.4, 10, 1.1, { size: 48, bold: true, color: '#FFFFFF' });
text(cover, 'Revenue, margin, and what we do next', 1.25, 3.55, 10, 0.6, { size: 22, color: '#B4B9C4' });
text(cover, 'Finance team, October 2026', 1.25, 6.4, 10, 0.4, { size: 14, color: '#8A909C' });

// Slide 2: three numbers and the chart behind them
const numbers = addBlankSlide(pres);
text(numbers, 'Revenue grew 2.5x in four quarters', 0.9, 0.55, 11.5, 0.8, { size: 30, bold: true, color: INK });

const kpis = [
  ['$300k', 'Q4 revenue'],
  ['47%', 'Gross margin'],
  ['+25%', 'Quarter on quarter'],
];
kpis.forEach(([value, label], i) => {
  const y = 1.7 + i * 1.75;
  block(numbers, 'roundRect', 0.9, y, 3.4, 1.5, WASH);
  text(numbers, value, 1.15, y + 0.15, 3, 0.75, { size: 34, bold: true, color: ACCENT });
  text(numbers, label, 1.15, y + 0.9, 3, 0.4, { size: 14, color: MUTED });
});

addSlideChart(numbers, {
  x: inches(4.7), y: inches(1.6), w: inches(7.8), h: inches(5.3),
  spec: {
    kind: 'column',
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Revenue', values: [120, 180, 240, 300], color: ACCENT },
      { name: 'Cost', values: [80, 90, 130, 160], color: '#C9CDD6' },
    ],
    legend: { position: 'b' },
    valueAxisMajorGridlines: true,
    valueAxisMajorGridlineColor: '#E2E4E9',
    gapWidthPct: 80,
  },
});

// Slide 3: a table with a styled header row
const plan = addBlankSlide(pres);
text(plan, 'What we do next', 0.9, 0.55, 11.5, 0.8, { size: 30, bold: true, color: INK });

const rows = [
  ['Owner', 'Action', 'Due'],
  ['Aiko', 'Renegotiate the hosting contract', 'Nov 15'],
  ['Ben', 'Ship annual billing', 'Dec 1'],
  ['Chloe', 'Hire two support engineers', 'Jan 10'],
];
const table = addSlideTable(plan, {
  x: inches(0.9), y: inches(1.7), w: inches(11.5), h: inches(3.2),
  rows,
  colWidths: [inches(2.2), inches(7), inches(2.3)],
});
rows.forEach((row, r) => {
  row.forEach((_, c) => {
    const cell = getTableCell(table, r, c);
    setTableCellFill(cell, r === 0 ? INK : r % 2 === 0 ? WASH : '#FFFFFF');
    setTableCellAnchor(cell, 'center');
    setTableCellTextFormat(cell, {
      size: 16, bold: r === 0, color: r === 0 ? '#FFFFFF' : INK,
    });
  });
});
`;

  let code = $state<string>(DEFAULT_CODE);
  let error = $state<string>('');
  let slides = $state<Array<{ index: number; svg: string; title: string }>>([]);
  let bytes = $state<Uint8Array | null>(null);
  let busy = $state<boolean>(false);
  // The first run waits for mount so prerendering never evaluates the code.
  let mounted = $state<boolean>(false);

  // CodeMirror instance — created on mount, replaces the
  // <textarea> from the previous version. We mirror its document
  // into the `code` state on every change so the existing $effect
  // / debounce / runner logic stays untouched.
  let editorContainer: HTMLDivElement | undefined = $state();
  let view: EditorView | null = null;

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
    mounted = true;
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
    if (!mounted) return;
    if (runTimer) clearTimeout(runTimer);
    runTimer = setTimeout(() => void run(), 250);
  });

  // `new Function` wraps the source in a two-line header, and we add two more
  // lines before the user's code, so a stack frame's line is off by four.
  const WRAPPER_LINES = 4;

  // A minified stack trace is noise to someone writing eight lines of code.
  // Show the message, the line it came from, and a hint for the one mistake
  // that is otherwise baffling: a `find…` that matched nothing.
  function describeError(err: unknown): string {
    if (!(err instanceof Error)) return String(err);
    const frame = /<anonymous>:(\d+):\d+/.exec(err.stack ?? '');
    const line = frame ? Number(frame[1]) - WRAPPER_LINES : 0;
    const where = line > 0 ? `Line ${line}: ` : '';
    const hint = /of (null|undefined)/.test(err.message)
      ? '\nSomething on that line is null. Functions named find… return null when nothing matches, so check the name or type you passed.'
      : '';
    return `${where}${err.message}${hint}`;
  }

  async function run() {
    busy = true;
    try {
      const pres = kit.createPresentation();
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
      await fn(...values, pres);

      const list = kit.getSlides(pres);
      slides = list.map((slide, i) => ({
        index: i + 1,
        title: kit.getSlideTitle(slide) ?? '',
        svg: renderSlideToSvg(pres, slide),
      }));
      bytes = await kit.savePresentation(pres);
      error = '';
    } catch (err) {
      // Code is broken for most keystrokes while someone types. Keep the last
      // deck that built on screen (dimmed) instead of blanking the preview.
      error = describeError(err);
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
      scope, and <code>pres</code> is a new 16:9 deck from <code>createPresentation()</code>. The
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
      <div class="slides" class:stale={error !== ''}>
        {#each slides as s (s.index)}
          <article class="slide-card">
            <h3>Slide {s.index}{s.title ? `: ${s.title}` : ''}</h3>
            <div class="slide-canvas">{@html s.svg}</div>
          </article>
        {/each}
        {#if slides.length === 0 && !error && !busy}
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

  .slides.stale .slide-card {
    opacity: 0.45;
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
