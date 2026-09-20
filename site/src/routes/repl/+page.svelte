<script lang="ts">
  import { onMount } from 'svelte';
  import { replaceState } from '$app/navigation';
  import * as kit from '@office-kit/pptx';
  import * as dsl from '@office-kit/pptx-dsl';
  import * as jsxRuntime from '@office-kit/pptx-dsl/jsx-runtime';
  import { renderSlideToSvg } from '@office-kit/pptx-preview';
  import { EditorState } from '@codemirror/state';
  import { EditorView, basicSetup } from 'codemirror';
  import { javascript } from '@codemirror/lang-javascript';
  import { oneDark } from '@codemirror/theme-one-dark';
  // The starter deck has to survive being typed into: nothing in it may return
  // null, or a half-edited string takes the whole deck down with a
  // null-dereference. (`findSlideLayout(pres, 'Title Slid')` did exactly that.)
  // It lives outside `src/` because it is a script body, not a module: every
  // library function is a free identifier, which `checkJs` would reject.
  // test/site-repl-starter.test.ts runs it against the library instead.
  import DEFAULT_CODE from '../../../repl/default-deck.js?raw';
  // The TSX starter is the DSL package's own example, so `pnpm typecheck`
  // there checks the very text the editor opens with.
  import DEFAULT_TSX from '../../../../packages/dsl/examples/review.tsx?raw';
  import { WRAPPER_LINES, asyncBody, evaluateTsx } from '$lib/repl/evaluate';

  type Mode = 'functions' | 'tsx';
  const MODES: ReadonlyArray<{ id: Mode; label: string }> = [
    { id: 'functions', label: 'Functions' },
    { id: 'tsx', label: 'TSX' },
  ];
  const STARTERS: Record<Mode, string> = { functions: DEFAULT_CODE, tsx: DEFAULT_TSX };
  // `?mode=tsx` opens the TSX tab, so the docs can link straight to it.
  const MODE_PARAM = 'mode';
  // What a TSX deck may import: the same three modules a project on disk has.
  const TSX_MODULES = {
    '@office-kit/pptx': kit,
    '@office-kit/pptx-dsl': dsl,
    '@office-kit/pptx-dsl/jsx-runtime': jsxRuntime,
  };

  let mode = $state<Mode>('functions');
  // Each tab keeps its own text, so switching back does not lose an edit.
  const drafts: Record<Mode, string> = { ...STARTERS };
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
            javascript({ typescript: true, jsx: true }),
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
                drafts[mode] = code;
              }
            }),
          ],
        }),
        parent: editorContainer,
      });
    }
    // Read here, not from `page.url`: the page is prerendered, where search
    // params are not available.
    if (new URLSearchParams(location.search).get(MODE_PARAM) === 'tsx') selectMode('tsx');
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
    drafts[mode] = next;
    if (view) {
      view.dispatch({
        changes: { from: 0, to: view.state.doc.length, insert: next },
      });
    }
  }

  function selectMode(next: Mode) {
    if (next === mode) return;
    mode = next;
    // A stale preview is kept while code is broken, but only a deck this tab
    // built: the other tab's deck under this tab's error would mislead.
    slides = [];
    bytes = null;
    setEditorText(drafts[next]);
  }

  function switchMode(next: Mode) {
    selectMode(next);
    const url = new URL(location.href);
    if (next === 'functions') url.searchParams.delete(MODE_PARAM);
    else url.searchParams.set(MODE_PARAM, next);
    replaceState(url, {});
  }

  let runTimer: ReturnType<typeof setTimeout> | null = null;
  $effect(() => {
    // Subscribe to `code` so each keystroke triggers a re-run, with a
    // 250 ms debounce so we don't recompile on every character.
    void code;
    void mode;
    if (!mounted) return;
    if (runTimer) clearTimeout(runTimer);
    runTimer = setTimeout(() => void run(), 250);
  });

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

  async function runFunctions(): Promise<kit.PresentationData> {
    const pres = kit.createPresentation();
    // Filter out the underscore-prefixed escape hatch and the
    // `VERSION` constant; expose everything else as a free function
    // parameter so the user can write `addSlide(...)` etc. directly.
    const entries = Object.entries(kit).filter(([k]) => !k.startsWith('_') && k !== 'VERSION');
    const names = entries.map((e) => e[0]);
    const values = entries.map((e) => e[1]);
    await new Function(...names, 'pres', asyncBody(code))(...values, pres);
    return pres;
  }

  // The default export comes out of evaluated text, so nothing vouches for its
  // shape until this check.
  const isDslNode = (value: unknown): value is dsl.Node =>
    typeof value === 'object' && value !== null && 'kind' in value;

  async function runTsx(): Promise<kit.PresentationData> {
    const root = await evaluateTsx(code, TSX_MODULES);
    if (!isDslNode(root)) throw new Error('The TSX file must default-export a Presentation.');
    return dsl.compile(root);
  }

  async function run() {
    busy = true;
    try {
      const pres = mode === 'tsx' ? await runTsx() : await runFunctions();
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
    setEditorText(STARTERS[mode]);
  }
</script>

<svelte:head>
  <title>REPL · @office-kit/pptx</title>
</svelte:head>

<section class="content">
  <header class="intro">
    <h1>REPL</h1>
    <p class="lede">
      Write code and the deck redraws as you type. In <strong>Functions</strong> every public
      function is already in scope, and <code>pres</code> is a new 16:9 deck from
      <code>createPresentation()</code>; the starter is a six-slide board deck driven by one data
      object. <strong>TSX</strong> runs a <code>@office-kit/pptx-dsl</code> file exactly as you
      would save it: imports at the top, a <code>&lt;Presentation&gt;</code> as the default export.
      Either way the download is the same bytes <code>savePresentation</code> writes in production.
    </p>
  </header>

  <div class="repl-grid">
    <div class="pane editor-pane">
      <div class="pane-head">
        <div class="modes" role="group" aria-label="Authoring style">
          {#each MODES as m (m.id)}
            <button type="button" aria-pressed={mode === m.id} onclick={() => switchMode(m.id)}>
              {m.label}
            </button>
          {/each}
        </div>
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
            The deck has no slides yet. Add one with
            {#if mode === 'tsx'}
              <code>&lt;Slide&gt;</code> inside <code>&lt;Presentation&gt;</code>.
            {:else}
              <code>addSlide(pres, &#123; layout &#125;)</code>.
            {/if}
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

  /* Two pressed-state buttons sharing one border, so they read as one control. */
  .modes {
    display: flex;
    border: 1px solid var(--night-line);
    border-radius: var(--radius-sm);
    overflow: hidden;
  }

  .modes button {
    height: 30px;
    padding: 0 0.8rem;
    border: none;
    background: transparent;
    color: var(--night-ink);
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 550;
    cursor: pointer;
    opacity: 0.7;
  }

  .modes button + button {
    border-left: 1px solid var(--night-line);
  }

  .modes button:hover {
    opacity: 1;
  }

  .modes button[aria-pressed='true'] {
    background: var(--night-2);
    opacity: 1;
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
