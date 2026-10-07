<script lang="ts">
  // Reading View for an editor without a host viewer. Mac PowerPoint opens it
  // as a separate "Slide Show - [document]" window: a 32 pt title bar and the
  // slide letterboxed on black below it. Here it fills the editor's window.
  // Click, Space, → and ↓ advance; ← and ↑ go back; Esc returns to the editing
  // view on the slide shown last. Hidden slides are skipped, as in a show.
  import { onMount, tick } from 'svelte';
  import { renderSlideToSvg } from '@office-kit/pptx-preview';
  import { getSlideSize, isSlideHidden } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const shown = $derived.by(() => { doc.version; return doc.slides.map((slide, index) => ({ slide, index })).filter((item) => !isSlideHidden(item.slide)); });
  let position = $state(0);
  const size = $derived.by(() => { doc.version; return getSlideSize(doc.pres) ?? { width: 12192000, height: 6858000 }; });
  const current = $derived(shown[position]);
  const svg = $derived.by(() => {
    if (!current) return '';
    try {
      return renderSlideToSvg(doc.pres, current.slide);
    } catch {
      return '';
    }
  });
  let root = $state<HTMLDivElement>();

  onMount(() => {
    const start = shown.findIndex((item) => item.index >= doc.selection.slideIndex);
    position = Math.max(0, start);
    void tick().then(() => root?.focus());
  });
  function go(delta: number) {
    position = Math.max(0, Math.min(shown.length - 1, position + delta));
  }
  function close() {
    if (current) doc.selectSlide(current.index);
    editor.readingView = false;
  }
  function keys(event: KeyboardEvent) {
    if (event.key === 'Escape') close();
    else if ([' ', 'ArrowRight', 'ArrowDown', 'PageDown', 'Enter', 'n'].includes(event.key)) go(1);
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(event.key)) go(-1);
    else if (event.key === 'Home') position = 0;
    else if (event.key === 'End') position = Math.max(0, shown.length - 1);
    else return;
    event.preventDefault();
    event.stopPropagation();
  }
</script>

<div class="reading-view" bind:this={root} role="dialog" aria-modal="true" aria-label={t('Reading View')} tabindex="-1" onkeydown={keys}>
  <header class="title-bar">
    <button class="close" aria-label={t('Close')} title={t('Close')} onclick={close}></button>
    <span class="title">{t('Slide Show')} - [{doc.fileName.replace(/\.pptx$/i, '')}]</span>
  </header>
  <!-- svelte-ignore a11y_click_events_have_key_events -->
  <div class="stage" role="presentation" onclick={() => go(1)}>
    <div class="slide" role="img" aria-label="{t('Slide')} {(current?.index ?? 0) + 1}" style:aspect-ratio="{size.width} / {size.height}" style:--ratio={size.width / size.height}>{@html svg}</div>
  </div>
</div>

<style>
  .reading-view { position: fixed; inset: 0; z-index: 900; display: grid; grid-template-rows: 32px minmax(0, 1fr); background: #000; outline: none; }
  .title-bar { position: relative; display: flex; align-items: center; justify-content: center; background: var(--ok-ribbon); color: var(--ok-text-2); font-size: 13px; font-weight: 600; border-bottom: 1px solid var(--ok-border); }
  .title { position: absolute; left: 82px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: calc(100% - 100px); }
  /* The window's close button: a 12 pt circle 8 pt from the corner. */
  .close { position: absolute; left: 10px; width: 12px; height: 12px; padding: 0; border: 0; border-radius: 50%; background: #ff5f57; cursor: pointer; }
  .stage { display: grid; place-items: center; min-height: 0; overflow: hidden; cursor: pointer; container-type: size; }
  .slide { width: min(100cqw, calc(100cqh * var(--ratio))); max-width: 100cqw; max-height: 100cqh; background: #fff; }
  .slide :global(svg) { display: block; width: 100%; height: 100%; }
</style>
