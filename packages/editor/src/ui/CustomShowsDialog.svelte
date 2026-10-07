<script lang="ts">
  import { onMount } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  let selectedId = $state<number | null>(null);
  let name = $state('');
  let slideIndices = $state<number[]>([]);
  let error = $state('');
  let dialog: HTMLDialogElement;

  const selectedShow = $derived(doc.customShows.find((show) => show.id === selectedId) ?? null);
  const slideIndexByObject = $derived(new Map(doc.slides.map((slide, index) => [slide, index] as const)));

  onMount(() => dialog.showModal());

  function edit(id: number): void {
    const show = doc.customShows.find((candidate) => candidate.id === id);
    if (!show) return;
    selectedId = id;
    name = show.name;
    slideIndices = show.slides
      .map((slide) => slideIndexByObject.get(slide))
      .filter((index): index is number => index !== undefined);
  }

  function newShow(): void {
    selectedId = null;
    name = '';
    slideIndices = [];
    error = '';
  }

  function runUpdate(action: () => void): void {
    error = '';
    try {
      action();
    } catch (cause) {
      error = `${t('Custom show update failed')}: ${cause instanceof Error ? cause.message : String(cause)}`;
    }
  }

  function addSlide(index: number): void { slideIndices = [...slideIndices, index]; }
  function removeSlide(sequenceIndex: number): void {
    slideIndices = slideIndices.filter((_, index) => index !== sequenceIndex);
  }
  function moveSlide(sequenceIndex: number, direction: -1 | 1): void {
    const next = sequenceIndex + direction;
    if (next < 0 || next >= slideIndices.length) return;
    const nextIndices = [...slideIndices];
    [nextIndices[sequenceIndex], nextIndices[next]] = [nextIndices[next]!, nextIndices[sequenceIndex]!];
    slideIndices = nextIndices;
  }

  function save(): void {
    runUpdate(() => {
      if (selectedId === null) editor.createCustomShow(name, slideIndices);
      else editor.updateCustomShow(selectedId, name, slideIndices);
      const show = selectedId === null ? doc.customShows.at(-1) : doc.customShows.find((item) => item.id === selectedId);
      if (show) edit(show.id);
    });
  }

  function close(): void { editor.closeDialog(); }
</script>

<dialog bind:this={dialog} aria-label={t('Custom Shows')} onclose={() => editor.closeDialog()}>
  <div class="dialog ok-editor">
    <header>
      <strong>{t('Custom Shows')}</strong>
      <button class="ok-btn" aria-label={t('Close')} onclick={close}>✕</button>
    </header>
    <div class="body">
      <section class="shows" aria-label={t('Custom Shows')}>
        {#each doc.customShows as show (show.id)}
          <button class:active={selectedId === show.id} class="show" onclick={() => edit(show.id)}>
            <span>{show.name}</span><small>{show.slides.length} {t('slides')}</small>
          </button>
        {/each}
        {#if doc.customShows.length === 0}<p class="empty">{t('No custom shows')}</p>{/if}
        <button class="ok-btn" onclick={newShow}>＋ {t('New')}</button>
      </section>
      <section class="editor-form">
        <label>{t('Name')}<input class="ok-input" bind:value={name} aria-label={t('Name')} /></label>
        <fieldset>
          <legend>{t('Slides')}</legend>
          {#each doc.slides as _slide, index (index)}
            <button type="button" class="slide-option" aria-label={`${t('Add slide')} ${index + 1}`} onclick={() => addSlide(index)}>＋ {t('Slide')} {index + 1}</button>
          {/each}
        </fieldset>
        {#if error}<p class="error" role="alert">{error}</p>{/if}
        <fieldset class="sequence">
          <legend>{t('Selected slides')}</legend>
          {#each slideIndices as index, sequenceIndex (sequenceIndex)}
            <div class="sequence-row">
              <span>{t('Slide')} {index + 1}</span>
              <button type="button" class="ok-btn" aria-label={t('Move slide in custom show up')} title={t('Move slide in custom show up')} onclick={() => moveSlide(sequenceIndex, -1)}>↑</button>
              <button type="button" class="ok-btn" aria-label={t('Move slide in custom show down')} title={t('Move slide in custom show down')} onclick={() => moveSlide(sequenceIndex, 1)}>↓</button>
              <button type="button" class="ok-btn" aria-label={t('Remove slide from show')} title={t('Remove slide from show')} onclick={() => removeSlide(sequenceIndex)}>×</button>
            </div>
          {/each}
          {#if slideIndices.length === 0}<p class="empty">{t('No slides selected')}</p>{/if}
        </fieldset>
        <div class="actions">
          <button class="ok-btn primary" disabled={!name.trim() || slideIndices.length === 0} onclick={save}>{t('Apply')}</button>
          {#if selectedShow}
            <button class="ok-btn" onclick={() => runUpdate(() => editor.copyCustomShow(selectedShow.id))}>{t('Copy')}</button>
            <button class="ok-btn" onclick={() => runUpdate(() => { editor.deleteCustomShow(selectedShow.id); newShow(); })}>{t('Delete')}</button>
            <button class="ok-btn" aria-label={t('Move custom show up')} title={t('Move custom show up')} onclick={() => runUpdate(() => editor.reorderCustomShow(selectedShow.id, -1))}>↑</button>
            <button class="ok-btn" aria-label={t('Move custom show down')} title={t('Move custom show down')} onclick={() => runUpdate(() => editor.reorderCustomShow(selectedShow.id, 1))}>↓</button>
          {/if}
        </div>
      </section>
    </div>
    <footer><button class="ok-btn" onclick={close}>{t('Close')}</button></footer>
  </div>
</dialog>

<style>
  dialog { width: min(680px, 94vw); max-height: 86vh; padding: 0; overflow: hidden; background: var(--ok-panel); color: var(--ok-text); border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: rgb(0 0 0 / 28%); }
  .dialog { display: flex; flex-direction: column; max-height: 86vh; }
  header, footer { display: flex; align-items: center; justify-content: space-between; padding: 10px 14px; border-bottom: 1px solid var(--ok-border); }
  footer { justify-content: flex-end; border-top: 1px solid var(--ok-border); border-bottom: 0; background: var(--ok-panel-2); }
  .body { display: grid; grid-template-columns: 0.8fr 1.2fr; min-height: 0; overflow: auto; }
  .shows, .editor-form { display: flex; flex-direction: column; gap: 8px; padding: 14px; }
  .shows { border-right: 1px solid var(--ok-border); }
  .show { display: flex; justify-content: space-between; gap: 8px; padding: 8px; border: 1px solid transparent; border-radius: var(--ok-radius); background: transparent; color: var(--ok-text); text-align: left; }
  .show:hover, .show.active { background: var(--ok-selected); border-color: var(--ok-selected-border); }
  small, .empty { color: var(--ok-text-2); }
  fieldset { display: flex; flex-direction: column; gap: 5px; max-height: 180px; overflow: auto; border: 1px solid var(--ok-border); }
  .slide-option { padding: 4px 6px; border: 1px solid transparent; border-radius: var(--ok-radius); background: transparent; color: var(--ok-text); text-align: left; font: inherit; cursor: pointer; }
  .slide-option:hover { background: var(--ok-hover); border-color: var(--ok-border); }
  .sequence { max-height: 180px; }
  .sequence-row { display: flex; align-items: center; gap: 4px; }
  .sequence-row span { flex: 1; }
  .error { margin: 0; color: var(--ok-danger, #bf3131); font-size: 12px; }
  .actions { display: flex; flex-wrap: wrap; gap: 6px; margin-top: auto; }
  .primary { background: var(--ok-accent); color: #fff; }
</style>
