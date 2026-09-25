<script lang="ts">
  import { onMount } from 'svelte';
  import { addSlideTable, emu, getShapeId, inches } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { slideMetrics } from '../canvas/geometry.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  let dialog: HTMLDialogElement;
  let rows = $state<number | undefined>(3);
  let columns = $state<number | undefined>(3);
  let header = $state(true);
  let banded = $state(true);
  let error = $state('');
  const valid = $derived(rows !== undefined && columns !== undefined && Number.isInteger(rows) && Number.isInteger(columns) && rows >= 1 && columns >= 1 && rows <= 50 && columns <= 50);
  onMount(() => dialog.showModal());

  function submit(event: SubmitEvent) {
    event.preventDefault();
    const doc = editor.doc;
    const slide = doc.currentSlide;
    if (!valid || rows === undefined || columns === undefined || !slide) return;
    const columnCount = columns;
    const cells = Array.from({ length: rows }, () => Array.from({ length: columnCount }, () => ''));
    const metrics = slideMetrics(doc.pres);
    const w = emu(Math.round(metrics.widthEmu * 0.8));
    const h = emu(Math.round(Math.min(metrics.heightEmu * 0.7, rows * inches(0.4))));
    try {
      doc.transact(t('Insert table'), () => {
        const table = addSlideTable(slide, {
          x: emu(Math.round((metrics.widthEmu - w) / 2)),
          y: emu(Math.round((metrics.heightEmu - h) / 2)),
          w, h, rows: cells, firstRow: header, bandRow: banded,
        });
        doc.selectCell(doc.selection.slideIndex, getShapeId(table), 0, 0);
      });
      editor.closeDialog();
    } catch (cause) {
      error = `${t('The table could not be inserted')}: ${cause instanceof Error ? cause.message : String(cause)}`;
    }
  }
</script>

<dialog bind:this={dialog} aria-label={t('Insert table')} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <header><strong>{t('Insert table')}</strong><button type="button" class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    <div class="dimensions">
      <label>{t('Number of rows')}<input class="ok-input" type="number" min="1" max="50" step="1" required bind:value={rows} /></label>
      <label>{t('Number of columns')}<input class="ok-input" type="number" min="1" max="50" step="1" required bind:value={columns} /></label>
    </div>
    <label class="option"><input type="checkbox" bind:checked={header} />{t('Header row')}</label>
    <label class="option"><input type="checkbox" bind:checked={banded} />{t('Alternating row colors')}</label>
    {#if valid && rows && columns}
      <div class="preview" role="img" aria-label={`${t('Table preview')}: ${rows} × ${columns}`} style="grid-template-columns: repeat({Math.min(columns, 8)}, 1fr);">
        {#each Array.from({ length: Math.min(rows, 8) }) as _, r}
          {#each Array.from({ length: Math.min(columns, 8) }) as _}
            <span class:header={header && r === 0} class:banded={banded && r % 2 === 1}></span>
          {/each}
        {/each}
      </div>
      <div class="size">{rows} × {columns}</div>
    {/if}
    <p>{t('The table is centered on the slide. Edit its cells after inserting it.')}</p>
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="ok-btn primary" disabled={!valid}>{t('Insert table')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(400px, 90vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form { display: flex; flex-direction: column; gap: 14px; }
  header, footer { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
  footer { justify-content: flex-end; }
  .dimensions { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  label { display: grid; gap: 6px; }
  input[type='number'] { width: 100%; }
  .option { display: flex; align-items: center; gap: 8px; }
  .preview { display: grid; gap: 1px; border: 1px solid var(--ok-border); background: var(--ok-border); }
  .preview span { height: 16px; background: var(--ok-panel); }
  .preview span.banded { background: var(--ok-hover); }
  .preview span.header { background: var(--ok-accent); }
  .size { text-align: center; font-size: 12px; }
  p { font-size: 12px; color: var(--ok-text-2); margin: 0; }
  [role='alert'] { color: #bf3131; }
</style>
