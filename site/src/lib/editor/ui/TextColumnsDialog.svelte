<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { cm, getShapeTextColumns, setShapeTextColumns } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const { onclose }: { onclose: () => void } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const initial = untrack(() => ({ selection: doc.selection, version: doc.version, shapes: editor.selectedShapes() }));
  const values = initial.shapes.map(getShapeTextColumns);
  const counts = values.map(value => value?.count ?? 1);
  const gaps = values.map(value => value?.gapEmu ?? 0);
  let count = $state<number | undefined>(counts.every(value => value === counts[0]) ? counts[0] : undefined);
  let spacing = $state<number | undefined>(gaps.every(value => value === gaps[0]) ? gaps[0]! / cm(1) : undefined);
  let dialog: HTMLDialogElement;
  onMount(() => dialog.showModal());
  $effect(() => { if (doc.selection !== initial.selection || doc.version !== initial.version) onclose(); });
  function submit(event: SubmitEvent) {
    event.preventDefault();
    if (doc.selection !== initial.selection || doc.version !== initial.version) { onclose(); return; }
    if (count !== undefined && (!Number.isInteger(count) || count < 1 || count > 16)) return;
    if (spacing !== undefined && (!Number.isFinite(spacing) || spacing < 0 || spacing > 40.64)) return;
    if (count !== undefined || spacing !== undefined) {
      doc.transact(t('Columns'), () => {
        for (let i = 0; i < initial.shapes.length; i++) {
          const columns = count ?? counts[i]!;
          setShapeTextColumns(initial.shapes[i]!, columns === 1 ? null : { count: columns, gapEmu: spacing === undefined ? gaps[i]! : cm(spacing) });
        }
      });
    }
    onclose();
  }
</script>

<dialog bind:this={dialog} aria-label={t('Columns')} {onclose} onkeydown={event => event.stopPropagation()}>
  <form onsubmit={submit}>
    <h2>{t('Columns')}</h2>
    <label>{t('Number of columns:')}<input type="number" min="1" max="16" step="1" bind:value={count} placeholder={t('Mixed')} /></label>
    <label>{t('Spacing between columns:')}<span><input type="number" min="0" max="40.64" step="any" bind:value={spacing} placeholder={t('Mixed')} /> cm</span></label>
    <footer><button type="button" onclick={onclose}>{t('Cancel')}</button><button type="submit">{t('OK')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: 340px; border: 1px solid #666; border-radius: 7px; background: #303030; color: #eee; padding: 0; box-shadow: 0 15px 60px #0008; font-size: 13px; color-scheme: dark; }
  dialog::backdrop { background: #0003; }
  h2 { font-size: 13px; font-weight: 600; text-align: center; margin: 0 0 12px; padding: 6px; background: #3a3a3a; border-bottom: 1px solid #242424; }
  label { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin: 14px 19px; }
  input { width: 70px; padding: 2px 4px; border: 1px solid #777; border-radius: 4px; background: #252525; color: inherit; font: inherit; }
  footer { display: flex; justify-content: flex-end; gap: 10px; padding: 0 17px 16px; }
  footer button { min-width: 70px; padding: 3px 8px; border: 0; border-radius: 5px; background: #626262; color: inherit; font-size: 12px; }
  footer button[type='submit'] { background: #087bfa; color: white; }
</style>
