<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { cm, type ParagraphProperties, type ParagraphTabStop } from '@office-kit/pptx';
  import { editTabStops, type TabStopEdit } from '../core/paragraph-tabs.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  let { properties, onapply, onclose }: { properties: ParagraphProperties[]; onapply: (edits: TabStopEdit[], defaultSize: number | undefined) => void; onclose: () => void } = $props();
  const initial = untrack(() => properties);
  const sizes = initial.map(p => p.defaultTabSizeEmu ?? cm(2.54));
  const originalSize = sizes.every(size => size === sizes[0]) ? sizes[0] : undefined;
  let defaultSize = $state(originalSize === undefined ? undefined : originalSize / cm(1));
  let positionTouched = $state(false);
  let position = $state<number | undefined>(0);
  let alignment = $state<ParagraphTabStop['alignment']>('left');
  let edits = $state<TabStopEdit[]>([]);
  const lists = $derived(initial.map(p => editTabStops(p.tabStops ?? [], edits)));
  const stops = $derived.by(() => {
    const common = new Map((lists[0] ?? []).map(stop => [stop.positionEmu, stop]));
    for (const list of lists.slice(1)) {
      const positions = new Map(list.map(stop => [stop.positionEmu, stop.alignment]));
      for (const [position, stop] of common) {
        if (positions.get(position) !== stop.alignment) common.delete(position);
      }
    }
    return [...common.values()];
  });
  const selected = $derived(stops.find(stop => stop.positionEmu === (position === undefined ? undefined : cm(position))));
  const positionEnabled = $derived((positionTouched || !!selected) && position !== undefined && position >= 0 && position <= 142.24);
  const cleared = $derived(edits.some(edit => edit.kind === 'clearAll') ? t('All') : edits.filter(edit => edit.kind === 'clear').map(edit => `${edit.positionEmu / cm(1)} cm`).join(', '));
  let dialog: HTMLDialogElement;
  onMount(() => dialog.showModal());
  function setStop() {
    if (position === undefined || position < 0 || position > 142.24) return;
    edits = [...edits, { kind: 'set', stop: { positionEmu: cm(position), alignment } }];
  }
  function clearStops(all: boolean) {
    if (all) edits = [...edits, { kind: 'clearAll' }];
    else if (position !== undefined) edits = [...edits, { kind: 'clear', positionEmu: cm(position) }];
    position = 0;
    positionTouched = false;
  }
  function submit(event: SubmitEvent) {
    event.preventDefault();
    const size = defaultSize === undefined ? undefined : cm(defaultSize);
    onapply(edits, size === originalSize ? undefined : size);
    onclose();
  }
</script>
<dialog bind:this={dialog} aria-label={t('Tabs')} {onclose} onkeydown={event => event.stopPropagation()}>
  <form onsubmit={submit}>
    <h2>{t('Tabs')}</h2>
    <div class="body">
      <div class="columns">
        <div class="settings">
          <label>{t('Tab stop position:')}<span class="measurement"><input type="number" aria-label={t('Tab stop position:')} min="0" max="142.24" step="any" bind:value={position} oninput={() => positionTouched = true} /> cm</span></label>
          <label>{t('Default tab stops:')}<span class="measurement"><input type="number" aria-label={t('Default tab stops:')} min="0" max="142.24" step="any" bind:value={defaultSize} placeholder={t('Mixed')} /> cm</span></label>
          <fieldset><legend>{t('Alignment:')}</legend>{#each [{ value: 'left', label: 'Left' }, { value: 'center', label: 'Center' }, { value: 'right', label: 'Right' }, { value: 'decimal', label: 'Decimal' }] as item}<label class="radio"><input type="radio" name="tab-alignment" value={item.value} bind:group={alignment} />{t(item.label)}</label>{/each}</fieldset>
        </div>
        <div>
          <select size="7" aria-label={t('Tab stops list')} onchange={event => { const stop = stops.find(item => item.positionEmu === Number(event.currentTarget.value)); if (stop) { position = stop.positionEmu / cm(1); alignment = stop.alignment; } }}>
            {#each stops as stop}<option value={stop.positionEmu} selected={selected?.positionEmu === stop.positionEmu}>{stop.positionEmu / cm(1)} cm</option>{/each}
          </select>
          <div class="actions"><button type="button" onclick={setStop} disabled={!positionEnabled}>{t('Set')}</button><button type="button" disabled={!positionEnabled} onclick={() => clearStops(false)}>{t('Clear')}</button></div>
        </div>
      </div>
      <p>{t('Tab stops to be cleared:')} {cleared}</p>
      <div class="clear-row"><button type="button" onclick={() => clearStops(true)}>{t('Clear All')}</button></div>
    </div>
    <footer><button type="button" onclick={onclose}>{t('Cancel')}</button><button type="submit">{t('OK')}</button></footer>
  </form>
</dialog>
<style>
  dialog { width: 432px; border: 1px solid #666; border-radius: 7px; background: #303030; color: #eee; padding: 0; box-shadow: 0 15px 60px #0008; font-size: 13px; color-scheme: dark; }
  dialog::backdrop { background: #0003; }
  h2 { font-size: 13px; text-align: left; margin: 0; padding: 8px 18px; background: #3a3a3a; }
  .body { padding: 12px 18px; }
  .columns { display: grid; grid-template-columns: 1.15fr 1fr; gap: 16px; }
  label { display: flex; align-items: center; justify-content: space-between; gap: 6px; margin-bottom: 12px; }
  input[type='number'] { width: 48px; }
  input, select { padding: 2px 4px; border: 1px solid #777; border-radius: 4px; background: #252525; color: inherit; font: inherit; }
  select { width: 100%; height: 154px; border-radius: 0; }
  .measurement { display: flex; align-items: center; gap: 3px; }
  fieldset { border: 0; padding: 8px 0; }
  .radio { justify-content: flex-start; gap: 7px; margin: 0; }
  .actions { display: flex; gap: 10px; margin-top: 8px; }
  button { min-width: 80px; padding: 3px 8px; border: 0; border-radius: 5px; background: #626262; color: inherit; font: inherit; }
  button:disabled { opacity: .4; }
  footer { display: flex; justify-content: flex-end; gap: 10px; padding: 4px 18px 16px; }
  .clear-row { display: flex; justify-content: flex-end; }
  p { margin: 0; min-height: 20px; }
  button[type='submit'] { background: #087bfa; color: white; }
</style>
