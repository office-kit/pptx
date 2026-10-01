<script lang="ts">
  import {
    getShapeMedia,
    getShapeMediaPlayback,
    setShapeMediaPlayback,
    type MediaPlayback,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  let error = $state('');

  const selected = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    if (shapes.length !== 1) return null;
    const shape = shapes[0]!;
    const media = getShapeMedia(shape);
    const playback = getShapeMediaPlayback(shape);
    return media !== null && playback !== null ? { shape, media, playback } : null;
  });

  const delaySeconds = $derived(selected?.playback.delayMs === undefined ? 0 : selected.playback.delayMs / 1000);

  function apply(label: string, options: Partial<MediaPlayback>): void {
    const target = selected;
    if (target === null) return;
    try {
      doc.transact(t(label), () => setShapeMediaPlayback(target.shape, options));
      error = '';
    } catch (cause) {
      error = cause instanceof Error ? cause.message : String(cause);
    }
  }

  function changeDelay(input: HTMLInputElement): void {
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) return;
    const milliseconds = Math.round(input.valueAsNumber * 1000);
    if (!Number.isSafeInteger(milliseconds)) {
      error = t('Delay must be a whole number of milliseconds');
      return;
    }
    apply('Start delay', { delayMs: milliseconds });
  }

  function changeVolume(input: HTMLInputElement): void {
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) return;
    apply('Volume', { volume: input.valueAsNumber / 100 });
  }
</script>

{#if selected}
  <div class="group">
    <div class="items">
      <label class="field">
        <span>{t('Start')}</span>
        <select aria-label={t('Start')} value={selected.playback.autoplay ? 'automatic' : 'click'} onchange={(event) => apply('Start', { autoplay: event.currentTarget.value === 'automatic' })}>
          <option value="automatic">{t('Automatically')}</option>
          <option value="click">{t('When Clicked')}</option>
        </select>
      </label>
      <label class="check">
        <input type="checkbox" checked={selected.playback.loop} onchange={(event) => apply('Loop until stopped', { loop: event.currentTarget.checked })} />
        <span>{t('Loop until stopped')}</span>
      </label>
      <label class="field number-field">
        <span>{t('Volume')}</span>
        <span class="number"><input class="ok-input" type="number" min="0" max="100" step="1" required aria-label={t('Volume')} value={Math.round(selected.playback.volume * 100)} onchange={(event) => changeVolume(event.currentTarget)} /><span>%</span></span>
      </label>
      <label class="check">
        <input type="checkbox" checked={selected.playback.muted} onchange={(event) => apply('Mute', { muted: event.currentTarget.checked })} />
        <span>{t('Mute')}</span>
      </label>
      {#if selected.media.kind === 'video'}
        <label class="check">
          <input type="checkbox" checked={selected.playback.fullScreen} onchange={(event) => apply('Play Full Screen', { fullScreen: event.currentTarget.checked })} />
          <span>{t('Play Full Screen')}</span>
        </label>
      {/if}
      <label class="check">
        <input type="checkbox" checked={selected.playback.hideWhenStopped} onchange={(event) => apply('Hide when not playing', { hideWhenStopped: event.currentTarget.checked })} />
        <span>{t('Hide when not playing')}</span>
      </label>
    </div>
    {#if error}<span class="error" role="alert">{error}</span>{/if}
    <span class="title">{t('Playback')}</span>
  </div>
  {#if selected.playback.autoplay}
    <div class="group timing">
      <div class="items">
        <label class="field number-field">
          <span>{t('Start delay (seconds)')}</span>
          <span class="number"><input class="ok-input" type="number" min="0" step="0.001" required aria-label={t('Start delay (seconds)')} value={delaySeconds} onchange={(event) => changeDelay(event.currentTarget)} /><span>s</span></span>
        </label>
      </div>
      <span class="title">{t('Timing')}</span>
    </div>
  {/if}
{/if}

<style>
  .group { display: flex; flex-direction: column; justify-content: space-between; flex-shrink: 0; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  .timing { min-width: 145px; }
  .items { display: flex; align-items: center; flex: 1; gap: 8px; }
  .field, .check { display: flex; align-items: center; gap: 4px; color: var(--ok-text); font: inherit; font-size: 11px; white-space: nowrap; }
  .field { flex-direction: column; align-items: stretch; gap: 2px; }
  select, input { font: inherit; color: inherit; }
  select { min-width: 112px; padding: 3px 4px; }
  .number { display: flex; align-items: center; gap: 3px; }
  .number input { width: 52px; min-width: 0; padding: 2px 4px; }
  .check { max-width: 125px; }
  .check input { margin: 0; }
  .check span { white-space: normal; }
  .error { color: var(--ok-danger, #b42318); font-size: 11px; padding-top: 3px; }
  .title { text-align: center; font-size: 10px; color: var(--ok-text-2); padding: 4px 0 1px; }
</style>
