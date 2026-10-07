<script lang="ts">
  import { onDestroy } from 'svelte';
  import PictureCropSection from './PictureCropSection.svelte';
  import { getShapeImageBrightness, getShapeImageContrast, getShapeMedia, resetShapeImageColorEffects, setShapeImageBrightness, setShapeImageContrast, type SlideShapeData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import VideoCorrectionsMenu from '../ribbon/VideoCorrectionsMenu.svelte';
  import VideoRecolorMenu from '../ribbon/VideoRecolorMenu.svelte';

  const editor = getEditor();
  const doc = editor.doc;
  const video = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    if (sel.kind !== 'shape' || sel.shapeIds.length !== 1) return null;
    const shape = doc.shapeById(sel.slideIndex, sel.shapeIds[0]!);
    return shape && getShapeMedia(shape)?.kind === 'video' ? shape : null;
  });
  const brightness = $derived.by(() => { doc.version; return video ? Math.round((getShapeImageBrightness(video) ?? 0) * 100) : 0; });
  const contrast = $derived.by(() => { doc.version; return video ? Math.round((getShapeImageContrast(video) ?? 0) * 100) : 0; });
  const locked = $derived(editor.selectionLocked());

  let live: { shape: SlideShapeData; kind: 'brightness' | 'contrast' } | null = null;

  function finishLive(): void {
    const gesture = live;
    live = null;
    if (gesture && doc.liveEditing) doc.commit(t(gesture.kind === 'brightness' ? 'Brightness' : 'Contrast'));
  }

  function previewValue(kind: 'brightness' | 'contrast', input: HTMLInputElement): void {
    const shape = video;
    if (!shape || locked) return;
    if (live && (live.shape !== shape || live.kind !== kind)) finishLive();
    const value = input.valueAsNumber / 100;
    const current = (kind === 'brightness' ? getShapeImageBrightness(shape) : getShapeImageContrast(shape)) ?? 0;
    if (value === current) return;
    live = { shape, kind };
    doc.applyLive(() => (kind === 'brightness' ? setShapeImageBrightness : setShapeImageContrast)(shape, value || null));
  }

  function cancelLive(): void {
    if (!live) return;
    live = null;
    void doc.cancelLive();
  }

  $effect(() => { const selected = video; const disabled = locked; if (live && (live.shape !== selected || disabled)) finishLive(); });
  onDestroy(finishLive);

  function resetColor(): void {
    const shape = video;
    if (!shape || locked) return;
    doc.transact(t('Reset picture color'), () => resetShapeImageColorEffects(shape));
  }

  function setValue(kind: 'brightness' | 'contrast', input: HTMLInputElement): void {
    if (!video || !input.reportValidity()) return;
    const value = input.valueAsNumber;
    if (!Number.isFinite(value)) return;
    editor.invoke(kind === 'brightness' ? 'setShapeImageBrightness' : 'setShapeImageContrast', { value: value / 100 || null });
  }
</script>

{#if video}
  <section class="video-controls" aria-label={t('Video options')}>
    <details open>
      <summary>{t('Video')}</summary>
      <div class="video-body">
        <div class="subheading"><span class="subheading-label">{t('Recolor')}</span> <VideoRecolorMenu variant="pane" /></div>
        <div class="subheading"><span class="subheading-label">{t('Brightness')} / {t('Contrast')}</span> <VideoCorrectionsMenu variant="pane" /></div>
        <div class="control">
          <div class="control-row">
            <label for="video-brightness-range">{t('Brightness')}</label>
            <input id="video-brightness-range" type="range" min="-100" max="100" step="1" value={brightness} aria-label={t('Brightness')} disabled={locked} oninput={(e) => previewValue('brightness', e.currentTarget)} onchange={finishLive} onblur={finishLive} onpointercancel={cancelLive} />
            <input class="ok-input value" type="number" min="-100" max="100" step="1" required value={brightness} aria-label={t('Brightness')} disabled={locked} onchange={(e) => setValue('brightness', e.currentTarget)} />
            <span>%</span>
          </div>
        </div>
        <div class="control">
          <div class="control-row">
            <label for="video-contrast-range">{t('Contrast')}</label>
            <input id="video-contrast-range" type="range" min="-100" max="100" step="1" value={contrast} aria-label={t('Contrast')} disabled={locked} oninput={(e) => previewValue('contrast', e.currentTarget)} onchange={finishLive} onblur={finishLive} onpointercancel={cancelLive} />
            <input class="ok-input value" type="number" min="-100" max="100" step="1" required value={contrast} aria-label={t('Contrast')} disabled={locked} onchange={(e) => setValue('contrast', e.currentTarget)} />
            <span>%</span>
          </div>
        </div>
        <button class="ok-btn reset" type="button" disabled={locked} onclick={resetColor}>{t('Reset')}</button>
      </div>
    </details>
  </section>
  <PictureCropSection picture={video} />
{/if}

<style>
  .video-controls { min-width: 0; padding: 12px; display: flex; flex-direction: column; gap: 14px; border-bottom: 1px solid var(--ok-border); }
  details { min-width: 0; font-size: 12px; }
  summary { cursor: pointer; }
  details { display: grid; gap: 10px; }
  .video-body { min-width: 0; display: grid; gap: 10px; }
  .subheading { min-width: 0; display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 11px; color: var(--ok-text-2); }
  .subheading-label { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  :global(.subheading .trigger) { flex: 0 1 auto; }
  .control { min-width: 0; display: grid; gap: 5px; }
  label { width: 66px; flex: 0 0 66px; font-size: 11px; color: var(--ok-text-2); }
  .control-row { display: flex; align-items: center; gap: 7px; }
  input[type=range] { flex: 1 1 auto; width: 0; min-width: 0; accent-color: var(--ok-accent); }
  .value { flex: 0 0 54px; width: 54px; box-sizing: border-box; }
  .reset { justify-self: start; }
  span { font-size: 11px; color: var(--ok-text-2); }
</style>
