<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getSlideTransition, getSlides, setSlideTransition, clearSlideTransition, type TransitionOptions } from '@office-kit/pptx';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const slide = untrack(() => doc.currentSlide);
  const selection = untrack(() => doc.selection);
  const targets = untrack(() => { const slides = getSlides(doc.pres); return selectedSlideIndices(selection).flatMap(index => slides[index] ? [slides[index]!] : []); });
  const version = untrack(() => doc.version);
  const original = slide ? getSlideTransition(slide) : null;
  const mixed = targets.some(target => JSON.stringify(getSlideTransition(target)) !== JSON.stringify(original));
  const effects = [['none', 'None'], ['fade', 'Fade'], ['push', 'Push'], ['wipe', 'Wipe'], ['cover', 'Cover'], ['pull', 'Uncover'], ['split', 'Split'], ['cut', 'Cut transition'], ['dissolve', 'Dissolve'], ['checker', 'Checkerboard'], ['blinds', 'Blinds'], ['comb', 'Comb'], ['randomBar', 'Random bars'], ['zoom', 'Zoom'], ['circle', 'Circle'], ['diamond', 'Diamond'], ['plus', 'Plus'], ['wedge', 'Wedge'], ['newsflash', 'Newsflash'], ['strips', 'Strips'], ['wheel', 'Wheel'], ['random', 'Random']];
  const directionLabels: Record<string, string> = { l: 'Left', r: 'Right', u: 'Up', d: 'Down', lu: 'Upper left', ru: 'Upper right', ld: 'Lower left', rd: 'Lower right', horz: 'Horizontal', vert: 'Vertical', in: 'Inward', out: 'Outward' };
  let effect = $state(original?.effect ?? 'none');
  let speed = $state<NonNullable<TransitionOptions['speed']>>(original?.speed ?? 'med');
  let direction = $state(original?.direction ?? '');
  let orientation = $state<NonNullable<TransitionOptions['orientation']>>(original?.orientation ?? 'horz');
  let thruBlack = $state(original?.thruBlack ?? false);
  let onClick = $state(original?.advanceOnClick ?? true);
  let auto = $state(original?.advanceAfterMs !== undefined);
  let seconds = $state<number | undefined>((original?.advanceAfterMs ?? 5000) / 1000);
  let allSlides = $state(false);
  let error = $state('');
  let dialog: HTMLDialogElement;
  const directions = $derived(['push', 'wipe'].includes(effect) ? ['l', 'r', 'u', 'd'] : ['cover', 'pull'].includes(effect) ? ['l', 'r', 'u', 'd', 'lu', 'ru', 'ld', 'rd'] : effect === 'strips' ? ['lu', 'ru', 'ld', 'rd'] : ['blinds', 'checker', 'comb', 'randomBar'].includes(effect) ? ['horz', 'vert'] : ['zoom', 'split'].includes(effect) ? ['in', 'out'] : []);
  const valid = $derived(effects.some(item => item[0] === effect) && (!auto || (seconds !== undefined && Number.isFinite(seconds) && seconds >= 0 && seconds <= 4294967.295)));
  onMount(() => dialog.showModal());
  function submit(event: SubmitEvent) {
    event.preventDefault();
    if (!slide || !valid) return;
    if (doc.version !== version || doc.selection !== selection) { error = t('The slide changed. Reopen this dialog.'); return; }
    const options: TransitionOptions = {
      effect, ...(effect !== 'none' ? { speed } : {}), advanceOnClick: onClick,
      ...(auto && seconds !== undefined ? { advanceAfterMs: Math.round(seconds * 1000) } : {}),
      ...(directions.includes(direction) ? { direction } : {}),
      ...(effect === 'split' ? { orientation } : {}),
      ...(['fade', 'cut'].includes(effect) ? { thruBlack } : {}),
    };
    try {
      doc.transact(t('Slide transition'), () => {
        for (const target of allSlides ? getSlides(doc.pres) : targets) {
          if (effect === 'none' && !auto && onClick) clearSlideTransition(target);
          else setSlideTransition(target, options);
        }
      });
      editor.closeDialog();
    } catch (cause) { error = `${t('Slide update failed')}: ${cause instanceof Error ? cause.message : String(cause)}`; }
  }
</script>

<dialog bind:this={dialog} aria-label={t('Slide transition')} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <header><strong>{t('Slide transition')}</strong><button type="button" class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    {#if targets.length > 1}<p>{t('Apply to selected slides')}: {targets.length}</p>{/if}
    {#if mixed}<p>{t('Slide transition')}: {t('Mixed')}</p>{/if}
    <label>{t('Transition effect')}<select class="ok-input" aria-label={t('Transition effect')} bind:value={effect} onchange={() => direction = ''}>
      {#if !effects.some(item => item[0] === effect)}<option value={effect}>{effect}</option>{/if}
      {#each effects as item}<option value={item[0]}>{t(item[1]!)}</option>{/each}
    </select></label>
    {#if effect !== 'none'}<label>{t('Transition speed')}<select class="ok-input" aria-label={t('Transition speed')} bind:value={speed}><option value="slow">{t('Slow')}</option><option value="med">{t('Medium')}</option><option value="fast">{t('Fast')}</option></select></label>{/if}
    {#if directions.length}<label>{t('Transition direction')}<select class="ok-input" aria-label={t('Transition direction')} bind:value={direction}><option value="">{t('Default')}</option>{#each directions as item}<option value={item}>{t(directionLabels[item]!)}</option>{/each}</select></label>{/if}
    {#if effect === 'split'}<label>{t('Split orientation')}<select class="ok-input" aria-label={t('Split orientation')} bind:value={orientation}><option value="horz">{t('Horizontal')}</option><option value="vert">{t('Vertical')}</option></select></label>{/if}
    {#if effect === 'fade' || effect === 'cut'}<label class="option"><input type="checkbox" bind:checked={thruBlack} />{t('Through black')}</label>{/if}
    <label class="option"><input type="checkbox" bind:checked={onClick} />{t('Advance on click')}</label>
    <label class="option"><input type="checkbox" bind:checked={auto} />{t('Advance automatically')}</label>
    {#if auto}<label>{t('Advance after (seconds)')}<input class="ok-input" type="number" min="0" max="4294967.295" step="0.001" required bind:value={seconds} /></label>{/if}
    {#if !valid}<p role="alert">{t('Choose a supported effect and enter a valid time in seconds.')}</p>{/if}
    <label class="option"><input type="checkbox" bind:checked={allSlides} />{t('Apply to all slides')}</label>
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="ok-btn primary" disabled={!valid}>{t('Apply')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(440px, 90vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form { display: grid; gap: 14px; }
  header, footer { display: flex; justify-content: space-between; align-items: center; gap: 12px; }
  footer { justify-content: flex-end; }
  label { display: grid; gap: 6px; }
  .option { display: flex; align-items: center; gap: 8px; }
  p { font-size: 12px; margin: 0; }
  [role='alert'] { color: #bf3131; }
</style>
