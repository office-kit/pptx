<script lang="ts">
  // The question the reference desktop app asks when a rehearsal ends: keep the new times as
  // each slide's "After" timing?
  import { onMount } from 'svelte';
  import {
    getSlides,
    getSlideShowProperties,
    getSlideTransition,
    setSlideShowProperties,
    setSlideTransition,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { writableTransition } from '../core/transition-effects.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const timings = editor.rehearsalTimings ?? [];
  const total = timings.reduce((sum, item) => sum + item.ms, 0);
  const seconds = Math.round(total / 1000);
  const clock = `${Math.floor(seconds / 3600)}:${String(Math.floor(seconds / 60) % 60).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
  let dialog: HTMLDialogElement;
  onMount(() => dialog.showModal());

  function close() {
    editor.rehearsalTimings = null;
    editor.closeDialog();
  }
  function keep() {
    const slides = getSlides(doc.pres);
    doc.transact(t('Rehearse Timings'), () => {
      for (const { slide, ms } of timings) {
        const target = slides[slide];
        if (!target) continue;
        const current = writableTransition(getSlideTransition(target));
        // A deck can carry effects the library cannot write; leave those slides alone.
        if (!current) continue;
        setSlideTransition(target, { ...current, advanceAfterMs: ms });
      }
      setSlideShowProperties(doc.pres, { ...getSlideShowProperties(doc.pres), useTimings: true });
    });
    close();
  }
</script>

<dialog bind:this={dialog} aria-label={t('Rehearse Timings')} onclose={close}>
  <p>{t('The total time for the slide show was {time}. Do you want to save the new slide timings?').replace('{time}', clock)}</p>
  <footer>
    <button type="button" class="ok-btn" onclick={close}>{t('No')}</button>
    <button type="button" class="ok-btn primary" onclick={keep}>{t('Yes')}</button>
  </footer>
</dialog>

<style>
  dialog { width: min(420px, 90vw); padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  p { margin: 0 0 16px; line-height: 1.5; }
  footer { display: flex; justify-content: flex-end; gap: 8px; }
</style>
