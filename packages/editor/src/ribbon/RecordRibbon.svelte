<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  // Recording narration and camera needs the PowerPoint recording studio; the
  // browser editor plays the show instead, so Clear Recording and Reset to
  // Cameo stay disabled exactly as PowerPoint shows them for a deck with no
  // recording.
  const UNAVAILABLE = 'Recording is not available in the browser.';
</script>

<!-- Mac PowerPoint 16's Record tab: Cameo | From Beginning, From Current
     Slide | Clear Recording ▾, Reset to Cameo ▾ | Learn More. -->
<div class="record" lang={getLocale()}>
  <section class="cluster" role="group" aria-label={t('Camera')}>
    <button class="big" style:--w="42px" disabled title={t(UNAVAILABLE)}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="14" rx="1" /><circle cx="15" cy="10" r="2.5" /><path d="M11 18c0-3 8-3 8 0" /></svg>
      <span>{t('Cameo')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Record')}>
    <button class="big" style:--w="57px" disabled={!editor.canPresent} onclick={() => editor.present('start')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="14" height="11" rx="1" /><path d="M10 14v5M6 19h8M8 6l4 2.5-4 2.5z" /><circle cx="19" cy="15" r="3" fill="#c92a2a" stroke="none" /></svg>
      <span>{t('From Beginning')}</span>
    </button>
    <button class="big" style:--w="73px" disabled={!editor.canPresent} onclick={() => editor.present('current')}>
      <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="14" height="11" rx="1" /><path d="M10 14v5M6 19h8M6 6h5v5H6z" /><circle cx="19" cy="15" r="3" fill="#c92a2a" stroke="none" /></svg>
      <span>{t('From Current Slide')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Edit')}>
    <button class="big" style:--w="58px" aria-haspopup="menu" disabled title={t(UNAVAILABLE)}>
      <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="16" height="12" stroke-dasharray="2 2" /><path d="m14 13 7 7M21 13l-7 7" stroke="#c92a2a" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
      <span>{t('Clear Recording')}</span>
    </button>
    <button class="big" style:--w="50px" aria-haspopup="menu" disabled title={t(UNAVAILABLE)}>
      <span class="icon-row"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="6" width="16" height="12" rx="1" /><circle cx="13" cy="12" r="2" /><path d="M5 3h6" /></svg><span class="arrow" aria-hidden="true">⌄</span></span>
      <span>{t('Reset to Cameo')}</span>
    </button>
  </section>
  <section class="cluster" role="group" aria-label={t('Help')}>
    <a class="big" style:--w="38px" href="https://support.microsoft.com/office/record-a-slide-show-with-narration-and-slide-timings-0b9502c6-5f6c-40ae-b1e7-e47d8741161c" target="_blank" rel="noreferrer">
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l4 4v14H6z" /><path d="M11 12a2 2 0 1 1 3 2c-1 .5-1 1-1 2M13 18.5v.5" style="stroke: var(--ok-accent)" /></svg>
      <span>{t('Learn More')}</span>
    </a>
  </section>
</div>

<style>
  /* Geometry measured from Mac PowerPoint 16 (NATIVE_PARITY.md, "Native
     geometry audit"). */
  .record { display: flex; align-items: stretch; width: 100%; min-width: 0; height: 72px; }
  .cluster { display: flex; flex: none; align-items: stretch; padding: 0 10px; border-right: 1px solid var(--ok-border); }
  .cluster:first-child { padding-left: 4px; }
  .cluster:last-child { border-right: none; }
  .big { box-sizing: border-box; display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: var(--w); padding: 4px 1px; font: inherit; font-size: 11px; line-height: 1.15; color: var(--ok-text); text-align: center; text-decoration: none; background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  .big > span:last-child { max-width: var(--w); margin: 0 -2px; }
  /* Japanese labels wrap per character, so they get at least six characters
     a line and a smaller size that fits three lines (PowerPoint widens them). */
  .big > span:last-child:lang(ja) { max-width: max(calc(var(--w) - 4px), 6em); font-size: 10px; line-height: 1.1; }
  .big:hover:not(:disabled) { background: var(--ok-hover); }
  .big:disabled { opacity: 0.4; cursor: default; }
  .icon-row { display: flex; align-items: center; gap: 1px; }
  .arrow { font-size: 11px; }
  svg { width: 32px; height: 32px; flex: none; fill: none; stroke: currentColor; stroke-width: 1.1; }
</style>
