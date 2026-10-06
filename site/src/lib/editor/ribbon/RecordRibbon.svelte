<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  // Recording narration and camera needs the PowerPoint recording studio; the
  // browser editor plays the show instead, so Clear Recording and Reset to
  // Cameo stay disabled exactly as PowerPoint shows them for a deck with no
  // recording.
  const UNAVAILABLE = 'Recording is not available in the browser.';
</script>

<div class="group">
  <button disabled title={t(UNAVAILABLE)}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="14" rx="1" /><circle cx="15" cy="10" r="2.5" /><path d="M11 18c0-3 8-3 8 0" /></svg>
    {t('Cameo')}
  </button>
</div>
<div class="group">
  <button disabled={!editor.canPresent} onclick={() => editor.present('start')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="14" height="11" rx="1" /><path d="M10 14v5M6 19h8M8 6l4 2.5-4 2.5z" /><circle cx="19" cy="15" r="3" fill="#c92a2a" stroke="none" /></svg>
    {t('From Beginning')}
  </button>
  <button disabled={!editor.canPresent} onclick={() => editor.present('current')}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="14" height="11" rx="1" /><path d="M10 14v5M6 19h8M6 6h5v5H6z" /><circle cx="19" cy="15" r="3" fill="#c92a2a" stroke="none" /></svg>
    {t('From Current Slide')}
  </button>
</div>
<div class="group">
  <button disabled title={t(UNAVAILABLE)}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="16" height="12" stroke-dasharray="2 2" /><path d="m14 13 7 7M21 13l-7 7" stroke="#c92a2a" /></svg>
    {t('Clear Recording')}
  </button>
  <button disabled title={t(UNAVAILABLE)}>
    <svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="6" width="16" height="12" rx="1" /><circle cx="13" cy="12" r="2" /><path d="M5 3h6" /></svg>
    {t('Reset to Cameo')}
  </button>
</div>
<div class="group">
  <a class="button" href="https://support.microsoft.com/office/record-a-slide-show-with-narration-and-slide-timings-0b9502c6-5f6c-40ae-b1e7-e47d8741161c" target="_blank" rel="noreferrer">
    <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l4 4v14H6z" /><path d="M11 12a2 2 0 1 1 3 2c-1 .5-1 1-1 2M13 18.5v.5" style="stroke: var(--ok-accent)" /></svg>
    {t('Learn More')}
  </a>
</div>

<style>
  .group { display: flex; align-items: center; flex-shrink: 0; gap: 2px; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  button, .button { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: 2px; min-width: 52px; max-width: 96px; min-height: 66px; padding: 3px 4px; font: inherit; font-size: 11px; line-height: 1.2; color: var(--ok-text); text-align: center; text-decoration: none; background: transparent; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:not(:disabled):hover, .button:hover { background: var(--ok-hover); border-color: var(--ok-border); }
  button:disabled { opacity: 0.45; cursor: default; }
  svg { width: 32px; height: 32px; flex-shrink: 0; stroke: currentColor; fill: none; stroke-width: 1.1; }
</style>
