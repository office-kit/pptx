<script lang="ts">
  import type { TextFormat } from '@office-kit/pptx';
  import { t } from '../i18n/i18n.svelte.ts';

  let { formats, selected, onformat, ondone }: {
    formats: TextFormat[];
    selected: boolean;
    onformat: (format: TextFormat) => void;
    ondone: () => void;
  } = $props();
  const bold = $derived(formats.length > 0 && formats.every((f) => f.bold === true));
  const italic = $derived(formats.length > 0 && formats.every((f) => f.italic === true));
  const underline = $derived(formats.length > 0 && formats.every((f) => f.underline === true || (typeof f.underline === 'string' && f.underline !== 'none')));
  const font = $derived(formats.length && formats.every((f) => f.font === formats[0]?.font) ? formats[0]?.font ?? '' : '');
  const size = $derived(formats.length && formats.every((f) => f.size === formats[0]?.size) ? formats[0]?.size : undefined);
</script>

<div class="text-format-bar" role="group" aria-label={t('Selected text formatting')}>
  <span>{t(selected ? 'Selected text' : 'Select text to format')}</span>
  <button class="ok-btn" aria-label={t('Bold')} aria-pressed={bold} disabled={!selected} onmousedown={(e) => e.preventDefault()} onclick={() => onformat({bold: !bold})}><b>B</b></button>
  <button class="ok-btn" aria-label={t('Italic')} aria-pressed={italic} disabled={!selected} onmousedown={(e) => e.preventDefault()} onclick={() => onformat({italic: !italic})}><i>I</i></button>
  <button class="ok-btn" aria-label={t('Underline')} aria-pressed={underline} disabled={!selected} onmousedown={(e) => e.preventDefault()} onclick={() => onformat({underline: !underline})}><u>U</u></button>
  <label>{t('Font')}<input class="ok-input font" aria-label={t('Font')} disabled={!selected} value={font} placeholder={t('Mixed or inherited')} onchange={(e) => { const font = e.currentTarget.value.trim(); if (font) onformat({font, fontEastAsian: font, fontComplexScript: font}); }} /></label>
  <label>{t('Font size')}<input class="ok-input size" aria-label={t('Font size')} type="number" min="1" max="4000" step="0.5" disabled={!selected} value={size ?? ''} placeholder="—" onchange={(e) => { if (e.currentTarget.value && e.currentTarget.reportValidity()) onformat({size:e.currentTarget.valueAsNumber}); }} /></label>
  <label>{t('Text color')}<input aria-label={t('Text color')} type="color" disabled={!selected} onchange={(e) => onformat({color:e.currentTarget.value})} /></label>
  <button class="ok-btn" onclick={ondone}>{t('Done')}</button>
</div>

<style>
  .text-format-bar { display: flex; flex-wrap: wrap; align-items: end; gap: 5px; padding: 6px 8px; background: var(--ok-panel); border-bottom: 1px solid var(--ok-border); }
  span { font-size: 11px; align-self: center; }
  label { display: grid; gap: 2px; font-size: 10px; }
  .font { width: 110px; }
  .size { width: 56px; }
  input[type='color'] { width: 30px; height: 26px; padding: 0; border: 1px solid var(--ok-border); }
  button[aria-pressed='true'] { background: var(--ok-accent); color: white; }
</style>
