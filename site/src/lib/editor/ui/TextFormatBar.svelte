<script lang="ts">
  import type { TextFormat } from '@office-kit/pptx';
  import { textFormatActive, toggleTextFormat, type TextFormatToggle } from '../core/text-format-toggle.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { formats, selected, typing = false, onformat, ontoggle, ondone, onlink, paragraph, onparagraph, context = 'text' }: {
    formats: TextFormat[];
    selected: boolean;
    typing?: boolean;
    onformat: (format: TextFormat, reset?: boolean) => void;
    ontoggle?: (property: TextFormatToggle) => void;
    ondone?: () => void;
    onlink?: () => void;
    paragraph?: { align: string; bullet: string; level: string; lineKind: string; lineValue: string; before: string; after: string };
    onparagraph?: (kind: 'align' | 'bullet' | 'level' | 'lineKind' | 'lineValue' | 'before' | 'after', value: string) => void;
    context?: 'text' | 'cells' | 'objects';
  } = $props();
  function toggle(property: TextFormatToggle) {
    if (ontoggle) ontoggle(property);
    else onformat(toggleTextFormat(formats, property));
  }
  const font = $derived(formats.length && formats.every((f) => f.font === formats[0]?.font) ? formats[0]?.font ?? '' : '');
  const size = $derived(formats.length && formats.every((f) => f.size === formats[0]?.size) ? formats[0]?.size : undefined);
  const highlight = $derived(formats.length && formats.every(f => f.highlight === formats[0]?.highlight) && /^#[0-9a-f]{6}$/i.test(formats[0]?.highlight ?? '') ? formats[0]!.highlight! : null);
  const color = $derived(formats.length && formats.every((f) => f.color === formats[0]?.color) && /^#[0-9a-f]{6}$/i.test(formats[0]?.color ?? '') ? formats[0]!.color! : null);
</script>

<div class="text-format-bar" role="group" aria-label={t(context === 'cells' ? 'Format selected cells' : context === 'objects' ? 'Format selected objects' : 'Selected text formatting')}>
  <span>{t(context === 'cells' ? 'Formatting applies to all selected cells' : context === 'objects' ? 'Formatting applies to all text in selected objects' : selected ? 'Selected text' : typing ? 'Text to type' : onparagraph ? 'Current paragraph' : 'Select text to format')}</span>
  {#each [
    { property: 'bold', label: 'Bold' },
    { property: 'italic', label: 'Italic' },
    { property: 'underline', label: 'Underline' },
    { property: 'strike', label: 'Strikethrough' },
    { property: 'superscript', label: 'Superscript' },
    { property: 'subscript', label: 'Subscript' },
  ] as const as item}
    <button class="ok-btn" aria-label={t(item.label)} title={t(item.label)} aria-pressed={textFormatActive(formats, item.property)} disabled={!(selected || typing)} onmousedown={e => e.preventDefault()} onclick={() => toggle(item.property)}>
      {#if item.property === 'bold'}<b>B</b>{:else if item.property === 'italic'}<i>I</i>{:else if item.property === 'underline'}<u>U</u>{:else if item.property === 'strike'}<s>S</s>{:else if item.property === 'superscript'}x<sup>2</sup>{:else}x<sub>2</sub>{/if}
    </button>
  {/each}
  <label>{t('Font')}<input class="ok-input font" aria-label={t('Font')} disabled={!(selected || typing)} value={font} placeholder={t('Mixed or inherited')} onchange={(e) => { const font = e.currentTarget.value.trim(); if (font) onformat({font, fontEastAsian: font, fontComplexScript: font}); }} /></label>
  <label>{t('Font size')}<input class="ok-input size" aria-label={t('Font size')} type="number" min="1" max="4000" step="0.5" disabled={!(selected || typing)} value={size ?? ''} placeholder="—" onchange={(e) => { if (e.currentTarget.value && e.currentTarget.reportValidity()) onformat({size:e.currentTarget.valueAsNumber}); }} /></label>
  <label>{t('Text color')}<input aria-label={t('Text color')} type="color" value={color ?? '#000000'} title={color ?? t('Mixed or inherited')} disabled={!(selected || typing)} onchange={(e) => onformat({color:e.currentTarget.value})} /></label>
  <label>{t('Highlight color')}<input aria-label={t('Highlight color')} type="color" value={highlight ?? '#ffff00'} title={highlight ?? t('Mixed or inherited')} disabled={!(selected || typing)} onchange={e => onformat({ highlight: e.currentTarget.value })} /></label>
  <button class="ok-btn" disabled={!(selected || typing)} onmousedown={e => e.preventDefault()} onclick={() => onformat({ highlight: highlight ?? '#FFFF00' })}>{t('Apply highlight')}</button>
  <button class="ok-btn" disabled={!(selected || typing)} onmousedown={e => e.preventDefault()} onclick={() => onformat({ highlight: null })}>{t('Remove highlight')}</button>
  <button class="ok-btn" disabled={!(selected || typing)} onmousedown={e => e.preventDefault()} onclick={() => onformat({}, true)}>{t('Clear text formatting')}</button>
  {#if onparagraph && paragraph}
    <label>{t('Paragraph alignment')}<select aria-label={t('Paragraph alignment')} value={paragraph.align} onchange={e => onparagraph?.('align', e.currentTarget.value)}>
      <option value="" disabled>{t('Mixed or inherited')}</option>
      <option value="left">{t('Left')}</option><option value="center">{t('Center')}</option><option value="right">{t('Right')}</option><option value="justify">{t('Justify')}</option>
    </select></label>
    <label>{t('List style')}<select aria-label={t('List style')} value={paragraph.bullet} onchange={e => onparagraph?.('bullet', e.currentTarget.value)}>
      <option value="" disabled>{t('Mixed or inherited')}</option><option value="none">{t('No list')}</option><option value="bullet">{t('Bulleted list')}</option><option value="number">{t('Numbered list')}</option>
    </select></label>
    <label>{t('List level')}<select aria-label={t('List level')} title={t('Tab / Shift+Tab in lists; Ctrl/Cmd+[ / ] changes level')} value={paragraph.level} onchange={e => onparagraph?.('level', e.currentTarget.value)}>
      <option value="" disabled>{t('Mixed')}</option>
      {#each Array.from({ length: 9 }, (_, i) => i) as value}<option value={String(value)}>{value + 1}</option>{/each}
    </select></label>
    <label>{t('Line spacing mode')}<select aria-label={t('Line spacing mode')} value={paragraph.lineKind} onchange={e => onparagraph?.('lineKind', e.currentTarget.value)}>
      <option value="" disabled>{t('Mixed')}</option><option value="inherit">{t('Inherit')}</option><option value="pct">{t('Multiple')}</option><option value="pts">{t('Points')}</option>
    </select></label>
    {#if paragraph.lineKind === 'pct' || paragraph.lineKind === 'pts'}
      <label>{t('Line spacing value')}<input class="ok-input size" aria-label={t('Line spacing value')} type="number" min="0" step="0.01" required value={paragraph.lineValue} placeholder={t('Mixed')} onchange={e => { if (e.currentTarget.reportValidity()) onparagraph?.('lineValue', e.currentTarget.value); }} /></label>
    {/if}
    <label>{t('Before paragraph (pt)')}<input class="ok-input size" aria-label={t('Before paragraph (pt)')} type="number" min="0" step="0.01" value={paragraph.before} placeholder={t('Mixed or inherited')} onchange={e => { if (e.currentTarget.reportValidity()) onparagraph?.('before', e.currentTarget.value); }} /></label>
    <label>{t('After paragraph (pt)')}<input class="ok-input size" aria-label={t('After paragraph (pt)')} type="number" min="0" step="0.01" value={paragraph.after} placeholder={t('Mixed or inherited')} onchange={e => { if (e.currentTarget.reportValidity()) onparagraph?.('after', e.currentTarget.value); }} /></label>
  {/if}
  {#if onlink}<button class="ok-btn" disabled={!selected} onmousedown={(e) => e.preventDefault()} onclick={onlink}>{t('Edit link')}</button>{/if}
  {#if ondone}<button class="ok-btn" onclick={ondone}>{t('Done')}</button>{/if}
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
