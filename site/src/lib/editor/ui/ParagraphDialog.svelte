<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { cm, getParagraphPropertiesEffective, setParagraphAlignment, setParagraphIndent, setParagraphSpacing, setParagraphLineSpacing, setParagraphTypography, type ParagraphProperties } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  let { properties, apply, onclose }: { properties: ParagraphProperties[]; apply: (edit: (shape: Parameters<typeof setParagraphAlignment>[0], index: number) => void) => void; onclose: () => void } = $props();
  const doc = getEditor().doc;
  const initial = untrack(() => ({ properties, selection: doc.selection, version: doc.version }));
  function common<T>(read: (p: ParagraphProperties) => T): T | undefined {
    const values = initial.properties.map(read);
    return values.every(value => value === values[0]) ? values[0] : undefined;
  }
  const original = {
    asianLineBreak: common(p => p.asianLineBreak ?? true),
    latinLineBreak: common(p => p.latinLineBreak ?? false),
    hangingPunctuation: common(p => p.hangingPunctuation ?? true),
    fontAlignment: common(p => p.fontAlignment ?? 'auto'),
    alignment: common(p => p.align ?? 'left'),
    left: common(p => (p.marL ?? 0) / cm(1)),
    special: common(p => !p.indent ? 'none' : p.indent > 0 ? 'first' : 'hanging'),
    by: common(p => Math.abs(p.indent ?? 0) / cm(1)),
    before: common(p => p.spcBefPts ?? 0),
    after: common(p => p.spcAftPts ?? 0),
    line: common(p => p.lineSpacing?.kind === 'pts' ? 'exact' : !p.lineSpacing || p.lineSpacing.value === 1 ? 'single' : p.lineSpacing.value === 1.5 ? 'oneHalf' : p.lineSpacing.value === 2 ? 'double' : 'multiple'),
    at: common(p => p.lineSpacing?.value ?? 1),
  };
  let activeTab = $state<'spacing' | 'breaking'>('spacing');
  let asianLineBreak = $state(original.asianLineBreak);
  let latinLineBreak = $state(original.latinLineBreak);
  let hangingPunctuation = $state(original.hangingPunctuation);
  let fontAlignment = $state(original.fontAlignment);
  let alignment = $state(original.alignment);
  let left = $state(original.left);
  let special = $state(original.special);
  let by = $state(original.by);
  let before = $state(original.before);
  let after = $state(original.after);
  let line = $state(original.line);
  let at = $state(original.at);
  let dialog: HTMLDialogElement;
  onMount(() => dialog.showModal());
  $effect(() => { if (doc.selection !== initial.selection || doc.version !== initial.version) onclose(); });
  function submit(event: SubmitEvent) {
    event.preventDefault();
    const changed = asianLineBreak !== original.asianLineBreak || latinLineBreak !== original.latinLineBreak || hangingPunctuation !== original.hangingPunctuation || fontAlignment !== original.fontAlignment || alignment !== original.alignment || left !== original.left || special !== original.special || by !== original.by || before !== original.before || after !== original.after || line !== original.line || at !== original.at;
    if (changed) apply((shape, index) => {
      const typography: Parameters<typeof setParagraphTypography>[2] = {};
      if (asianLineBreak !== original.asianLineBreak) typography.asianLineBreak = asianLineBreak;
      if (latinLineBreak !== original.latinLineBreak) typography.latinLineBreak = latinLineBreak;
      if (hangingPunctuation !== original.hangingPunctuation) typography.hangingPunctuation = hangingPunctuation;
      if (fontAlignment !== original.fontAlignment) typography.fontAlignment = fontAlignment;
      if (Object.keys(typography).length) setParagraphTypography(shape, index, typography);
      if (alignment && alignment !== original.alignment) setParagraphAlignment(shape, index, alignment);
      if (left !== undefined && left !== original.left) setParagraphIndent(shape, index, { leftEmu: cm(left) });
      if (special && (special !== original.special || by !== original.by)) setParagraphIndent(shape, index, { firstLineEmu: special === 'none' ? 0 : (by === undefined ? Math.abs(getParagraphPropertiesEffective(doc.pres, shape, index).indent ?? 0) : cm(by)) * (special === 'hanging' ? -1 : 1) });
      if (before !== undefined && before !== original.before) setParagraphSpacing(shape, index, { beforePts: before });
      if (after !== undefined && after !== original.after) setParagraphSpacing(shape, index, { afterPts: after });
      if (line && (line !== original.line || at !== original.at)) {
        const value = line === 'single' ? 1 : line === 'oneHalf' ? 1.5 : line === 'double' ? 2 : at;
        if (value !== undefined) setParagraphLineSpacing(shape, index, { kind: line === 'exact' ? 'pts' : 'pct', value });
      }
    });
    onclose();
  }
</script>
<dialog bind:this={dialog} aria-label={t('Paragraph')} {onclose} onkeydown={event => event.stopPropagation()}>
  <form onsubmit={submit}>
    <h2>{t('Paragraph')}</h2>
    <div class="tabs" role="tablist" aria-label={t('Paragraph')}>
      {#each [{ id: 'spacing', label: 'Indents and Spacing' }, { id: 'breaking', label: 'Line Breaks and Alignment' }] as tab}
        <button type="button" role="tab" id={'paragraph-tab-' + tab.id} aria-controls={'paragraph-panel-' + tab.id} aria-selected={activeTab === tab.id} tabindex={activeTab === tab.id ? 0 : -1} onclick={() => activeTab = tab.id === 'spacing' ? 'spacing' : 'breaking'} onkeydown={event => { if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) { event.preventDefault(); activeTab = event.key === 'Home' ? 'spacing' : event.key === 'End' ? 'breaking' : activeTab === 'spacing' ? 'breaking' : 'spacing'; dialog.querySelector<HTMLButtonElement>('#paragraph-tab-' + activeTab)?.focus(); } }}>{t(tab.label)}</button>
      {/each}
    </div>
    <div role="tabpanel" id="paragraph-panel-spacing" aria-labelledby="paragraph-tab-spacing" hidden={activeTab !== 'spacing'}>
    <fieldset><legend>{t('General')}</legend>
      <label>{t('Alignment:')}<select aria-label={t('Alignment:')} bind:value={alignment}><option value={undefined} disabled>{t('Mixed')}</option>{#each [{ value: 'left', label: 'Left' }, { value: 'center', label: 'Center' }, { value: 'right', label: 'Right' }, { value: 'justify', label: 'Justify' }, { value: 'distribute', label: 'Distributed' }] as item}<option value={item.value}>{t(item.label)}</option>{/each}</select></label>
    </fieldset>
    <fieldset><legend>{t('Indentation')}</legend>
      <label>{t('Before text:')}<span><input aria-label={t('Before text:')} type="number" min="0" max="142.24" step="any" bind:value={left} placeholder={t('Mixed')} /> cm</span></label>
      <div class="row"><label>{t('Special:')}<select aria-label={t('Special:')} bind:value={special}><option value={undefined} disabled>{t('Mixed')}</option><option value="none">{t('(None)')}</option><option value="first">{t('First line')}</option><option value="hanging">{t('Hanging')}</option></select></label><label>{t('By:')}<span><input aria-label={t('By:')} type="number" min="0" max="142.24" step="any" bind:value={by} disabled={!special || special === 'none'} placeholder={t('Mixed')} /> cm</span></label></div>
    </fieldset>
    <fieldset><legend>{t('Spacing')}</legend>
      <div class="row"><div><label>{t('Before:')}<span><input aria-label={t('Before:')} type="number" min="0" max="1584" step="any" bind:value={before} placeholder={t('Mixed')} /> pt</span></label><label>{t('After:')}<span><input aria-label={t('After:')} type="number" min="0" max="1584" step="any" bind:value={after} placeholder={t('Mixed')} /> pt</span></label></div><div><label>{t('Line spacing:')}<select aria-label={t('Line spacing:')} bind:value={line} onchange={() => { if (line === 'exact') at = 18; if (line === 'multiple') at = 1; }}><option value={undefined} disabled>{t('Mixed')}</option><option value="single">{t('Single')}</option><option value="oneHalf">{t('1.5 lines')}</option><option value="double">{t('Double')}</option><option value="exact">{t('Exactly')}</option><option value="multiple">{t('Multiple')}</option></select></label><label>{t('At:')}<input type="number" min="0" max={line === 'exact' ? 1584 : 9.99} step="any" bind:value={at} disabled={line !== 'exact' && line !== 'multiple'} placeholder={t('Mixed')} /></label></div></div>
    </fieldset>
    </div>
    <div role="tabpanel" id="paragraph-panel-breaking" aria-labelledby="paragraph-tab-breaking" hidden={activeTab !== 'breaking'}>
      <fieldset><legend>{t('Line Breaking:')}</legend>
        <label class="check"><input type="checkbox" checked={asianLineBreak === true} indeterminate={asianLineBreak === undefined} onchange={event => asianLineBreak = event.currentTarget.checked} />{t('Use Asian typography rules')}</label>
        <label class="check"><input type="checkbox" checked={latinLineBreak === true} indeterminate={latinLineBreak === undefined} onchange={event => latinLineBreak = event.currentTarget.checked} />{t('Allow Latin text to wrap in the middle of a word')}</label>
        <label class="check"><input type="checkbox" checked={hangingPunctuation === true} indeterminate={hangingPunctuation === undefined} onchange={event => hangingPunctuation = event.currentTarget.checked} />{t('Allow hanging punctuation')}</label>
      </fieldset>
      <fieldset><legend>{t('Alignment:')}</legend>
        <label>{t('Text Alignment:')}<select aria-label={t('Text Alignment:')} bind:value={fontAlignment}><option value={undefined} disabled>{t('Mixed')}</option>{#each [{ value: 'top', label: 'Top' }, { value: 'center', label: 'Center' }, { value: 'baseline', label: 'Baseline' }, { value: 'bottom', label: 'Bottom' }, { value: 'auto', label: 'Auto' }] as item}<option value={item.value}>{t(item.label)}</option>{/each}</select></label>
      </fieldset>
    </div>
    <footer><button type="button" onclick={onclose}>{t('Cancel')}</button><button type="submit">{t('OK')}</button></footer>
  </form>
</dialog>
<style>
  dialog { width: 490px; border: 1px solid #666; border-radius: 7px; background: #303030; color: #eee; padding: 0; box-shadow: 0 15px 60px #0008; font-size: 13px; color-scheme: dark; }
  dialog::backdrop { background: #0003; }
  h2 { font-size: 13px; text-align: center; margin: 0; padding: 8px; background: #3a3a3a; }
  .tabs { display: flex; justify-content: center; margin: 14px auto; }
  .tabs button { background: #454545; color: inherit; font: inherit; border: 1px solid #777; padding: 4px 10px; }
  .tabs button:first-child { border-radius: 4px 0 0 4px; }
  .tabs button:last-child { border-radius: 0 4px 4px 0; }
  .tabs button[aria-selected='true'] { background: #626262; }
  [role='tabpanel'] { min-height: 292px; }
  .check { justify-content: flex-start; }
  .check input { width: auto; }
  fieldset { border: 0; border-top: 1px solid #666; margin: 10px 18px; padding: 8px 0; }
  legend { padding-right: 8px; }
  label { display: flex; align-items: center; justify-content: space-between; gap: 8px; margin: 8px 0; }
  .row { display: flex; justify-content: space-between; gap: 20px; }
  input { width: 65px; }
  input, select { padding: 2px 4px; border: 1px solid #777; border-radius: 4px; background: #252525; color: inherit; font: inherit; }
  input:disabled { opacity: .4; }
  footer { display: flex; justify-content: flex-end; gap: 10px; padding: 4px 18px 16px; }
  footer button { min-width: 70px; padding: 3px 8px; border: 0; border-radius: 5px; background: #626262; color: inherit; font: inherit; }
  footer button[type='submit'] { background: #087bfa; color: white; }
</style>
