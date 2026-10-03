<script lang="ts">
  import {
    getShapeKind, getShapeText, getTableCells, getTableCellText,
    getPresentationFonts, setTableCellTextFormat,
    type TextFormat,
  } from '@office-kit/pptx';
  import { onMount } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { textFormatsInRange } from '../core/text-format-selection.ts';
  import { tableCellsInRange, tableSelectionBlock } from '../core/table-selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import FontFamilyInput from './FontFamilyInput.svelte';
  import ColorPicker from './ColorPicker.svelte';

  const editor = getEditor();
  const doc = editor.doc;
  let dialog: HTMLDialogElement;
  let tab = $state<'font' | 'character'>('font');
  let family = $state('');
  let familyEastAsian = $state('');
  let style = $state('');
  let size = $state<number>();
  let underline = $state('');
  let strike = $state<string>();
  let cap = $state<TextFormat['cap']>();
  let offset = $state<number>();
  let normalizeHeight = $state<boolean>();
  let spacingMode = $state('');
  let spacingAmount = $state<number>();
  let useKerning = $state<boolean>();
  let kerningThreshold = $state<number>();
  let patch = $state<TextFormat>({});
  const formats = $derived.by(() => {
    doc.version;
    if (editor.inlineTextFormat) return editor.inlineTextFormat.formats;
    const selection = doc.selection;
    if (selection.kind === 'cell') {
      const table = doc.shapeById(selection.slideIndex, selection.shapeId);
      if (!table) return [];
      const tableCells = getTableCells(table);
      const block = tableSelectionBlock(selection);
      const result: TextFormat[] = [];
      for (let row = block.row; row < block.row + block.rowSpan; row++) {
        for (let col = block.col; col < block.col + block.colSpan; col++) {
          const cell = tableCells[row]?.[col];
          if (cell) result.push(...textFormatsInRange(table, { start: 0, end: getTableCellText(cell).length }, { row, col }, { pres: doc.pres }));
        }
      }
      return result;
    }
    return editor.selectedShapes().filter(shape => getShapeKind(shape) === 'shape').flatMap(shape =>
      textFormatsInRange(shape, { start: 0, end: getShapeText(shape).length }, undefined, { pres: doc.pres }),
    );
  });
  const common = <T,>(read: (format: TextFormat) => T): T | undefined => {
    const values = formats.map(read);
    return values.length && values.every(value => value === values[0]) ? values[0] : undefined;
  };
  const themeFonts = $derived.by(() => {
    doc.version;
    const fonts = getPresentationFonts(doc.pres);
    return fonts ? [fonts.majorLatin, fonts.majorEastAsian, fonts.majorComplexScript, fonts.minorLatin, fonts.minorEastAsian, fonts.minorComplexScript].filter((font): font is string => !!font) : [];
  });
  const families = $derived([...new Set([...themeFonts, ...formats.flatMap(f => [f.font, f.fontEastAsian, f.fontComplexScript].filter((font): font is string => !!font))])]);

  function readState() {
    family = common(f => f.font) ?? '';
    familyEastAsian = common(f => f.fontEastAsian) ?? '';
    size = common(f => f.size);
    style = common(f => `${f.bold ?? false}:${f.italic ?? false}`) ?? '';
    underline = common(f => f.underline === true ? 'sng' : !f.underline ? 'none' : f.underline) ?? '';
    strike = common(f => f.strike === true ? 'sngStrike' : !f.strike ? 'noStrike' : f.strike);
    cap = common(f => f.cap ?? 'none');
    const base = common(f => f.baseline ?? 0);
    offset = base === undefined ? undefined : base * 100;
    normalizeHeight = common(f => f.normalizeHeight ?? false);
    const spc = common(f => f.spc ?? 0);
    spacingMode = spc === undefined ? '' : spc === 0 ? 'normal' : spc > 0 ? 'expanded' : 'condensed';
    spacingAmount = spc === undefined || spc === 0 ? undefined : Math.abs(spc) / 100;
    const kern = common(f => f.kern ?? 1200);
    useKerning = kern === undefined ? undefined : kern !== 0;
    kerningThreshold = kern === undefined ? undefined : kern === 0 ? 12 : kern / 100;
  }
  function setSpacing() {
    if (spacingMode === 'normal') patch.spc = 0;
    else if (spacingMode && spacingAmount !== undefined) patch.spc = Math.round(spacingAmount * 100) * (spacingMode === 'condensed' ? -1 : 1);
  }
  function setSpacingAmount() {
    if (spacingMode === 'normal' && spacingAmount !== undefined && spacingAmount > 0) spacingMode = 'expanded';
    setSpacing();
  }
  function setKerning() {
    if (useKerning === false) patch.kern = 0;
    else if (useKerning && kerningThreshold !== undefined) patch.kern = Math.round(kerningThreshold * 100);
  }
  function setOffset(value: number | undefined) {
    offset = value;
    if (value !== undefined) patch.baseline = value / 100;
  }
  function apply() {
    if (Object.keys(patch).length) {
      if (editor.inlineTextFormat) editor.inlineTextFormat.apply(patch);
      else if (doc.selection.kind === 'cell') {
        const table = doc.shapeById(doc.selection.slideIndex, doc.selection.shapeId);
        if (table) {
          const cells = [...tableCellsInRange(getTableCells(table), tableSelectionBlock(doc.selection))];
          doc.transact(t('Format selected cells'), () => { for (const cell of cells) setTableCellTextFormat(cell, patch); });
        }
      } else editor.invoke('setShapeTextFormat', { format: patch, options: {} });
    }
    editor.closeDialog();
  }
  onMount(() => { tab = editor.fontDialogTab; readState(); dialog.showModal(); });
</script>

<dialog class="font-dialog" bind:this={dialog} aria-label={t('Font')} oncancel={() => editor.closeDialog()} onclose={() => editor.closeDialog()} onkeydown={event => event.stopPropagation()}>
  <form onsubmit={event => { event.preventDefault(); apply(); }}>
    <div class="tabs" role="tablist" aria-label={t('Font')}>
      <button type="button" role="tab" aria-selected={tab === 'font'} onclick={() => tab = 'font'}>{t('Font')}</button>
      <button type="button" role="tab" aria-selected={tab === 'character'} onclick={() => tab = 'character'}>{t('Character Spacing')}</button>
    </div>
    <div class="panel">
    {#if tab === 'font'}
      <div class="font-grid">
        <div class="family"><span>{t('Latin text font')}</span><FontFamilyInput label={t('Latin text font')} value={family} {families} choose={font => { family = font; patch.font = font; }} /></div>
        <label>{t('Font style')}<select bind:value={style} onchange={() => { if (style) { patch.bold = style.startsWith('true:'); patch.italic = style.endsWith(':true'); } }}><option value="">{t('Mixed')}</option><option value="false:false">{t('Regular')}</option><option value="false:true">{t('Italic')}</option><option value="true:false">{t('Bold')}</option><option value="true:true">{t('Bold Italic')}</option></select></label>
        <label>{t('Font size')}<input type="number" min="1" max="4000" step="any" bind:value={size} placeholder={t('Mixed')} onchange={() => { if (size !== undefined) patch.size = size; }} /></label>
        <div class="family"><span>{t('Asian text font')}</span><FontFamilyInput label={t('Asian text font')} value={familyEastAsian} {families} choose={font => { familyEastAsian = font; patch.fontEastAsian = font; }} /></div>
      </div>
      <fieldset><legend>{t('Color & Underline')}</legend><div class="color-grid">
        <div><span>{t('Font color')}</span><ColorPicker label={t('Font color')} value={patch.color ?? common(f => f.color) ?? undefined} choose={color => { patch.color = color; }} /></div>
        <label>{t('Underline style')}<select bind:value={underline} onchange={() => { if (underline) patch.underline = underline === 'none' ? false : underline; }}>
          <option value="">{t('Mixed')}</option>
          <option value="none">{t('(None)')}</option>
          <option value="words">{t('Words only')}</option>
          <option value="sng">{t('Single')}</option>
          <option value="dbl">{t('Double')}</option>
          <option value="heavy">{t('Heavy')}</option>
          <option value="dotted">{t('Dotted')}</option>
          <option value="dottedHeavy">{t('Dotted heavy')}</option>
          <option value="dash">{t('Dashed')}</option>
          <option value="dashHeavy">{t('Dashed heavy')}</option>
          <option value="dashLong">{t('Long dashed')}</option>
          <option value="dashLongHeavy">{t('Long dashed heavy')}</option>
          <option value="dotDash">{t('Dash dot')}</option>
          <option value="dotDashHeavy">{t('Dash dot heavy')}</option>
          <option value="dotDotDash">{t('Dash dot dot')}</option>
          <option value="dotDotDashHeavy">{t('Dash dot dot heavy')}</option>
          <option value="wavy">{t('Wavy')}</option>
          <option value="wavyHeavy">{t('Wavy heavy')}</option>
          <option value="wavyDbl">{t('Double wavy')}</option>
        </select></label>
        <div><span>{t('Underline color')}</span><ColorPicker label={t('Underline color')} disabled={underline === 'none' || !underline} value={(patch.underlineColor !== undefined ? patch.underlineColor : common(f => f.underlineColor)) ?? undefined} automaticSelected={(patch.underlineColor !== undefined ? patch.underlineColor : common(f => f.underlineColor ?? null)) === null} automatic={() => { patch.underlineColor = null; }} choose={color => { patch.underlineColor = color; }} /></div>
      </div></fieldset>
      <fieldset><legend>{t('Effects')}</legend><div class="effects-grid">
        <div>
          <label class="check"><input type="checkbox" checked={strike === 'sngStrike'} indeterminate={strike === undefined} onchange={e => { strike = e.currentTarget.checked ? 'sngStrike' : 'noStrike'; patch.strike = strike; }} />{t('Strikethrough')}</label>
          <label class="check"><input type="checkbox" checked={strike === 'dblStrike'} indeterminate={strike === undefined} onchange={e => { strike = e.currentTarget.checked ? 'dblStrike' : 'noStrike'; patch.strike = strike; }} />{t('Double strikethrough')}</label>
          <label class="check"><input type="checkbox" checked={offset !== undefined && offset > 0} indeterminate={offset === undefined} onchange={e => setOffset(e.currentTarget.checked ? 30 : 0)} />{t('Superscript')}</label>
          <label class="check"><input type="checkbox" checked={offset !== undefined && offset < 0} indeterminate={offset === undefined} onchange={e => setOffset(e.currentTarget.checked ? -25 : 0)} />{t('Subscript')}</label>
          <label class="offset">{t('Offset (%)')}<input type="number" min="-100" max="100" step="any" bind:value={offset} placeholder={t('Mixed')} onchange={() => setOffset(offset)} /></label>
        </div>
        <div>
          <label class="check"><input type="checkbox" checked={cap === 'small'} indeterminate={cap === undefined} onchange={e => { cap = e.currentTarget.checked ? 'small' : 'none'; patch.cap = cap; }} />{t('Small caps')}</label>
          <label class="check"><input type="checkbox" checked={cap === 'all'} indeterminate={cap === undefined} onchange={e => { cap = e.currentTarget.checked ? 'all' : 'none'; patch.cap = cap; }} />{t('All caps')}</label>
          <label class="check"><input type="checkbox" checked={normalizeHeight ?? false} indeterminate={normalizeHeight === undefined} onchange={e => { normalizeHeight = e.currentTarget.checked; patch.normalizeHeight = normalizeHeight; }} />{t('Equalize character height')}</label>
        </div>
      </div></fieldset>
    {:else}
      <div class="spacing-row">
        <label>{t('Character spacing mode')}<select aria-label={t('Spacing')} bind:value={spacingMode} onchange={() => { if (spacingMode === 'normal') spacingAmount = undefined; else if (spacingAmount === undefined || spacingAmount === 0) spacingAmount = 1; setSpacing(); }}><option value="">{t('Mixed')}</option><option value="normal">{t('Normal')}</option><option value="expanded">{t('Expanded')}</option><option value="condensed">{t('Condensed')}</option></select></label>
        <label>{t('By (pt)')}<input type="number" min="0" max="1000" step="any" disabled={!spacingMode} bind:value={spacingAmount} onchange={setSpacingAmount} /></label>
      </div>
      <div class="kerning-row">
        <label class="check"><input type="checkbox" checked={useKerning ?? false} indeterminate={useKerning === undefined} onchange={e => { useKerning = e.currentTarget.checked; kerningThreshold ??= 12; setKerning(); }} />{t('Use kerning for fonts')}</label>
        <label>{t('For fonts')}<input type="number" min="1" max="1000" step="any" disabled={!useKerning} bind:value={kerningThreshold} onchange={setKerning} /><span>{t('points and above')}</span></label>
      </div>
    {/if}
    </div>
    <footer><button type="button" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit">{t('OK')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: 574px; max-width: calc(100vw - 32px); box-sizing: border-box; border: 1px solid var(--ok-border); border-radius: 7px; background: var(--ok-panel); color: var(--ok-text); padding: 0; box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0005; }
  .tabs { display: flex; justify-content: center; padding: 18px 14px 6px; }
  button, input, select { font: inherit; }
  .tabs button { padding: 5px 12px; border: 1px solid var(--ok-border); background: transparent; color: inherit; cursor: pointer; }
  .tabs button:first-child { border-radius: 5px 0 0 5px; }
  .tabs button:last-child { border-radius: 0 5px 5px 0; }
  .tabs button[aria-selected='true'] { background: var(--ok-accent); color: white; }
  .panel { min-height: 360px; padding: 12px 24px 0; font-size: 12px; }
  .font-grid { display: grid; grid-template-columns: minmax(0, 2fr) minmax(0, 1fr) minmax(0, .8fr); gap: 12px; }
  label, .family, .color-grid > div { display: flex; flex-direction: column; gap: 5px; min-width: 0; }
  .family :global(.font-field) { width: 100%; }
  .family :global(.font-field input) { min-width: 0; width: 100%; }
  .check { flex-direction: row; align-items: center; gap: 7px; margin: 9px 0; }
  input, select { box-sizing: border-box; padding: 4px 6px; border: 1px solid var(--ok-border); border-radius: 4px; background: var(--ok-input); color: inherit; min-width: 0; width: 100%; }
  input[type='checkbox'] { width: auto; margin: 0; }
  input:disabled { opacity: .45; }
  fieldset { border: 0; border-top: 1px solid var(--ok-border); margin: 18px 0 0; padding: 10px 0 0; }
  legend { padding-right: 8px; }
  .color-grid, .effects-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 22px; }
  .color-grid { grid-template-columns: 1fr 1.4fr 1fr; }
  .offset { flex-direction: row; align-items: center; margin-top: 8px; }
  .offset input { width: 78px; }
  .spacing-row { display: flex; gap: 24px; margin: 22px 0; }
  .spacing-row label { width: 150px; }
  .kerning-row > label:last-child { flex-direction: row; align-items: center; margin-top: 14px; }
  .kerning-row input[type='number'] { width: 80px; }
  footer { display: flex; justify-content: flex-end; gap: 8px; padding: 18px 24px; }
  footer button { min-width: 70px; padding: 5px 10px; border: 1px solid var(--ok-border); border-radius: 4px; background: transparent; color: inherit; cursor: pointer; }
  footer button[type='submit'] { background: var(--ok-accent); border-color: var(--ok-accent); color: #fff; }
</style>
