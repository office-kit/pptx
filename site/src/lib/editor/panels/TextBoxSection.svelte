<script lang="ts">
  import {
    cm, getShapeKind, getShapeBodyPrEffective, getShapeTextAutoFit,
    setShapeTextAnchor, setShapeTextDirection, setShapeTextAutoFit, setShapeTextMargins, setShapeTextWrap,
    type SlideShapeData, type TextAnchor, type TextAutoFit,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { shapeTextDefaults } from '../core/text-layout-defaults.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import TextColumnsDialog from '../ui/TextColumnsDialog.svelte';

  const editor = getEditor();
  const doc = editor.doc;
  let columnsOpen = $state(false);
  const shapes = $derived.by(() => {
    doc.version;
    const selected = editor.selectedShapes();
    return selected.every(shape => getShapeKind(shape) === 'shape') ? selected : [];
  });
  const values = $derived(shapes.map(shape => ({ shape, body: getShapeBodyPrEffective(doc.pres, shape), fit: getShapeTextAutoFit(shape) })));
  function common<T>(items: T[]): T | undefined { return items.every(item => item === items[0]) ? items[0] : undefined; }
  const anchor = $derived(common(values.map(item => item.body.anchor ?? shapeTextDefaults(item.shape).anchor)));
  const direction = $derived(common(values.map(item => item.body.vert ?? 'horz')));
  const autoFit = $derived(common(values.map(item => item.fit)));
  const wrap = $derived(common(values.map(item => (item.body.wrap ?? 'square') === 'square')));
  const sides = [['left', 'Left margin'], ['right', 'Right margin'], ['top', 'Top margin'], ['bottom', 'Bottom margin']] as const;
  type Side = typeof sides[number][0];
  const margins = $derived.by(() => {
    const value = (side: Side) => {
      // DrawingML's default insets: 0.1 inch horizontally, 0.05 inch vertically.
      const fallback = side === 'left' || side === 'right' ? 91440 : 45720;
      const result = common(values.map(item => item.body.margins[side] ?? fallback));
      return result === undefined ? undefined : Math.round(result / cm(1) * 100) / 100;
    };
    return { left: value('left'), right: value('right'), top: value('top'), bottom: value('bottom') };
  });
  const directions = [
    ['horz', 'Horizontal'], ['eaVert', 'Vertical'], ['vert', 'Rotate all text 90°'],
    ['vert270', 'Rotate all text 270°'], ['wordArtVert', 'Stacked'],
  ] as const;
  function apply(label: string, operation: (shape: SlideShapeData) => void) {
    if (shapes.length) doc.transact(t(label), () => { for (const shape of shapes) operation(shape); });
  }
  function changeAnchor(value: string) {
    if (value === 'top' || value === 'center' || value === 'bottom') apply('Vertical alignment', shape => setShapeTextAnchor(shape, value));
  }
  function changeDirection(value: string) {
    const option = directions.find(([key]) => key === value);
    if (option) apply('Text direction', shape => setShapeTextDirection(shape, option[0]));
  }
  function margin(side: Side, input: HTMLInputElement) {
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) { input.value = margins[side] === undefined ? '' : String(margins[side]); return; }
    const value = cm(input.valueAsNumber);
    apply('Text margins', shape => setShapeTextMargins(shape, { [side]: value }));
  }
  const anchors: ReadonlyArray<readonly [TextAnchor, string]> = [['top', 'Top'], ['center', 'Middle'], ['bottom', 'Bottom']];
  const fits: ReadonlyArray<readonly [TextAutoFit, string]> = [['none', 'Do not Autofit'], ['normal', 'Shrink text on overflow'], ['shape', 'Resize shape to fit text']];
</script>

{#if shapes.length}
  <details class="text-box">
    <summary>{t('Text Box')}</summary>
    <div class="fields">
      <label><span>{t('Vertical alignment')}</span><select class="ok-input" value={anchor ?? ''} onchange={e => changeAnchor(e.currentTarget.value)}>
        <option value="" disabled>{t('Mixed')}</option>
        {#each anchors as [value, label]}<option {value}>{t(label)}</option>{/each}
      </select></label>
      <label><span>{t('Text direction')}</span><select class="ok-input" value={direction ?? ''} onchange={e => changeDirection(e.currentTarget.value)}>
        <option value="" disabled>{t('Mixed')}</option>
        {#each directions as [value, label]}<option {value}>{t(label)}</option>{/each}
        {#if direction && !directions.some(([value]) => value === direction)}<option value={direction}>{t('Custom')}</option>{/if}
      </select></label>
      <div role="radiogroup" aria-label={t('Autofit')}>
        {#each fits as [value, label]}
          <label class="check"><input type="radio" name="text-autofit" checked={autoFit === value} onchange={() => apply(label, shape => setShapeTextAutoFit(shape, value))} /><span>{t(label)}</span></label>
        {/each}
      </div>
      {#each sides as [side, label]}
        <label><span>{t(label)}</span><span class="number"><input class="ok-input" aria-label={t(label)} type="number" min="0" max="55.88" step="any" value={margins[side] ?? ''} placeholder={margins[side] === undefined ? t('Mixed') : undefined} onchange={e => margin(side, e.currentTarget)} /><span>cm</span></span></label>
      {/each}
      <label class="check"><input type="checkbox" checked={wrap ?? false} indeterminate={wrap === undefined} onchange={e => { const value = e.currentTarget.checked ? 'square' : 'none'; apply('Wrap text in shape', shape => setShapeTextWrap(shape, value)); }} /><span>{t('Wrap text in shape')}</span></label>
      <button class="ok-btn columns" onclick={() => columnsOpen = true}>{t('Columns...')}</button>
    </div>
  </details>
{/if}
{#if columnsOpen}<TextColumnsDialog onclose={() => columnsOpen = false} />{/if}

<style>
  .text-box { font-size: 11px; margin: 0 -10px; }
  summary { padding: 4px 8px; background: var(--ok-hover); cursor: pointer; }
  .fields { display: flex; flex-direction: column; gap: 8px; padding: 12px; }
  label { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-height: 20px; }
  select { max-width: 145px; font-size: inherit; }
  .number { display: flex; align-items: center; gap: 3px; width: 96px; flex: 0 0 96px; }
  .number input { width: 72px; min-width: 0; padding: 2px 4px; font-size: inherit; }
  .check { justify-content: flex-start; }
  .check input { margin: 0; }
  .columns { align-self: flex-start; }
</style>
