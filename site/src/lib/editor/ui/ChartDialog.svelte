<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { addSlideChart, setChartSpec, getShapeId, getSlideCharts, emu, type ChartKind, type ChartSeries, type ChartSpec } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { selectedShapeId } from '../core/selection.ts';
  import { slideMetrics } from '../canvas/geometry.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { edit = false }: { edit?: boolean } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const initial = untrack(() => {
    const id = selectedShapeId(doc.selection);
    return edit && doc.currentSlide ? getSlideCharts(doc.currentSlide).find(chart => getShapeId(chart.shape) === id) : undefined;
  });
  const original = initial?.spec;
  const presentation = doc.pres;
  const version = doc.version;
  const kinds: { value: ChartKind; label: string }[] = [
    { value: 'column', label: 'Column chart' }, { value: 'bar', label: 'Bar chart' },
    { value: 'line', label: 'Line chart' }, { value: 'area', label: 'Area chart' },
    { value: 'pie', label: 'Pie chart' }, { value: 'doughnut', label: 'Doughnut chart' },
    { value: 'radar', label: 'Radar chart' },
  ];
  const markerSymbols: { value: NonNullable<ChartSeries['markerSymbol']>; label: string }[] = [
    { value: 'auto', label: 'Automatic' }, { value: 'none', label: 'None' },
    { value: 'circle', label: 'Circle' }, { value: 'square', label: 'Square' },
    { value: 'diamond', label: 'Diamond' }, { value: 'triangle', label: 'Triangle' },
    { value: 'star', label: 'Star' }, { value: 'x', label: 'Cross' },
    { value: 'plus', label: 'Plus' }, { value: 'dash', label: 'Dash' }, { value: 'dot', label: 'Dot' },
  ];
  let dialog: HTMLDialogElement;
  let kind = $state<ChartKind>(original?.kind ?? 'column');
  let title = $state(original?.title ?? '');
  let stacking = $state<'none' | 'stacked' | 'percentStacked'>(original?.grouping === 'stacked' || original?.grouping === 'percentStacked' ? original.grouping : 'none');
  let stackingChanged = $state(false);
  let legendPosition = $state<NonNullable<ChartSpec['legend']>['position']>(original?.legend?.position ?? null);
  let legendChanged = $state(false);
  let showValue = $state(original?.dataLabels?.showValue ?? false);
  let showCategory = $state(original?.dataLabels?.showCategory ?? false);
  let showSeriesName = $state(original?.dataLabels?.showSeriesName ?? false);
  let showPercent = $state(original?.dataLabels?.showPercent ?? false);
  let labelsChanged = $state(false);
  let categoryAxisTitle = $state(original?.categoryAxisTitle ?? '');
  let valueAxisTitle = $state(original?.valueAxisTitle ?? '');
  let axisMin = $state<number | null | undefined>(original?.valueAxis?.min);
  let axisMax = $state<number | null | undefined>(original?.valueAxis?.max);
  let majorUnit = $state<number | null | undefined>(original?.valueAxis?.majorUnit);
  let minorUnit = $state<number | null | undefined>(original?.valueAxis?.minorUnit);
  let numberFormat = $state(original?.valueAxis?.numberFormat ?? '');
  let axesChanged = $state(false);
  let showCategoryAxis = $state(!(original?.categoryAxisHidden ?? false));
  let showValueAxis = $state(!(original?.valueAxisHidden ?? false));
  let showMajorGridlines = $state(original?.valueAxisMajorGridlines ?? false);
  let showMinorGridlines = $state(original?.valueAxisMinorGridlines ?? false);
  let axisVisibilityChanged = $state(false);
  const hasAxes = $derived(['column', 'bar', 'line', 'area'].includes(kind));
  const validAxes = $derived(!hasAxes || !axesChanged || (
    [axisMin, axisMax, majorUnit, minorUnit].every(value => value == null || Number.isFinite(value)) &&
    (axisMin == null || axisMax == null || axisMin < axisMax) &&
    (majorUnit == null || majorUnit > 0) && (minorUnit == null || minorUnit > 0) &&
    (!original?.valueAxis?.logBase || ((axisMin == null || axisMin > 0) && (axisMax == null || axisMax > 0)))
  ));
  let categories = $state([...(original?.categories ?? [t('Category') + ' 1', t('Category') + ' 2', t('Category') + ' 3'])]);
  type SeriesDraft = { base: ChartSeries; name: string; color?: string; values: (number | undefined)[] };
  // Match the default series colors emitted by the ChartML writer.
  const defaultSeriesColors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
  const initialSeries = original?.series ?? [{ name: t('Series') + ' 1', values: [10, 20, 15], color: defaultSeriesColors[0] }];
  let series = $state<SeriesDraft[]>(initialSeries.map(base => ({ base, name: base.name, color: base.color, values: base.values.map(value => value ?? undefined) })));
  const validMarkers = $derived(kind !== 'line' || series.every(entry => entry.base.markerSizePt == null || (Number.isInteger(entry.base.markerSizePt) && entry.base.markerSizePt >= 2 && entry.base.markerSizePt <= 72)));
  let error = $state('');
  const supported = $derived(!edit || (original && kinds.some(item => item.value === original.kind)));
  const singleSeries = $derived(kind === 'pie' || kind === 'doughnut');
  const validSeries = $derived(!singleSeries || series.length === 1);
  const heading = $derived(t(edit ? 'Edit chart' : 'Insert chart'));
  onMount(() => dialog.showModal());

  function addCategory() {
    categories.push(t('Category') + ' ' + (categories.length + 1));
    for (const entry of series) entry.values.push(0);
  }
  function removeCategory(index: number) {
    categories.splice(index, 1);
    for (const entry of series) {
      entry.values.splice(index, 1);
      entry.base = {
        ...entry.base,
        pointColors: entry.base.pointColors?.filter((_, i) => i !== index),
        pointExplosions: entry.base.pointExplosions?.filter((_, i) => i !== index),
        pointDataLabels: entry.base.pointDataLabels?.filter((_, i) => i !== index),
      };
    }
  }
  function addSeries() {
    series.push({ base: { name: '', values: [] }, name: t('Series') + ' ' + (series.length + 1), values: categories.map(() => 0) });
  }
  function submit(event: SubmitEvent) {
    event.preventDefault();
    if (!supported || !validSeries || !validAxes || !validMarkers || !doc.currentSlide) return;
    if (doc.pres !== presentation || doc.version !== version) { error = t('The document changed. Reopen the chart editor.'); return; }
    const spec: ChartSpec = {
      ...original,
      ...(hasAxes && stackingChanged ? {
        grouping: stacking === 'none' ? (kind === 'column' || kind === 'bar' ? 'clustered' : 'standard') : stacking,
        overlapPct: undefined,
      } : {}),
      ...(hasAxes && axesChanged ? {
        categoryAxisTitle: categoryAxisTitle || undefined,
        valueAxisTitle: valueAxisTitle || undefined,
        valueAxis: { ...original?.valueAxis, min: axisMin ?? undefined, max: axisMax ?? undefined, majorUnit: majorUnit ?? undefined, minorUnit: minorUnit ?? undefined, numberFormat: numberFormat || undefined },
      } : {}),
      ...(hasAxes && axisVisibilityChanged ? {
        categoryAxisHidden: !showCategoryAxis,
        valueAxisHidden: !showValueAxis,
        valueAxisMajorGridlines: showMajorGridlines,
        valueAxisMinorGridlines: showMinorGridlines,
      } : {}),
      kind, title: title || undefined, categories: [...categories],
      legend: legendChanged ? { ...original?.legend, position: legendPosition, layout: undefined } : original?.legend,
      dataLabels: labelsChanged ? { ...original?.dataLabels, showValue, showCategory, showSeriesName, showPercent } : original?.dataLabels,
      series: series.map(entry => ({
        ...entry.base, name: entry.name, color: entry.color, values: entry.values.map(value => value ?? null),
        dataLabels: labelsChanged && entry.base.dataLabels ? { ...entry.base.dataLabels, showValue, showCategory, showSeriesName, showPercent } : entry.base.dataLabels,
        pointDataLabels: labelsChanged ? entry.base.pointDataLabels?.map(label => label ? { ...label, showValue, showCategory, showSeriesName, showPercent } : null) : entry.base.pointDataLabels,
      })),
    };
    try {
      doc.transact(heading, () => {
        if (edit && initial) setChartSpec(initial, spec);
        else {
          const size = slideMetrics(doc.pres);
          const shape = addSlideChart(doc.currentSlide!, { x: emu(Math.round(size.widthEmu * 0.1)), y: emu(Math.round(size.heightEmu * 0.1)), w: emu(Math.round(size.widthEmu * 0.8)), h: emu(Math.round(size.heightEmu * 0.8)), spec });
          doc.selectShape(doc.selection.slideIndex, getShapeId(shape));
        }
      });
      editor.closeDialog();
    } catch (cause) {
      error = `${t('The chart could not be updated')}: ${cause instanceof Error ? cause.message : String(cause)}`;
    }
  }
</script>

<dialog bind:this={dialog} aria-label={heading} onclose={() => editor.closeDialog()}>
  <form onsubmit={submit}>
    <header><strong>{heading}</strong><button class="ok-btn" type="button" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    {#if supported}
      <div class="settings">
        <label>{t('Chart type')}<select class="ok-input" aria-label={t('Chart type')} bind:value={kind}>{#each kinds as item}<option value={item.value}>{t(item.label)}</option>{/each}</select></label>
        <label>{t('Chart title')}<input class="ok-input" bind:value={title} /></label>
        {#if hasAxes}
          <label>{t('Series stacking')}<select class="ok-input" aria-label={t('Series stacking')} bind:value={stacking} onchange={() => stackingChanged = true}>
            <option value="none">{t('No stacking')}</option><option value="stacked">{t('Stacked')}</option><option value="percentStacked">{t('100% stacked')}</option>
          </select></label>
        {/if}
      </div>
      <div class="chart-format">
        <label>{t('Legend')}<select class="ok-input" aria-label={t('Legend')} bind:value={legendPosition} onchange={() => legendChanged = true}>
          <option value={null}>{t('None')}</option><option value="r">{t('Right')}</option><option value="l">{t('Left')}</option><option value="t">{t('Top')}</option><option value="b">{t('Bottom')}</option><option value="tr">{t('Top right')}</option>
        </select></label>
        <fieldset onchange={() => labelsChanged = true}><legend>{t('Data labels')}</legend>
          <label><input type="checkbox" bind:checked={showValue} />{t('Show values')}</label>
          <label><input type="checkbox" bind:checked={showCategory} />{t('Show categories')}</label>
          <label><input type="checkbox" bind:checked={showSeriesName} />{t('Show series names')}</label>
          <label><input type="checkbox" bind:checked={showPercent} />{t('Show percentages')}</label>
        </fieldset>
      </div>
      {#if hasAxes}
        <details class="axis-settings"><summary>{t('Chart axes')}</summary>
          <fieldset class="axis-visibility" onchange={() => axisVisibilityChanged = true}>
            <legend>{t('Axis visibility')}</legend>
            <label><input type="checkbox" bind:checked={showCategoryAxis} />{t('Show category axis')}</label>
            <label><input type="checkbox" bind:checked={showValueAxis} />{t('Show value axis')}</label>
            <label><input type="checkbox" bind:checked={showMajorGridlines} />{t('Show major gridlines')}</label>
            <label><input type="checkbox" bind:checked={showMinorGridlines} />{t('Show minor gridlines')}</label>
          </fieldset>
          <div class="axis-grid" oninput={() => axesChanged = true}>
            <label>{t('Category axis title')}<input class="ok-input" bind:value={categoryAxisTitle} /></label>
            <label>{t('Value axis title')}<input class="ok-input" bind:value={valueAxisTitle} /></label>
            <label>{t('Axis minimum')}<input class="ok-input" type="number" step="any" placeholder={t('Automatic')} bind:value={axisMin} /></label>
            <label>{t('Axis maximum')}<input class="ok-input" type="number" step="any" placeholder={t('Automatic')} bind:value={axisMax} /></label>
            <label>{t('Major tick interval')}<input class="ok-input" type="number" step="any" placeholder={t('Automatic')} bind:value={majorUnit} /></label>
            <label>{t('Minor tick interval')}<input class="ok-input" type="number" step="any" placeholder={t('Automatic')} bind:value={minorUnit} /></label>
            <label>{t('Axis number format')}<input class="ok-input" placeholder={t('Automatic')} bind:value={numberFormat} /></label>
          </div>
          <p>{t('Leave values blank for automatic axis scaling.')}</p>
        </details>
      {/if}
      <div class="data-grid">
        <table aria-label={t('Chart data')}>
          <thead><tr><th>{t('Category')}</th>{#each series as entry, s}<th>
            <input class="ok-input" aria-label={`${t('Series name')} ${s + 1}`} bind:value={entry.name} />
            <div class="series-tools"><input type="color" aria-label={`${t('Series color')} ${s + 1}`} value={/^#[0-9a-f]{6}$/i.test(entry.color ?? '') ? entry.color : defaultSeriesColors[s % defaultSeriesColors.length]} onchange={e => entry.color = e.currentTarget.value} />
            <button type="button" class="ok-btn" disabled={series.length <= 1} aria-label={`${t('Remove series')} ${s + 1}`} onclick={() => series.splice(s, 1)}>×</button></div>
            {#if kind === 'line'}
              <details class="series-format"><summary>{t('Line and markers')}</summary>
                <label>{t('Marker shape')}<select class="ok-input" aria-label={`${t('Marker shape')} ${s + 1}`} value={entry.base.markerSymbol ?? 'auto'} onchange={e => entry.base = { ...entry.base, markerSymbol: e.currentTarget.value as ChartSeries['markerSymbol'] }}>
                  {#each markerSymbols as marker}<option value={marker.value}>{t(marker.label)}</option>{/each}
                  {#if entry.base.markerSymbol === 'picture'}<option value="picture">{t('Picture')}</option>{/if}
                </select></label>
                <label>{t('Marker size (pt)')}<input class="ok-input" type="number" min="2" max="72" step="1" placeholder={t('Automatic')} aria-label={`${t('Marker size (pt)')} ${s + 1}`} value={entry.base.markerSizePt ?? ''} oninput={e => entry.base = { ...entry.base, markerSizePt: e.currentTarget.value === '' ? undefined : e.currentTarget.valueAsNumber }} /></label>
                <label><input type="checkbox" aria-label={`${t('Smooth line')} ${s + 1}`} checked={entry.base.smooth ?? false} onchange={e => entry.base = { ...entry.base, smooth: e.currentTarget.checked }} />{t('Smooth line')}</label>
              </details>
            {/if}
          </th>{/each}<th></th></tr></thead>
          <tbody>{#each categories as _, r}<tr>
            <th><input class="ok-input" aria-label={`${t('Category')} ${r + 1}`} bind:value={categories[r]} /></th>
            {#each series as entry, s}<td><input class="ok-input" type="number" step="any" aria-label={`${t('Value')} ${r + 1}, ${s + 1}`} bind:value={entry.values[r]} /></td>{/each}
            <td><button type="button" class="ok-btn" disabled={categories.length <= 1} aria-label={`${t('Remove category')} ${r + 1}`} onclick={() => removeCategory(r)}>×</button></td>
          </tr>{/each}</tbody>
        </table>
      </div>
      <div class="add"><button class="ok-btn" type="button" onclick={addCategory}>{t('Add category')}</button><button class="ok-btn" type="button" disabled={singleSeries} onclick={addSeries}>{t('Add series')}</button></div>
    {:else}<p role="alert">{t('Select a supported chart to edit its data.')}</p>{/if}
    {#if !validSeries}<p role="alert">{t('Pie and doughnut charts require one series. Remove extra series or choose another chart type.')}</p>{/if}
    {#if !validAxes}<p role="alert">{t('Axis minimum must be below maximum, tick intervals must be positive, and logarithmic bounds must be positive.')}</p>{/if}
    {#if !validMarkers}<p role="alert">{t('Marker size must be a whole number from 2 to 72 points.')}</p>{/if}
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button class="ok-btn" type="button" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button class="ok-btn primary" type="submit" disabled={!supported || !validSeries || !validAxes || !validMarkers}>{t(edit ? 'Apply changes' : 'Insert chart')}</button></footer>
  </form>
</dialog>

<style>
  dialog { width: min(760px, 94vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form { display: flex; flex-direction: column; gap: 16px; }
  header, footer, .add, .series-tools { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
  footer { justify-content: flex-end; }
  .add { justify-content: flex-start; }
  .settings { display: grid; grid-template-columns: 1fr 2fr; gap: 12px; }
  label { display: grid; gap: 6px; }
  .chart-format { display: grid; grid-template-columns: 140px 1fr; gap: 12px; align-items: start; }
  fieldset { display: flex; flex-wrap: wrap; gap: 8px 16px; border: 1px solid var(--ok-border); }
  fieldset label { display: flex; align-items: center; gap: 4px; }
  .axis-visibility { margin-top: 12px; }
  .axis-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; margin-top: 12px; }
  summary { cursor: pointer; }
  .axis-settings p { font-size: 12px; opacity: 0.8; }
  .data-grid { overflow: auto; max-height: 44vh; }
  table { border-collapse: collapse; width: 100%; }
  th, td { border: 1px solid var(--ok-border); padding: 5px; font-weight: normal; }
  td input, th input { width: 130px; }
  .series-format { text-align: left; font-weight: normal; margin-top: 8px; }
  .series-format label { display: block; margin-top: 8px; }
  .data-grid .series-format input[type="checkbox"] { width: auto; margin-right: 6px; }
  .series-tools { margin-top: 6px; }
  input[type='color'] { width: 32px; height: 24px; }
  [role='alert'] { color: #bf3131; }
</style>
