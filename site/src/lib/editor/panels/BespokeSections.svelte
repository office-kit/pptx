<script lang="ts">
  // Hand-tuned quick controls for the most-used shape properties, shown at the
  // top of the properties panel. These are ergonomic shortcuts; the full,
  // exhaustive list still lives below in the auto-generated section, so nothing
  // here is the *only* path to a capability.
  import { getEditor } from '../core/context.ts';
  import {
    getShapeBoundsResolved,
    getShapeRotation,
    getShapeFlip,
    getShapeText,
    inches,
    emu,
    setShapeBounds,
    getShapeFill,
    getShapeFillColorResolved,
    getShapeStroke,
    getShapeStrokeDash,
    getShapeStrokeWidth,
    type LineDash,
    getShapeStrokeColorResolved,
    getShapeId,
    getSlideShapes,
    setShapeText,
  } from '@office-kit/pptx';
  import { selectedShapeId } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;

  const shape = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    if (sel.kind !== 'shape') return null;
    const id = selectedShapeId(sel);
    return id == null ? null : doc.shapeById(sel.slideIndex, id);
  });

  const geometry = $derived.by(() => {
    doc.version;
    const items = [];
    for (const target of editor.selectedShapes()) {
      const bounds = getShapeBoundsResolved(doc.pres, target);
      if (!bounds) return [];
      items.push({ shape: target, bounds });
    }
    return items;
  });
  let lockAspectRatio = $state(false);
  const canLockAspectRatio = $derived(geometry.length > 0 && geometry.every(item => item.bounds.w > 0 && item.bounds.h > 0));
  const emuPerInch = inches(1);
  // DrawingML ST_Coordinate limits, expressed in the panel's inches.
  const minPosition = -27273042329600 / emuPerInch;
  const maxDimension = 27273042316900 / emuPerInch;
  function emuToIn(value: number): number {
    return Math.round((value / emuPerInch) * 100) / 100;
  }

  const bounds = $derived.by(() => {
    if (!geometry.length) return null;
    const common = (field: 'x' | 'y' | 'w' | 'h') => {
      const values = new Set(geometry.map(item => item.bounds[field]));
      return values.size === 1 ? emuToIn(geometry[0]!.bounds[field]) : null;
    };
    return { x: common('x'), y: common('y'), w: common('w'), h: common('h') };
  });
  const rotation = $derived.by(() => {
    doc.version;
    const values = new Set(editor.selectedShapes().map(getShapeRotation));
    return values.size > 1 ? null : [...values][0] ?? 0;
  });

  const flips = $derived.by(() => {
    doc.version;
    const values = editor.selectedShapes().map(target => getShapeFlip(target));
    const horizontal = new Set(values.map(value => value?.horizontal ?? false));
    const vertical = new Set(values.map(value => value?.vertical ?? false));
    return {
      horizontal: horizontal.size > 1 ? null : horizontal.has(true),
      vertical: vertical.size > 1 ? null : vertical.has(true),
    };
  });

  const text = $derived.by(() => {
    doc.version;
    const s = shape;
    if (!s) return '';
    try {
      return getShapeText(s);
    } catch {
      return '';
    }
  });

  const paint = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    if (sel.kind !== 'shape') return { fill: 'inherit', stroke: 'inherit', width: 'inherit', dash: 'inherit' };
    const slide = doc.slideAt(sel.slideIndex);
    if (!slide) return { fill: 'inherit', stroke: 'inherit', width: 'inherit', dash: 'inherit' };
    const ids = new Set(sel.shapeIds);
    const fills = new Set<string>();
    const strokes = new Set<string>();
    const widths = new Set<string>();
    const dashes = new Set<string>();
    for (const target of getSlideShapes(slide)) {
      if (!ids.has(getShapeId(target))) continue;
      const fill = getShapeFill(target);
      const stroke = getShapeStroke(target);
      const width = getShapeStrokeWidth(target);
      widths.add(width == null ? 'inherit' : String(width / 12700));
      dashes.add(getShapeStrokeDash(target) ?? 'inherit');
      fills.add(fill.kind === 'solid' ? getShapeFillColorResolved(doc.pres, target) ?? fill.color : fill.kind);
      strokes.add(stroke.kind === 'solid' ? getShapeStrokeColorResolved(doc.pres, target) ?? stroke.color : stroke.kind);
    }
    return {
      fill: fills.size > 1 ? 'mixed' : [...fills][0] ?? 'inherit',
      stroke: strokes.size > 1 ? 'mixed' : [...strokes][0] ?? 'inherit',
      width: widths.size > 1 ? 'mixed' : [...widths][0] ?? 'inherit',
      dash: dashes.size > 1 ? 'mixed' : [...dashes][0] ?? 'inherit',
    };
  });

  function paintLabel(value: string): string {
    switch (value) {
      case 'mixed': return t('Mixed');
      case 'none': return t('None');
      case 'inherit': return t('Inherited');
      case 'gradient': return t('Gradient');
      case 'pattern': return t('Pattern');
      case 'image': return t('Picture');
      default: return value;
    }
  }
  function colorValue(value: string): string {
    return /^#[0-9a-f]{6}$/i.test(value) ? value : '#000000';
  }
  function applyFill(value: string) {
    editor.invoke('setShapeFill', { color: value.replace('#', '') });
  }
  function applyStroke(value: string) {
    editor.invoke('setShapeStroke', { options: { color: value.replace('#', '') } });
  }
  const dashStyles: Array<[LineDash, string]> = [
    ['solid', 'Solid line'], ['dot', 'Dotted line'], ['dash', 'Dashed line'],
    ['lgDash', 'Long dashed line'], ['dashDot', 'Dash-dot line'],
    ['lgDashDot', 'Long dash-dot line'], ['lgDashDotDot', 'Long dash-dot-dot line'],
    ['sysDash', 'System dashed line'], ['sysDot', 'System dotted line'],
    ['sysDashDot', 'System dash-dot line'], ['sysDashDotDot', 'System dash-dot-dot line'],
  ];
  function widthValue(): string {
    return paint.width === 'mixed' || paint.width === 'inherit' ? '' : paint.width;
  }
  function applyWidth(input: HTMLInputElement) {
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) {
      input.value = widthValue();
      return;
    }
    editor.invoke('setShapeStroke', { options: { widthEmu: Math.round(input.valueAsNumber * 12700) } });
  }
  function applyDash(value: string) {
    const dash = dashStyles.find(([key]) => key === value)?.[0];
    if (dash) editor.invoke('setShapeStrokeDash', { dash });
  }
  function setBoundsField(field: 'x' | 'y' | 'w' | 'h', input: HTMLInputElement) {
    if (!bounds || !geometry.length) return;
    const restore = () => { input.value = bounds?.[field] == null ? '' : String(bounds[field]); };
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) {
      restore();
      return;
    }
    const value = inches(input.valueAsNumber);
    const updates = geometry.map(({ shape, bounds: current }) => {
      const next = { ...current, [field]: value };
      if (lockAspectRatio && canLockAspectRatio) {
        if (field === 'w') next.h = emu(current.h * next.w / current.w);
        if (field === 'h') next.w = emu(current.w * next.h / current.h);
      }
      return { shape, bounds: next };
    });
    if (updates.some(item => item.bounds.w > maxDimension * emuPerInch || item.bounds.h > maxDimension * emuPerInch)) {
      restore();
      editor.toast('error', t('The proportional size is too large'));
      return;
    }
    doc.transact(t('Set bounds'), () => {
      for (const item of updates) setShapeBounds(item.shape, item.bounds);
    });
  }
  function applyRotation(input: HTMLInputElement) {
    const s = shape;
    if (!s) return;
    if (!input.reportValidity() || !Number.isFinite(input.valueAsNumber)) {
      input.value = rotation === null ? '' : String(rotation);
      return;
    }
    editor.invoke('setShapeRotation', { degrees: input.valueAsNumber });
  }
  function applyText(value: string) {
    const s = shape;
    if (!s) return;
    doc.transact(t('Edit text'), () => setShapeText(s, value, { preserveFormatting: true }));
  }
</script>

{#if shape}
  <div class="bespoke">
    <div class="sec">
      <div class="sec-title">{t('Fill & outline')}</div>
      <div class="row2">
        <div class="mini">
          <span>{t('Fill')}</span>
          <span class="colorwrap">
            <input type="color" aria-label={t('Fill')} value={colorValue(paint.fill)} onchange={(e) => applyFill(e.currentTarget.value)} />
            <span data-paint-state="fill">{paintLabel(paint.fill)}</span>
          </span>
          <button class="ok-btn" onclick={() => editor.invoke('setShapeNoFill')}>{t('No fill')}</button>
        </div>
        <div class="mini">
          <span>{t('Outline')}</span>
          <span class="colorwrap">
            <input type="color" aria-label={t('Outline')} value={colorValue(paint.stroke)} onchange={(e) => applyStroke(e.currentTarget.value)} />
            <span data-paint-state="stroke">{paintLabel(paint.stroke)}</span>
          </span>
          <button class="ok-btn" onclick={() => editor.invoke('setShapeNoStroke')}>{t('No outline')}</button>
        </div>
      </div>
    </div>

    <div class="row2">
      <label class="mini">
        <span>{t('Outline width (points)')}</span>
        <input class="ok-input" type="number" min="0" max="1584" step="any"
          value={widthValue()} placeholder={paintLabel(paint.width)}
          onchange={(e) => applyWidth(e.currentTarget)} />
      </label>
      <label class="mini">
        <span>{t('Outline style')}</span>
        <select class="ok-input" value={paint.dash} onchange={(e) => applyDash(e.currentTarget.value)}>
          {#if !dashStyles.some(([key]) => key === paint.dash)}
            <option value={paint.dash} disabled>{paintLabel(paint.dash)}</option>
          {/if}
          {#each dashStyles as [value, label]}<option {value}>{t(label)}</option>{/each}
        </select>
      </label>
    </div>

    {#if bounds}
      <div class="sec">
        <div class="sec-title">{t('Position & size (in)')}</div>
        {#if geometry.length > 1}<p class="scope">{t('Values apply to each selected object')}</p>{/if}
        <label class="aspect-lock">
          <input type="checkbox" bind:checked={lockAspectRatio} disabled={!canLockAspectRatio} />
          <span>{t('Lock aspect ratio')}</span>
        </label>
        <div class="grid4">
          <label class="mini"><span>X</span>
            <input class="ok-input" type="number" step="any" min={minPosition} max={maxDimension} value={bounds.x ?? ''} placeholder={bounds.x === null ? t('Mixed') : undefined}
              onchange={(e) => setBoundsField('x', e.currentTarget)} /></label>
          <label class="mini"><span>Y</span>
            <input class="ok-input" type="number" step="any" min={minPosition} max={maxDimension} value={bounds.y ?? ''} placeholder={bounds.y === null ? t('Mixed') : undefined}
              onchange={(e) => setBoundsField('y', e.currentTarget)} /></label>
          <label class="mini"><span>W</span>
            <input class="ok-input" type="number" step="any" min={0} max={maxDimension} value={bounds.w ?? ''} placeholder={bounds.w === null ? t('Mixed') : undefined}
              onchange={(e) => setBoundsField('w', e.currentTarget)} /></label>
          <label class="mini"><span>H</span>
            <input class="ok-input" type="number" step="any" min={0} max={maxDimension} value={bounds.h ?? ''} placeholder={bounds.h === null ? t('Mixed') : undefined}
              onchange={(e) => setBoundsField('h', e.currentTarget)} /></label>
        </div>
      </div>
    {/if}

    <div class="sec">
      <div class="sec-title">{t('Rotation')}</div>
      <div class="rotrow">
        <input class="ok-input" type="number" aria-label={t('Rotation')} step="any" value={rotation ?? ''} placeholder={rotation === null ? t('Mixed') : undefined}
          onchange={(e) => applyRotation(e.currentTarget)} />
        <span class="deg">°</span>
      </div>
    </div>

    <div class="row2">
      {#each ['horizontal', 'vertical'] as axis}
        {@const value = axis === 'horizontal' ? flips.horizontal : flips.vertical}
        <label class="aspect-lock">
          <input type="checkbox" checked={value ?? false} indeterminate={value === null}
            onchange={event => editor.invoke('setShapeFlip', { options: { [axis]: event.currentTarget.checked } })} />
          <span>{t(axis === 'horizontal' ? 'Flip horizontally' : 'Flip vertically')}{value === null ? ` (${t('Mixed')})` : ''}</span>
        </label>
      {/each}
    </div>

    <div class="sec">
      <div class="sec-title">{t('Text')}</div>
      <textarea class="ok-input" aria-label={t('Text')} rows="2" value={text}
        onchange={(e) => applyText(e.currentTarget.value)}></textarea>
    </div>
  </div>
{/if}

<style>
  .bespoke {
    padding: 8px 10px;
    border-bottom: 1px solid var(--ok-border);
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .sec-title {
    font-size: 11px;
    font-weight: 600;
    color: var(--ok-text-2);
    margin-bottom: 5px;
  }
  .row2 {
    display: flex;
    gap: 10px;
  }
  .mini {
    display: flex;
    flex-direction: column;
    gap: 3px;
    font-size: 11px;
    color: var(--ok-text-2);
    flex: 1;
    min-width: 0;
  }
  .colorwrap input[type='color'] {
    width: 100%;
    height: 26px;
    border: 1px solid var(--ok-border-strong);
    border-radius: var(--ok-radius);
    background: none;
    padding: 0;
    cursor: pointer;
  }
  .scope { font-size: 11px; color: var(--ok-text-2); margin: 0 0 6px; }
  .aspect-lock {
    display: flex;
    align-items: center;
    gap: 5px;
    margin-bottom: 6px;
    font-size: 11px;
  }
  .grid4 {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px;
  }
  .rotrow {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .rotrow .ok-input {
    width: 80px;
  }
  .deg {
    color: var(--ok-text-2);
  }
  textarea.ok-input {
    resize: vertical;
    width: 100%;
    font-family: var(--ok-font);
  }
</style>
