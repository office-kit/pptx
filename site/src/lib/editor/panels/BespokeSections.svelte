<script lang="ts">
  // Hand-tuned quick controls for the most-used shape properties, shown at the
  // top of the properties panel. These are ergonomic shortcuts; the full,
  // exhaustive list still lives below in the auto-generated section, so nothing
  // here is the *only* path to a capability.
  import { getEditor } from '../core/context.ts';
  import {
    getShapeText,
    getShapeTextAnchor,
    getShapeParagraphCount,
    getParagraphPropertiesEffective,
    getShapeKind,
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
  import SizePositionSection from './SizePositionSection.svelte';
  import TextFormatBar from '../ui/TextFormatBar.svelte';
  import { textFormatsInRange } from '../core/text-format-selection.ts';
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

  const textShape = $derived.by(() => {
    const selection = doc.selection;
    if (selection.kind !== 'shape' || selection.shapeIds.length !== 1) return null;
    return shape && getShapeKind(shape) === 'shape' ? shape : null;
  });
  const text = $derived.by(() => {
    doc.version;
    return textShape ? getShapeText(textShape) : '';
  });

  const objectFormats = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    if (!shapes.length || !shapes.every(target => getShapeKind(target) === 'shape')) return null;
    return shapes.flatMap(target => {
      const formats = textFormatsInRange(target, { start: 0, end: getShapeText(target).length }, undefined, { pres: doc.pres });
      return formats.length ? formats : [{}];
    });
  });

  const textAlignment = $derived.by(() => {
    doc.version;
    if (!objectFormats) return { align: '', anchor: '' };
    const horizontal = new Set<string>();
    const vertical = new Set<string>();
    for (const target of editor.selectedShapes()) {
      vertical.add(getShapeTextAnchor(target) ?? '');
      const count = getShapeParagraphCount(target);
      if (!count) horizontal.add('');
      for (let index = 0; index < count; index++) {
        const align = getParagraphPropertiesEffective(doc.pres, target, index).align ?? '';
        horizontal.add(['left', 'center', 'right', 'justify'].includes(align) ? align : '');
      }
    }
    return { align: horizontal.size === 1 ? [...horizontal][0] : '', anchor: vertical.size === 1 ? [...vertical][0] : '' };
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
  function applyText(value: string) {
    const s = textShape;
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

    <SizePositionSection />

    {#if objectFormats}
      <TextFormatBar formats={objectFormats} selected context="objects"
        onformat={(format, reset) => editor.invoke('setShapeTextFormat', { format, options: { reset } })} />
      <div class="row2">
        <label>{t('Paragraph alignment')}
          <select class="ok-input" aria-label={t('Paragraph alignment')} value={textAlignment.align}
            onchange={event => editor.invoke('setShapeAlignment', { align: event.currentTarget.value })}>
            <option value="" disabled>{t('Mixed or inherited')}</option>
            {#each [['left', 'Left'], ['center', 'Center'], ['right', 'Right'], ['justify', 'Justify']] as [value, label]}
              <option {value}>{t(label)}</option>
            {/each}
          </select>
        </label>
        <label>{t('Vertical alignment')}
          <select class="ok-input" aria-label={t('Vertical alignment')} value={textAlignment.anchor}
            onchange={event => editor.invoke('setShapeTextAnchor', { anchor: event.currentTarget.value })}>
            <option value="" disabled>{t('Mixed or inherited')}</option>
            {#each [['top', 'Top'], ['center', 'Center'], ['bottom', 'Bottom']] as [value, label]}
              <option {value}>{t(label)}</option>
            {/each}
          </select>
        </label>
      </div>
    {/if}

    {#if textShape}
      <div class="sec">
        <div class="sec-title">{t('Text')}</div>
        <textarea class="ok-input" aria-label={t('Text')} rows="2" value={text}
          onchange={(e) => applyText(e.currentTarget.value)}></textarea>
      </div>
    {:else if objectFormats}
      <p class="scope">{t('Select one text shape to edit its content.')}</p>
    {/if}
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
  textarea.ok-input {
    resize: vertical;
    width: 100%;
    font-family: var(--ok-font);
  }
</style>
