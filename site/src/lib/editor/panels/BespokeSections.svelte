<script lang="ts">
  import { getEditor } from '../core/context.ts';
  import {
    getShapeText,
    getShapeParagraphCount,
    getParagraphPropertiesEffective,
    getShapeKind,
    getShapeFill,
    getShapeFillEffective,
    getShapeFillOpacity,
    getShapeGradientFillEffective,
    setShapeFill,
    setShapeNoFill,
    setShapeSlideBackgroundFill,
    setShapeGradientFill,
    getShapePatternFill,
    setShapePatternFill,
    asColor,
    getShapeFillColorResolved,
    getShapeStroke,
    getShapeStrokeWidth,
    getShapeStrokeColorResolved,
    getShapeId,
    getSlidePartName,
    getSlideShapes,
    setShapeText,
  } from '@office-kit/pptx';
  import GradientFillSection from './GradientFillSection.svelte';
  import PatternFillSection from './PatternFillSection.svelte';
  import TransparencyField from './TransparencyField.svelte';
  import LineStyleFields from './LineStyleFields.svelte';
  import SizePositionSection from './SizePositionSection.svelte';
  import TextBoxSection from './TextBoxSection.svelte';
  import TextFormatBar from '../ui/TextFormatBar.svelte';
  import { textFormatsInRange } from '../core/text-format-selection.ts';
  import { selectedShapeId } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { tab }: { tab: 'paint' | 'effects' | 'size' | 'all' } = $props();
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
    if (!objectFormats) return { align: '' };
    const horizontal = new Set<string>();
    for (const target of editor.selectedShapes()) {
      const count = getShapeParagraphCount(target);
      if (!count) horizontal.add('');
      for (let index = 0; index < count; index++) {
        const align = getParagraphPropertiesEffective(doc.pres, target, index).align ?? '';
        horizontal.add(['left', 'center', 'right', 'justify'].includes(align) ? align : '');
      }
    }
    return { align: horizontal.size === 1 ? [...horizontal][0] : '' };
  });

  const paint = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    if (sel.kind !== 'shape') return { fill: 'inherit', stroke: 'inherit', width: 'inherit' };
    const slide = doc.slideAt(sel.slideIndex);
    if (!slide) return { fill: 'inherit', stroke: 'inherit', width: 'inherit' };
    const ids = new Set(sel.shapeIds);
    const fills = new Set<string>();
    const strokes = new Set<string>();
    const widths = new Set<string>();
    for (const target of getSlideShapes(slide)) {
      if (!ids.has(getShapeId(target))) continue;
      const fill = getShapeFill(target);
      const stroke = getShapeStroke(target);
      const width = getShapeStrokeWidth(target);
      widths.add(width == null ? 'inherit' : String(width / 12700));
      fills.add(fill.kind === 'solid' ? getShapeFillColorResolved(doc.pres, target) ?? fill.color : fill.kind);
      strokes.add(stroke.kind === 'solid' ? getShapeStrokeColorResolved(doc.pres, target) ?? stroke.color : stroke.kind);
    }
    return {
      fill: fills.size > 1 ? 'mixed' : [...fills][0] ?? 'inherit',
      stroke: strokes.size > 1 ? 'mixed' : [...strokes][0] ?? 'inherit',
      width: widths.size > 1 ? 'mixed' : [...widths][0] ?? 'inherit',
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
  const fillKind = $derived.by(() => {
    doc.version;
    const kinds = new Set(editor.selectedShapes().map(target => getShapeFillEffective(doc.pres, target).kind));
    return kinds.size === 1 ? [...kinds][0] : 'mixed';
  });
  function changeFill(kind: 'none' | 'solid' | 'gradient' | 'pattern' | 'background') {
    if (editor.selectionLocked() || doc.selection.kind !== 'shape' || fillKind === kind) return;
    const slideKey = getSlidePartName(doc.slideAt(doc.selection.slideIndex)!);
    const shapes = editor.selectedShapes();
    doc.transact(t('Fill'), () => {
      for (const target of shapes) {
        const key = `${slideKey}:${getShapeId(target)}`;
        const remembered = doc.rememberedFills.get(key) ?? {};
        const current = getShapeFillEffective(doc.pres, target);
        if (current.kind === kind) continue;
        if (current.kind === 'pattern') remembered.pattern = getShapePatternFill(doc.pres, target) ?? undefined;
        if (current.kind === 'solid') remembered.solid = { color: asColor(getShapeFillColorResolved(doc.pres, target) ?? current.color) ?? 'accent1', opacity: getShapeFillOpacity(target) ?? undefined };
        if (current.kind === 'gradient') {
          const gradient = getShapeGradientFillEffective(doc.pres, target);
          if (gradient) remembered.gradient = { ...gradient, stops: gradient.stops.map(stop => {
            const color = asColor(stop.color);
            return { ...stop, color: color ?? asColor(stop.resolvedColor ?? '') ?? 'accent1', brightness: color ? stop.brightness : 0 };
          }) };
        }
        doc.rememberedFills.set(key, remembered);
        if (kind === 'background') setShapeSlideBackgroundFill(target);
        else if (kind === 'none') setShapeNoFill(target);
        else if (kind === 'solid') setShapeFill(target, remembered.solid ?? { color: 'accent1' });
        else if (kind === 'pattern') setShapePatternFill(target, remembered.pattern ?? { preset: 'pct5', foreground: 'accent1', background: 'bg1' });
        else setShapeGradientFill(target, remembered.gradient ?? {
          path: 'linear', angleDeg: 90, scaled: true,
          stops: [
            { offset: 0, color: 'accent1', brightness: 0.95 },
            { offset: 0.74, color: 'accent1', brightness: 0.55 },
            { offset: 0.83, color: 'accent1', brightness: 0.55 },
            { offset: 1, color: 'accent1', brightness: 0.7 },
          ],
        });
      }
    });
  }
  function applyFill(value: string) {
    editor.invoke('setShapeFill', { color: { color: value.replace('#', '') } });
  }
  function applyStroke(value: string) {
    editor.invoke('setShapeStroke', { options: { color: value.replace('#', '') } });
  }
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
  function applyText(value: string) {
    const s = textShape;
    if (!s) return;
    doc.transact(t('Edit text'), () => setShapeText(s, value, { preserveFormatting: true }));
  }
</script>

{#if shape}
  <div class="bespoke">
    <div hidden={tab !== 'paint' && tab !== 'all'} class="paint-controls">
      <details class="paint-section" open>
        <summary>{t('Fill')}</summary>
        <div class="paint-fields">
          <fieldset class="fill-types" disabled={editor.selectionLocked()} aria-label={t('Fill type')}>
            {#each [['none', 'No fill'], ['solid', 'Solid fill'], ['gradient', 'Gradient fill'], ['pattern', 'Pattern fill'], ['background', 'Slide background fill']] as [kind, label]}
              <label><input type="radio" name="shape-fill-type" checked={fillKind === kind} disabled={kind === 'background' && editor.selectedShapes().some(target => getShapeKind(target) !== 'shape')}
                onchange={() => { if (kind === 'none' || kind === 'solid' || kind === 'gradient' || kind === 'pattern' || kind === 'background') changeFill(kind); }} />{t(label)}</label>
            {/each}
          </fieldset>
          {#if fillKind === 'gradient'}
            <GradientFillSection />
          {:else if fillKind === 'pattern'}
            <PatternFillSection />
          {:else if fillKind !== 'none' && fillKind !== 'background'}
          <div class="mini">
            <span>{t('Color')}</span>
            <span class="colorwrap">
              <input type="color" aria-label={t('Fill')} value={colorValue(paint.fill)} onchange={(e) => applyFill(e.currentTarget.value)} />
              <span data-paint-state="fill">{paintLabel(paint.fill)}</span>
            </span>
          </div>
          <TransparencyField paint="fill" />
          {/if}
        </div>
      </details>

      <details class="paint-section" open>
        <summary>{t('Line')}</summary>
        <div class="paint-fields">
          <div class="mini">
            <span>{t('Color')}</span>
            <span class="colorwrap">
              <input type="color" aria-label={t('Outline')} value={colorValue(paint.stroke)} onchange={(e) => applyStroke(e.currentTarget.value)} />
              <span data-paint-state="stroke">{paintLabel(paint.stroke)}</span>
            </span>
            <button class="ok-btn" onclick={() => editor.invoke('setShapeNoStroke')}>{t('No outline')}</button>
          </div>
          <TransparencyField paint="line" />
          <label class="paint-field">
            <span>{t('Width')}</span>
            <span class="number"><input class="ok-input" aria-label={t('Outline width (points)')} type="number" min="0" max="1584" step="any"
              value={widthValue()} placeholder={paintLabel(paint.width)}
              onchange={(e) => applyWidth(e.currentTarget)} /><span>pt</span></span>
          </label>
          <LineStyleFields />
        </div>
      </details>

    </div>
    <div hidden={tab !== 'size' && tab !== 'all'} class="size-controls">
      <SizePositionSection />
      <TextBoxSection />

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
  </div>
{/if}

<style>
  .fill-types { border: 0; margin: 0 0 8px; padding: 0; display: flex; flex-direction: column; gap: 5px; }
  .fill-types label { display: flex; align-items: center; gap: 6px; font-size: 12px; }
  .fill-types input { margin: 0; accent-color: var(--ok-accent); }

  .paint-controls, .size-controls {
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  [hidden] {
    display: none;
  }
  .bespoke {
    padding: 8px 10px;
    border-bottom: 1px solid var(--ok-border);
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .paint-section {
    margin: 0 -10px;
    font-size: 11px;
  }
  .paint-section summary {
    padding: 4px 8px;
    background: var(--ok-hover);
    cursor: pointer;
  }
  .paint-fields {
    display: flex;
    flex-direction: column;
    gap: 8px;
    padding: 12px;
  }
  .paint-field {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
  }
  .number {
    display: flex;
    align-items: center;
    gap: 3px;
    width: 96px;
  }
  .number input {
    width: 72px;
    min-width: 0;
    padding: 2px 4px;
    font-size: inherit;
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
  .scope {
    font-size: 11px;
    color: var(--ok-text-2);
    margin: 0 0 6px;
  }
  textarea.ok-input {
    resize: vertical;
    width: 100%;
    font-family: var(--ok-font);
  }
</style>
