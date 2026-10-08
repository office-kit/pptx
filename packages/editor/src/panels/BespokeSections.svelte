<script lang="ts">
  import { restoreRememberedImageFill } from '../core/remembered-image-fill.ts';
  import { insertRememberedTextureFill, rememberShapeFill } from '../core/remembered-fill.ts';
  import { defaultTexture, texturePng } from '../core/textures.ts';
  import { getEditor } from '../core/context.ts';
  import {
    getShapeText,
    getShapeParagraphCount,
    getParagraphPropertiesEffective,
    getShapeKind,
    getShapeFill,
    getShapeFillEffective,
    setShapeFill,
    setShapeNoFill,
    setShapeSlideBackgroundFill,
    setShapeGradientFill,
    setShapePatternFill,
    getShapeFillColorResolved,
    getShapeStroke,
    getShapeStrokeEffective,
    getShapeStrokeWidth,
    getShapeStrokeColorResolved,
    getShapeId,
    getSlidePartName,
    getSlideShapes,
    setShapeText,
    type GradientFillOptions,
  } from '@office-kit/pptx';
  import GradientFillSection from './GradientFillSection.svelte';
  import PatternFillSection from './PatternFillSection.svelte';
  import PictureFillSection from './PictureFillSection.svelte';
  import TransparencyField from './TransparencyField.svelte';
  import PaneSection from './PaneSection.svelte';
  import EffectSections from './EffectSections.svelte';
  import { shapeEffects } from './effects-model.ts';
  import LineStyleFields from './LineStyleFields.svelte';
  import SizePositionSection from './SizePositionSection.svelte';
  import TextBoxSection from './TextBoxSection.svelte';
  import TextFormatBar from '../ui/TextFormatBar.svelte';
  import ColorPicker from '../ui/ColorPicker.svelte';
  import { textFormatsInRange } from '../core/text-format-selection.ts';
  import { selectedShapeId } from '../core/selection.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';

  let { tab }: { tab: 'paint' | 'effects' | 'size' | 'all' } = $props();
  const DEFAULT_GRADIENT: GradientFillOptions = {
    path: 'linear', angleDeg: 90, scaled: true,
    stops: [
      { offset: 0, color: 'accent1', brightness: 0.95 },
      { offset: 0.74, color: 'accent1', brightness: 0.55 },
      { offset: 0.83, color: 'accent1', brightness: 0.55 },
      { offset: 1, color: 'accent1', brightness: 0.7 },
    ],
  };
  const editor = getEditor();
  const doc = editor.doc;
  const shape = $derived.by(() => {
    doc.version;
    const sel = doc.selection;
    if (sel.kind !== 'shape') return null;
    const id = selectedShapeId(sel);
    return id == null ? null : doc.shapeById(sel.slideIndex, id);
  });

  const effectsApplicable = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    return shapes.length > 0 && shapes.every(target => ['shape', 'connector', 'picture'].includes(getShapeKind(target)));
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
  const paintColors = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    const common = (read: typeof getShapeFill | typeof getShapeStroke) => {
      const values = new Set(shapes.map(target => {
        const paint = read(target);
        return paint.kind === 'solid' ? paint.color : undefined;
      }));
      return values.size === 1 ? [...values][0] : undefined;
    };
    return { fill: common(getShapeFill), stroke: common(getShapeStroke) };
  });
  const fillKind = $derived.by(() => {
    doc.version;
    const kinds = new Set(editor.selectedShapes().map(target => getShapeFillEffective(doc.pres, target).kind));
    return kinds.size === 1 ? [...kinds][0] : 'mixed';
  });
  async function changeFill(kind: 'none' | 'solid' | 'gradient' | 'pattern' | 'background' | 'image') {
    const selection = doc.selection;
    if (editor.selectionLocked() || selection.kind !== 'shape' || fillKind === kind) return;
    const slideKey = getSlidePartName(doc.slideAt(selection.slideIndex)!);
    const shapes = editor.selectedShapes();
    // Like the reference desktop app, a shape with no picture to restore gets the default texture.
    const needsTexture = kind === 'image' && shapes.some(target => getShapeFillEffective(doc.pres, target).kind !== 'image' && !doc.rememberedFills.get(`${slideKey}:${getShapeId(target)}`)?.image);
    const presentation = doc.pres, version = doc.version;
    const texture = needsTexture ? await texturePng(defaultTexture()) : null;
    if (doc.pres !== presentation || doc.version !== version || doc.selection !== selection) return;
    doc.transact(t('Fill'), () => {
      for (const target of shapes) {
        const key = `${slideKey}:${getShapeId(target)}`;
        const remembered = doc.rememberedFills.get(key) ?? {};
        const current = getShapeFillEffective(doc.pres, target);
        if (current.kind === kind) continue;
        rememberShapeFill(doc.pres, target, remembered);
        doc.rememberedFills.set(key, remembered);
        if (kind === 'image' && remembered.image) restoreRememberedImageFill(target, remembered.image);
        else if (kind === 'image') insertRememberedTextureFill(doc.pres, target, texture!, remembered);
        else if (kind === 'background') setShapeSlideBackgroundFill(target);
        else if (kind === 'none') setShapeNoFill(target);
        else if (kind === 'solid') setShapeFill(target, remembered.solid ?? { color: 'accent1' });
        else if (kind === 'pattern') setShapePatternFill(target, remembered.pattern ?? { preset: 'pct5', foreground: 'accent1', background: 'bg1' });
        else setShapeGradientFill(target, remembered.gradient ?? DEFAULT_GRADIENT);
      }
    });
  }
  const lineKind = $derived.by(() => {
    doc.version;
    const kinds = new Set(editor.selectedShapes().map(target => getShapeStrokeEffective(doc.pres, target).kind));
    return kinds.size === 1 ? [...kinds][0] : 'mixed';
  });
  // The reference desktop app's English UI calls both the line type and the solid dash "Solid
  // line"; Japanese distinguishes them (線 (単色) here, 実線 for the dash), so
  // the shared English key cannot carry this one.
  const solidLineLabel = $derived(getLocale() === 'ja' ? '線 (単色)' : 'Solid line');
  function changeLine(kind: 'none' | 'solid' | 'gradient') {
    if (editor.selectionLocked() || lineKind === kind) return;
    if (kind === 'none') editor.invoke('setShapeNoStroke');
    else if (kind === 'solid') editor.invoke('setShapeStroke', { options: { color: 'accent1' } });
    // The reference desktop app starts a gradient line from the same accent 1 ramp as a
    // gradient fill.
    else editor.invoke('setShapeStroke', { options: { fill: { kind: 'gradient', ...DEFAULT_GRADIENT } } });
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
      <PaneSection id="fill" label={t('Fill')}>
          <fieldset class="fill-types" disabled={editor.selectionLocked()} aria-label={t('Fill type')}>
            {#each [['none', 'No fill'], ['solid', 'Solid fill'], ['gradient', 'Gradient fill'], ['image', 'Picture or texture fill'], ['pattern', 'Pattern fill'], ['background', 'Slide background fill']] as [kind, label]}
              <label><input type="radio" name="shape-fill-type" checked={fillKind === kind} disabled={(kind === 'background' || kind === 'image') && editor.selectedShapes().some(target => getShapeKind(target) !== 'shape')}
                onclick={() => { if (kind === 'image') changeFill('image'); }}
                onchange={() => { if (kind === 'none' || kind === 'solid' || kind === 'gradient' || kind === 'pattern' || kind === 'background') changeFill(kind); }} />{t(label)}</label>
            {/each}
          </fieldset>
          <PictureFillSection visible={fillKind === 'image'} />
          {#if fillKind === 'gradient'}
            <GradientFillSection />
          {:else if fillKind === 'pattern'}
            <PatternFillSection />
          {:else if fillKind !== 'none' && fillKind !== 'background' && fillKind !== 'image'}
          <hr class="rule" />
          <div class="paint-field">
            <span>{t('Color')}</span>
            <span class="colorwrap">
              <span class="paint-state" data-paint-state="fill">{paintLabel(paint.fill)}</span>
              <ColorPicker label={t('Fill')} value={paintColors.fill} resolvedColor={colorValue(paint.fill)} disabled={editor.selectionLocked()} choose={applyFill} />
            </span>
          </div>
          <TransparencyField paint="fill" />
          {/if}
      </PaneSection>

      <PaneSection id="line" label={t('Line')}>
          <!-- The reference desktop app's Line section opens with the line type, like Fill. -->
          <fieldset class="fill-types" disabled={editor.selectionLocked()} aria-label={t('Line type')}>
            <label><input type="radio" name="shape-line-type" checked={lineKind === 'none'} onchange={() => changeLine('none')} />{t('No line')}</label>
            <label><input type="radio" name="shape-line-type" checked={lineKind === 'solid'} onchange={() => changeLine('solid')} />{solidLineLabel}</label>
            <label><input type="radio" name="shape-line-type" checked={lineKind === 'gradient'} onchange={() => changeLine('gradient')} />{t('Gradient line')}</label>
          </fieldset>
          {#if lineKind === 'gradient'}
          <GradientFillSection target="line" />
          {/if}
          {#if lineKind !== 'none'}
          <hr class="rule" />
          {#if lineKind !== 'gradient'}
          <div class="paint-field">
            <span>{t('Color')}</span>
            <span class="colorwrap">
              <span class="paint-state" data-paint-state="stroke">{paintLabel(paint.stroke)}</span>
              <ColorPicker label={t('Outline')} value={paintColors.stroke} resolvedColor={colorValue(paint.stroke)} disabled={editor.selectionLocked()} choose={applyStroke} />
            </span>
          </div>
          <TransparencyField paint="line" />
          {/if}
          <label class="paint-field">
            <span>{t('Width')}</span>
            <span class="number"><input class="ok-input" aria-label={t('Outline width (points)')} type="number" min="0" max="1584" step="any"
              value={widthValue()} placeholder={paintLabel(paint.width)}
              onchange={(e) => applyWidth(e.currentTarget)} /><span>pt</span></span>
          </label>
          <LineStyleFields />
          {/if}
      </PaneSection>

    </div>
    <div hidden={tab !== 'effects' && !(tab === 'all' && effectsApplicable)} class="effect-controls">
      <EffectSections target={shapeEffects} />
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
  .fill-types { border: 0; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 2px; }
  .fill-types label { min-height: 18px; }
  .fill-types label { display: flex; align-items: center; gap: 6px; font-size: 12px; }
  .fill-types input { margin: 0; accent-color: var(--ok-accent); }

  .paint-controls, .effect-controls { display: flex; flex-direction: column; }
  .size-controls {
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
  /* Label-left / control-right rows (the reference desktop app, Mac: 30 pt pitch, 26 pt controls). */
  .paint-field {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    min-height: 26px;
  }
  .rule { width: 100%; margin: 2px 0; border: none; border-top: 1px solid var(--ok-border); }
  .colorwrap { display: flex; align-items: center; gap: 6px; }
  .colorwrap :global(.trigger) { width: 39px; height: 26px; }
  .paint-state { color: var(--ok-text-3); }
  .number {
    display: flex;
    align-items: center;
    gap: 3px;
    width: 82px;
  }
  .number input {
    box-sizing: border-box;
    width: 64px;
    height: 26px;
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
