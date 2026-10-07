<script lang="ts">
  import { asColor, getShapeKind, getShapeText, setShapeTextFormat, type Color, type ColorTransform, type SlideShapeData, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { textFormatsInRange } from '../core/text-format-selection.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import ColorPicker from '../ui/ColorPicker.svelte';
  import PaneSection from './PaneSection.svelte';
  import SliderField from './SliderField.svelte';
  import { EMU_PER_POINT } from './effects-model.ts';

  // Text Options ▸ Text Fill & Outline: the fill and outline of every run in
  // the selected shapes, as PowerPoint applies them with no text selected.
  const editor = getEditor();
  const doc = editor.doc;
  const shapes = $derived.by(() => { doc.version; return editor.selectedShapes().filter((shape) => getShapeKind(shape) === 'shape'); });
  const disabled = $derived(!shapes.length || editor.selectionLocked());
  const formats = $derived.by(() => {
    doc.version;
    return shapes.map((shape) => textFormatsInRange(shape, { start: 0, end: getShapeText(shape).length }, undefined, { pres: doc.pres })[0] ?? {});
  });
  function common<T>(read: (format: TextFormat) => T | undefined): T | undefined {
    const values = formats.map(read);
    return values.length && values.every((value) => value === values[0]) ? values[0] : undefined;
  }
  const fillKind = $derived(common((format) => format.textFill?.kind ?? 'solid'));
  const lineKind = $derived(common((format) => (format.outline ? 'solid' : 'none')));
  const opacityOf = (transforms: readonly ColorTransform[] | undefined) => {
    const alpha = transforms?.find((transform) => transform.kind === 'alpha');
    return alpha && 'value' in alpha ? alpha.value : 1;
  };
  function apply(label: string, change: (shape: SlideShapeData, format: TextFormat) => TextFormat) {
    if (disabled) return;
    const items = shapes.map((shape, index) => ({ shape, format: formats[index]! }));
    doc.transact(t(label), () => { for (const { shape, format } of items) setShapeTextFormat(shape, change(shape, format)); });
  }
  const withoutAlpha = (transforms: readonly ColorTransform[] | undefined) => (transforms ?? []).filter((transform) => transform.kind !== 'alpha');
  function solidColor(color: Color, transforms: readonly ColorTransform[], opacity: number): TextFormat {
    const kept = withoutAlpha(transforms);
    return { color, colorTransforms: opacity < 1 ? [...kept, { kind: 'alpha', value: opacity }] : kept };
  }
  function changeFill(kind: 'solid' | 'gradient' | 'pattern') {
    if (fillKind === kind) return;
    apply('Text Fill', (_shape, format) => {
      const base = (format.color && asColor(String(format.color))) || 'tx1';
      if (kind === 'solid') return { color: base };
      if (kind === 'pattern') return { color: null, textFill: { kind: 'pattern', preset: 'pct5', foreground: 'accent1', background: 'bg1' } };
      // PowerPoint's default gradient from the run's color: light to dark.
      return { color: null, textFill: { kind: 'gradient', path: 'linear', angleDeg: 90, scaled: true, stops: [
        { offset: 0, color: base, brightness: 0.95 }, { offset: 1, color: base, brightness: 0.55 },
      ] } };
    });
  }
  // English PowerPoint calls both the line type and the solid dash "Solid line";
  // Japanese has 線 (単色) for the type.
  const solidLineLabel = $derived(getLocale() === 'ja' ? '線 (単色)' : 'Solid line');
</script>

<div class="text-fill">
  <PaneSection id="textFill" label={t('Text Fill')}>
    <fieldset class="pane-radios" {disabled} aria-label={t('Text fill type')}>
      <label title={t('No text fill is not supported by the library yet.')}><input type="radio" name="text-fill-type" disabled />{t('No fill')}</label>
      <label><input type="radio" name="text-fill-type" checked={fillKind === 'solid'} onchange={() => changeFill('solid')} />{t('Solid fill')}</label>
      <label><input type="radio" name="text-fill-type" checked={fillKind === 'gradient'} onchange={() => changeFill('gradient')} />{t('Gradient fill')}</label>
      <label title={t('Picture fills for text are not supported by the library yet.')}><input type="radio" name="text-fill-type" disabled />{t('Picture or texture fill')}</label>
      <label><input type="radio" name="text-fill-type" checked={fillKind === 'pattern'} onchange={() => changeFill('pattern')} />{t('Pattern fill')}</label>
    </fieldset>
    {#if fillKind === 'solid'}
      <hr class="pane-rule" />
      <div class="pane-row"><span>{t('Color')}</span><span class="pane-color"><ColorPicker label={t('Text Fill Color')} value={common((format) => format.color ? String(format.color) : undefined)} {disabled}
        choose={(color, transforms) => apply('Text Fill', (_shape, format) => solidColor(color, transforms ?? [], opacityOf(format.colorTransforms)))} /></span></div>
      <SliderField label={t('Transparency')} name={t('Text fill transparency')} value={common((format) => Math.round((1 - opacityOf(format.colorTransforms)) * 100))} min={0} max={100} unit="%" {disabled}
        apply={(value) => apply('Text Fill', (_shape, format) => solidColor((format.color && asColor(String(format.color))) || 'tx1', format.colorTransforms ?? [], 1 - value / 100))} />
    {/if}
  </PaneSection>

  <PaneSection id="textOutline" label={t('Text Outline')}>
    <fieldset class="pane-radios" {disabled} aria-label={t('Text outline type')}>
      <label><input type="radio" name="text-line-type" checked={lineKind === 'none'} onchange={() => apply('Text Outline', () => ({ outline: null }))} />{t('No line')}</label>
      <label><input type="radio" name="text-line-type" checked={lineKind === 'solid'} onchange={() => apply('Text Outline', () => ({ outline: { color: 'tx1', widthEmu: 9525 } }))} />{solidLineLabel}</label>
      <label title={t('Gradient lines are not supported by the library yet.')}><input type="radio" name="text-line-type" disabled />{t('Gradient line')}</label>
    </fieldset>
    {#if lineKind === 'solid'}
      <hr class="pane-rule" />
      <div class="pane-row"><span>{t('Color')}</span><span class="pane-color"><ColorPicker label={t('Text Outline Color')} value={common((format) => format.outline?.color ? String(format.outline.color) : undefined)} {disabled}
        choose={(color, colorTransforms) => apply('Text Outline', (_shape, format) => { const { colorTransforms: _previous, ...outline } = format.outline ?? {}; return { outline: { ...outline, color, ...(colorTransforms?.length ? { colorTransforms } : {}) } }; })} /></span></div>
      <SliderField label={t('Width')} name={t('Text outline width')} slider={false} value={common((format) => format.outline?.widthEmu === undefined ? undefined : format.outline.widthEmu / EMU_PER_POINT)} min={0} max={1584} unit="pt" {disabled}
        apply={(points) => apply('Text Outline', (_shape, format) => ({ outline: { ...format.outline, widthEmu: Math.round(points * EMU_PER_POINT) } }))} />
    {/if}
  </PaneSection>
</div>

<style>
  .text-fill { display: flex; flex-direction: column; }
</style>
