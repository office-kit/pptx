<script lang="ts">
  import { asColor, getShapeKind, getShapeText, setShapeTextFormat, type Color, type ColorTransform, type LineFill, type SlideShapeData, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { textFormatsInRange } from '../core/text-format-selection.ts';
  import { getLocale, t } from '../i18n/i18n.svelte.ts';
  import ColorPicker from '../ui/ColorPicker.svelte';
  import PaneSection from './PaneSection.svelte';
  import SliderField from './SliderField.svelte';
  import { EMU_PER_POINT } from './effects-model.ts';
  import GradientFillSection from './GradientFillSection.svelte';
  import TexturePicker from '../ui/TexturePicker.svelte';
  import { defaultTexture, texturePng } from '../core/textures.ts';

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
  const lineKind = $derived(common((format) => (format.outline?.fill ? 'gradient' : format.outline ? 'solid' : 'none')));
  const opacityOf = (transforms: readonly ColorTransform[] | undefined) => {
    const alpha = transforms?.find((transform) => transform.kind === 'alpha');
    return alpha && 'value' in alpha ? alpha.value : 1;
  };
  let error = $state('');
  function apply(label: string, change: (shape: SlideShapeData, format: TextFormat) => TextFormat) {
    if (disabled) return;
    const items = shapes.map((shape, index) => ({ shape, format: formats[index]! }));
    try {
      doc.transact(t(label), () => { for (const { shape, format } of items) setShapeTextFormat(shape, change(shape, format)); });
      error = '';
    } catch (cause) { error = cause instanceof Error ? cause.message : String(cause); }
  }
  // Like PowerPoint, choosing Picture or texture fill starts from the default
  // texture; Insert, Clipboard and Texture then replace the picture.
  let input = $state<HTMLInputElement>();
  let loading = $state(false);
  async function fillWithPicture(read: () => Promise<Uint8Array>) {
    if (disabled || loading) return;
    const presentation = doc.pres, version = doc.version;
    loading = true;
    try {
      const bytes = await read();
      if (doc.pres !== presentation || doc.version !== version) return;
      apply('Text Fill', () => ({ textFill: { kind: 'image', bytes } }));
    } catch (cause) {
      error = cause instanceof DOMException && cause.name === 'NotAllowedError' ? t('Clipboard access was denied') : cause instanceof Error ? cause.message : String(cause);
    } finally { loading = false; }
  }
  async function upload(event: Event) {
    const element = event.currentTarget;
    if (!(element instanceof HTMLInputElement)) return;
    const file = element.files?.[0];
    if (!file) return;
    try { await fillWithPicture(async () => new Uint8Array(await file.arrayBuffer())); }
    finally { element.value = ''; }
  }
  function pastePicture() {
    return fillWithPicture(async () => {
      for (const item of await navigator.clipboard.read()) {
        const type = item.types.find((type) => type.startsWith('image/'));
        if (type) return new Uint8Array(await (await item.getType(type)).arrayBuffer());
      }
      throw new Error(t('The clipboard does not contain a picture.'));
    });
  }
  const withoutAlpha = (transforms: readonly ColorTransform[] | undefined) => (transforms ?? []).filter((transform) => transform.kind !== 'alpha');
  function solidColor(color: Color, transforms: readonly ColorTransform[], opacity: number): TextFormat {
    const kept = withoutAlpha(transforms);
    return { color, colorTransforms: opacity < 1 ? [...kept, { kind: 'alpha', value: opacity }] : kept };
  }
  function changeFill(kind: 'none' | 'solid' | 'gradient' | 'image' | 'pattern') {
    if (fillKind === kind) return;
    if (kind === 'image') { void fillWithPicture(() => texturePng(defaultTexture())); return; }
    apply('Text Fill', (_shape, format) => {
      if (kind === 'none') return { textFill: { kind: 'none' } };
      const base = (format.color && asColor(String(format.color))) || 'tx1';
      if (kind === 'solid') return { color: base };
      if (kind === 'pattern') return { textFill: { kind: 'pattern', preset: 'pct5', foreground: 'accent1', background: 'bg1' } };
      // PowerPoint's default gradient from the run's color: light to dark.
      return { textFill: { kind: 'gradient', path: 'linear', angleDeg: 90, scaled: true, stops: [
        { offset: 0, color: base, brightness: 0.95 }, { offset: 1, color: base, brightness: 0.55 },
      ] } };
    });
  }
  // English PowerPoint calls both the line type and the solid dash "Solid line";
  // Japanese has 線 (単色) for the type.
  const solidLineLabel = $derived(getLocale() === 'ja' ? '線 (単色)' : 'Solid line');
  // PowerPoint starts a gradient text outline from the accent 1 ramp it uses
  // for gradient fills.
  const DEFAULT_OUTLINE_GRADIENT = {
    kind: 'gradient', path: 'linear', angleDeg: 90, scaled: true,
    stops: [
      { offset: 0, color: 'accent1', brightness: 0.95 },
      { offset: 0.74, color: 'accent1', brightness: 0.55 },
      { offset: 0.83, color: 'accent1', brightness: 0.55 },
      { offset: 1, color: 'accent1', brightness: 0.7 },
    ],
  } as const satisfies LineFill;
</script>

<input type="file" accept="image/*" hidden bind:this={input} onchange={upload} aria-label={t('Picture source')} />

<div class="text-fill">
  <PaneSection id="textFill" label={t('Text Fill')}>
    <fieldset class="pane-radios" {disabled} aria-label={t('Text fill type')}>
      <label><input type="radio" name="text-fill-type" checked={fillKind === 'none'} onchange={() => changeFill('none')} />{t('No fill')}</label>
      <label><input type="radio" name="text-fill-type" checked={fillKind === 'solid'} onchange={() => changeFill('solid')} />{t('Solid fill')}</label>
      <label><input type="radio" name="text-fill-type" checked={fillKind === 'gradient'} onchange={() => changeFill('gradient')} />{t('Gradient fill')}</label>
      <label><input type="radio" name="text-fill-type" checked={fillKind === 'image'} onchange={() => changeFill('image')} />{t('Picture or texture fill')}</label>
      <label><input type="radio" name="text-fill-type" checked={fillKind === 'pattern'} onchange={() => changeFill('pattern')} />{t('Pattern fill')}</label>
    </fieldset>
    {#if fillKind === 'image'}
      <hr class="pane-rule" />
      <fieldset class="picture" disabled={disabled || loading}>
        <span>{t('Picture source')}</span>
        <button class="ok-btn" onclick={() => input?.click()}>{t('Insert...')}</button>
        <button class="ok-btn" disabled={!navigator.clipboard?.read} onclick={pastePicture}>{t('Clipboard')}</button>
        <TexturePicker disabled={disabled || loading} choose={(id) => fillWithPicture(() => texturePng(id))} />
      </fieldset>
    {/if}
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
      <label><input type="radio" name="text-line-type" checked={lineKind === 'gradient'} onchange={() => apply('Text Outline', (_shape, format) => ({ outline: { widthEmu: format.outline?.widthEmu ?? 9525, fill: DEFAULT_OUTLINE_GRADIENT } }))} />{t('Gradient line')}</label>
    </fieldset>
    {#if lineKind === 'gradient'}
      <hr class="pane-rule" />
      <GradientFillSection target="textOutline" />
      <SliderField label={t('Width')} name={t('Text outline width')} slider={false} value={common((format) => format.outline?.widthEmu === undefined ? undefined : format.outline.widthEmu / EMU_PER_POINT)} min={0} max={1584} unit="pt" {disabled}
        apply={(points) => apply('Text Outline', (_shape, format) => ({ outline: { ...format.outline, widthEmu: Math.round(points * EMU_PER_POINT) } }))} />
    {/if}
    {#if lineKind === 'solid'}
      <hr class="pane-rule" />
      <div class="pane-row"><span>{t('Color')}</span><span class="pane-color"><ColorPicker label={t('Text Outline Color')} value={common((format) => format.outline?.color ? String(format.outline.color) : undefined)} {disabled}
        choose={(color, colorTransforms) => apply('Text Outline', (_shape, format) => { const { colorTransforms: _previous, ...outline } = format.outline ?? {}; return { outline: { ...outline, color, ...(colorTransforms?.length ? { colorTransforms } : {}) } }; })} /></span></div>
      <SliderField label={t('Width')} name={t('Text outline width')} slider={false} value={common((format) => format.outline?.widthEmu === undefined ? undefined : format.outline.widthEmu / EMU_PER_POINT)} min={0} max={1584} unit="pt" {disabled}
        apply={(points) => apply('Text Outline', (_shape, format) => ({ outline: { ...format.outline, widthEmu: Math.round(points * EMU_PER_POINT) } }))} />
    {/if}
  </PaneSection>
</div>

{#if error}<p role="alert">{error}</p>{/if}

<style>
  .text-fill { display: flex; flex-direction: column; }
  .picture { border: 0; padding: 0; margin: 0; display: grid; gap: 8px; min-width: 0; font-size: 11px; }
  [role=alert] { color: #bf3131; font-size: 11px; }
</style>
