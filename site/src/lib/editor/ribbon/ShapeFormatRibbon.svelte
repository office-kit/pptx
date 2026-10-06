<script lang="ts">
  // Mac PowerPoint's Shape Format tab, in its order: Insert Shapes, Shape
  // Styles, WordArt Styles, Alt Text, Arrange, Size and Format Pane.
  import { placeBelowTrigger } from './place-menu.ts';
  import { cm, emu, getShapeBoundsResolved, getShapeKind, setShapeBounds, setShapePreset, setShapeText3D, setShapeTextFormat, type Color, type PresetShape, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { applyWordArtPreset, type WordArtPreset } from '../core/wordart-presets.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import ColorPicker from '../ui/ColorPicker.svelte';
  import ShapeFillPicker from './ShapeFillPicker.svelte';
  import Icon from '../ui/Icon.svelte';
  import WordArtGallery from '../ui/WordArtGallery.svelte';
  import ArrangeMenu from './ArrangeMenu.svelte';
  import { PRESET } from './config.ts';
  import ShapeQuickStyles from './ShapeQuickStyles.svelte';

  const editor = getEditor();
  const doc = editor.doc;
  const shapes = $derived.by(() => { doc.version; return editor.selectedShapes(); });
  const editable = $derived(shapes.length > 0 && !editor.selectionLocked());
  const paintable = $derived(editable && shapes.every((shape) => ['shape', 'connector'].includes(getShapeKind(shape))));
  const texty = $derived(editable && shapes.some((shape) => getShapeKind(shape) === 'shape'));
  let open = $state<'editShape' | 'shapeEffects' | 'wordArt' | 'textEffects' | null>(null);
  let wordArtButton = $state<HTMLButtonElement>();

  const CHANGE_SHAPES: readonly [PresetShape, string][] = [
    ['rect', 'Rectangle'],
    ['roundRect', 'Rounded Rectangle'],
    ['ellipse', 'Oval'],
    ['triangle', 'Isosceles Triangle'],
    ['diamond', 'Diamond'],
    ['rightArrow', 'Right Arrow'],
    ['hexagon', 'Hexagon'],
    ['star5', '5-Point Star'],
    ['leftArrow', 'Left Arrow'],
  ];
  // PowerPoint's outer shadow and glow presets for text ("Offset: Bottom Right",
  // "Glow: 5 point; Accent color 1").
  const TEXT_SHADOW = { color: '#000000', blurEmu: 38100, offsetEmu: 38100, angleDeg: 45, opacity: 0.4 } as const;
  const TEXT_GLOW = { color: 'accent1', radiusEmu: 63500, opacity: 0.4 } as const;
  const TEXT_REFLECTION = { blurEmu: 6350, offsetEmu: 0, angleDeg: 90, startOpacity: 0.5, opacity: 0.003, endPosition: 0.55 } as const;
  const THIN_OUTLINE_EMU = 9525;

  function textFormat(format: TextFormat) {
    open = null;
    editor.invoke('setShapeTextFormat', { format, options: { reset: false } });
  }
  function wordArt(preset: WordArtPreset) {
    open = null;
    doc.transact(t('WordArt Styles'), () => { for (const shape of shapes) if (getShapeKind(shape) === 'shape') applyWordArtPreset(shape, preset); });
  }
  // Clear WordArt removes the run effects and outline and the bevel's 3-D, and keeps the fill.
  function clearWordArt() {
    open = null;
    doc.transact(t('Clear WordArt'), () => {
      for (const shape of shapes) {
        if (getShapeKind(shape) !== 'shape') continue;
        setShapeTextFormat(shape, { outline: null, shadow: null, innerShadow: null, glow: null, reflection: null });
        setShapeText3D(shape, null);
      }
    });
  }
  function changeShape(preset: PresetShape) {
    open = null;
    doc.transact(t('Change Shape'), () => { for (const shape of shapes) if (getShapeKind(shape) === 'shape') setShapePreset(shape, preset); });
  }
  function outline(color: Color) { editor.invoke('setShapeStroke', { options: { color } }); }

  const geometry = $derived(shapes.flatMap((shape) => {
    const bounds = getShapeBoundsResolved(doc.pres, shape);
    return bounds ? [{ shape, bounds }] : [];
  }));
  const dimension = (field: 'w' | 'h') => {
    const values = geometry.map((item) => item.bounds[field]);
    return values.length > 0 && values.every((value) => value === values[0]) ? Math.round((values[0]! / cm(1)) * 100) / 100 : null;
  };
  function resize(field: 'w' | 'h', input: HTMLInputElement) {
    if (!editable || !input.reportValidity() || !Number.isFinite(input.valueAsNumber) || input.valueAsNumber <= 0) { input.value = String(dimension(field) ?? ''); return; }
    const value = emu(cm(input.valueAsNumber));
    doc.transact(t(field === 'w' ? 'Width' : 'Height'), () => { for (const { shape, bounds } of geometry) setShapeBounds(shape, { ...bounds, [field]: value }); });
  }
</script>

<svelte:window onpointerdown={(event) => { if (open && !(event.target as Element).closest?.('.shape-format .anchor')) open = null; }} onkeydown={(event) => { if (open && event.key === 'Escape') open = null; }} />

{#snippet menuButton(name: NonNullable<typeof open>, icon: string, label: string, enabled: boolean)}
  <button class="cmd" aria-label={t(label)} aria-haspopup="menu" aria-expanded={open === name} disabled={!enabled} onclick={() => (open = open === name ? null : name)}><Icon name={icon} size={20} /><span>{t(label)} ▾</span></button>
{/snippet}

<div class="shape-format">
  <section class="group" aria-label={t('Insert Shapes')}>
    <button class="big" aria-label={t('Shapes')} aria-haspopup="menu" aria-expanded={!!editor.shapeGallery} disabled={!editor.canRun('addSlideShape')} onclick={(event) => editor.openShapeGallery(event.currentTarget)}><Icon name="shapes" size={32} /><span>{t('Shapes')}</span></button>
    <div class="stack">
      <div class="anchor">
        {@render menuButton('editShape', 'shapes', 'Edit Shape', texty)}
        {#if open === 'editShape'}
          <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Edit Shape')}>
            <div class="heading">{t('Change Shape')}</div>
            {#each CHANGE_SHAPES as [preset, label] (preset)}<button role="menuitem" onclick={() => changeShape(preset)}>{t(label)}</button>{/each}
            <hr />
            <button role="menuitem" title={t('Edit Points is not available in this editor yet.')} disabled>{t('Edit Points')}</button>
          </div>
        {/if}
      </div>
      <button class="cmd" aria-label={t('Text Box')} disabled={!editor.canRun('addSlideTextBox')} onclick={() => editor.runOrPrompt('addSlideTextBox', PRESET.textBox)}><Icon name="textbox" size={20} /><span>{t('Text Box')}</span></button>
      <button class="cmd" aria-label={t('Merge Shapes')} title={t('Merging shapes needs boolean geometry, which the library does not compute.')} disabled><Icon name="group" size={20} /><span>{t('Merge Shapes')} ▾</span></button>
    </div>
  </section>

  <section class="group" aria-label={t('Shape Styles')}>
    <ShapeQuickStyles inline />
    <div class="stack">
      <span class="paint"><Icon name="fill" size={18} /><span>{t('Shape Fill')}</span><ShapeFillPicker disabled={!paintable} /></span>
      <span class="paint"><Icon name="outline" size={18} /><span>{t('Shape Outline')}</span><ColorPicker compact label={t('Shape Outline')} disabled={!paintable} choose={outline} /></span>
      <div class="anchor">
        {@render menuButton('shapeEffects', 'shadow', 'Shape Effects', paintable)}
        {#if open === 'shapeEffects'}
          <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Shape Effects')}>
            <button role="menuitem" disabled={!editor.canRun('setShapeShadow')} onclick={() => { open = null; editor.runOrPrompt('setShapeShadow'); }}>{t('Shadow')}</button>
            <button role="menuitem" disabled={!editor.canRun('setShapeGlow')} onclick={() => { open = null; editor.runOrPrompt('setShapeGlow'); }}>{t('Glow')}</button>
            <hr />
            <button role="menuitem" disabled={!editor.canRun('clearShapeEffects')} onclick={() => { open = null; editor.invoke('clearShapeEffects'); }}>{t('No Effects')}</button>
          </div>
        {/if}
      </div>
    </div>
  </section>

  <section class="group" aria-label={t('WordArt Styles')}>
    <div class="anchor">
      <button class="big" bind:this={wordArtButton} aria-label={t('WordArt Quick Styles')} aria-haspopup="menu" aria-expanded={open === 'wordArt'} disabled={!texty} onclick={() => (open = open === 'wordArt' ? null : 'wordArt')}><span class="wordart" aria-hidden="true">A</span><span>{t('Quick Styles')} ▾</span></button>
      {#if open === 'wordArt' && wordArtButton}
        <WordArtGallery anchor={wordArtButton} label={t('WordArt Quick Styles')} choose={wordArt} close={() => (open = null)} clear={clearWordArt} />
      {/if}
    </div>
    <div class="stack">
      <span class="paint"><Icon name="font-color" size={18} /><span>{t('Text Fill')}</span><ColorPicker compact label={t('Text Fill')} disabled={!texty} choose={(color) => textFormat({ color })} /></span>
      <span class="paint"><Icon name="outline" size={18} /><span>{t('Text Outline')}</span><ColorPicker compact label={t('Text Outline')} disabled={!texty} choose={(color) => textFormat({ outline: { color, widthEmu: THIN_OUTLINE_EMU } })} /></span>
      <div class="anchor">
        {@render menuButton('textEffects', 'glow', 'Text Effects', texty)}
        {#if open === 'textEffects'}
          <div class="menu" role="menu" use:placeBelowTrigger aria-label={t('Text Effects')}>
            <button role="menuitem" onclick={() => textFormat({ shadow: TEXT_SHADOW })}>{t('Shadow')}</button>
            <button role="menuitem" onclick={() => textFormat({ reflection: TEXT_REFLECTION })}>{t('Reflection')}</button>
            <button role="menuitem" onclick={() => textFormat({ glow: TEXT_GLOW })}>{t('Glow')}</button>
            <hr />
            <button role="menuitem" onclick={() => textFormat({ shadow: null, reflection: null, glow: null })}>{t('No Effects')}</button>
          </div>
        {/if}
      </div>
    </div>
  </section>

  <section class="group" aria-label={t('Accessibility')}>
    <button class="big" aria-label={t('Alt Text')} disabled={!editor.canRun('setShapeDescription')} onclick={() => editor.runOrPrompt('setShapeDescription')}><Icon name="text-format" size={32} /><span>{t('Alt Text')}</span></button>
  </section>

  <section class="group" aria-label={t('Arrange')}><ArrangeMenu /></section>

  <section class="group size" aria-label={t('Size')}>
    <label><Icon name="height" size={16} /><span class="sr">{t('Height')}</span><input class="ok-input" type="number" min="0" step="any" aria-label={t('Height')} disabled={!editable} value={dimension('h') ?? ''} onchange={(event) => resize('h', event.currentTarget)} /><span>cm</span></label>
    <label><Icon name="width" size={16} /><span class="sr">{t('Width')}</span><input class="ok-input" type="number" min="0" step="any" aria-label={t('Width')} disabled={!editable} value={dimension('w') ?? ''} onchange={(event) => resize('w', event.currentTarget)} /><span>cm</span></label>
  </section>

  <section class="group" aria-label={t('Format Pane')}>
    <button class="big" aria-label={t('Format Pane')} disabled={shapes.length === 0} onclick={() => editor.showShapeFormat()}><Icon name="format-pane" size={32} /><span>{t('Format Pane')}</span></button>
  </section>
</div>

<style>
  .shape-format { display: flex; align-items: stretch; min-width: 0; }
  .group { position: relative; display: flex; align-items: center; gap: 4px; padding: 0 6px; border-right: 1px solid var(--ok-border); flex: none; }
  .group:last-child { border-right: none; }
  button { font: inherit; color: var(--ok-text); background: none; border: 1px solid transparent; border-radius: var(--ok-radius); cursor: pointer; }
  button:hover:not(:disabled) { background: var(--ok-hover); }
  button:disabled { opacity: 0.4; cursor: default; }
  .big { display: flex; flex-direction: column; align-items: center; gap: 2px; min-width: 44px; padding: 3px; font-size: 11px; line-height: 1.15; }
  .big > span:last-child { max-width: 60px; text-align: center; }
  .cmd { display: flex; align-items: center; gap: 4px; padding: 1px 4px; font-size: 11px; white-space: nowrap; }
  .stack { display: flex; flex-direction: column; align-items: flex-start; gap: 1px; }
  .anchor { position: relative; }
  .paint { display: flex; align-items: center; gap: 4px; padding: 1px 4px; font-size: 11px; white-space: nowrap; }
  .wordart { font: 700 26px/32px Georgia, serif; color: var(--ok-accent); text-shadow: 1px 1px 1px #0004; }
  .size { flex-direction: column; justify-content: center; align-items: flex-start; }
  .size label { display: flex; align-items: center; gap: 4px; font-size: 11px; }
  .size input { width: 64px; padding: 2px 3px; font-size: 11px; }
  .sr { position: absolute; width: 1px; height: 1px; overflow: hidden; clip-path: inset(50%); }
  .menu { position: fixed; z-index: 400; display: flex; flex-direction: column; min-width: 220px; padding: 4px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .menu button { padding: 5px 8px; text-align: left; font-size: 12px; white-space: nowrap; }
  .menu .heading { padding: 4px 8px; font-size: 11px; font-weight: 600; color: var(--ok-text-2); }
  .menu hr { width: 100%; border: none; border-top: 1px solid var(--ok-border); margin: 4px 0; }
</style>
