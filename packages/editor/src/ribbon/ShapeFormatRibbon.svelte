<script lang="ts">
  // The reference desktop app's (Mac) Shape Format tab, in its order and sizes: Insert Shapes
  // (shape strip, Text Box ▾, Edit Shape ▾, Merge Shapes ▾), Shape Styles
  // (style strip, Shape Fill, Shape Outline ▾, Shape Effects ▾), Text Art Styles
  // (Text Art strip, Text Fill, Text Outline ▾, Text Effects ▾), Alt Text,
  // Arrange, Size and Format Pane. Below 1300 pt Insert Shapes becomes Shapes
  // plus three icon menus and Arrange collapses into one button.
  import { addSlideTextBox, getPresentationTheme, getShapeId, getShapeKind, inches, setShapePreset, setShapeText3D, setShapeTextDirection, setShapeTextFormat, type Color, type PresetShape, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { SHAPE_GALLERY, galleryShapeLabel, shapeSprite, type GalleryShape } from '../core/shape-gallery.ts';
  import { TEXT_ART_PRESETS, applyTextArtPreset, type TextArtPreset } from '../core/text-art-presets.ts';
  import { textArtSwatchStyle } from '../core/text-art-swatch.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import ColorPicker from '../ui/ColorPicker.svelte';
  import Icon from '../ui/Icon.svelte';
  import TextArtGallery from '../ui/TextArtGallery.svelte';
  import ArrangeGroup from './ArrangeGroup.svelte';
  import MenuButton from './MenuButton.svelte';
  import ShapeFillPicker from './ShapeFillPicker.svelte';
  import ShapeQuickStyles from './ShapeQuickStyles.svelte';
  import SizeSpinners from './SizeSpinners.svelte';
  import { captionLines } from './caption.ts';
  import { PRESET } from './config.ts';
  import { RibbonCollapse } from './ribbon-collapse.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const collapse = new RibbonCollapse(1);
  const compact = $derived(collapse.step >= 1);
  const shapes = $derived.by(() => { doc.version; return editor.selectedShapes(); });
  const editable = $derived(shapes.length > 0 && !editor.selectionLocked());
  const paintable = $derived(editable && shapes.every((shape) => ['shape', 'connector'].includes(getShapeKind(shape))));
  const texty = $derived(editable && shapes.some((shape) => getShapeKind(shape) === 'shape'));
  const theme = $derived.by(() => { doc.version; return getPresentationTheme(doc.pres); });
  let textArtOpen = $state(false);
  let textArtButton = $state<HTMLButtonElement>();

  // The in-ribbon shape strip: six columns by three rows per page, like
  // the reference desktop app's 120 × 60 pt gallery.
  const STRIP_COLUMNS = 6;
  const STRIP_ROWS = 3;
  const STRIP_CELL = { w: 20, h: 18, icon: 16 };
  const stripShapes = SHAPE_GALLERY.flatMap((group) => group.shapes);
  const stripPages = Math.ceil(stripShapes.length / (STRIP_COLUMNS * STRIP_ROWS));
  let stripPage = $state(0);
  const stripVisible = $derived(stripShapes.slice(stripPage * STRIP_COLUMNS * STRIP_ROWS, (stripPage + 1) * STRIP_COLUMNS * STRIP_ROWS));
  const sprite = $derived(compact ? null : shapeSprite());

  // The in-ribbon Text Art strip shows three swatches; the arrow opens all twenty.
  const TEXT_ART_STRIP = TEXT_ART_PRESETS.slice(0, 3);

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
  // The reference desktop app's outer shadow and glow presets for text ("Offset: Bottom Right",
  // "Glow: 5 point; Accent color 1").
  const TEXT_SHADOW = { color: '#000000', blurEmu: 38100, offsetEmu: 38100, angleDeg: 45, opacity: 0.4 } as const;
  const TEXT_GLOW = { color: 'accent1', radiusEmu: 63500, opacity: 0.4 } as const;
  const TEXT_REFLECTION = { blurEmu: 6350, offsetEmu: 0, angleDeg: 90, startOpacity: 0.5, opacity: 0.003, endPosition: 0.55 } as const;
  const THIN_OUTLINE_EMU = 9525;
  // Insert ▸ Vertical Text Box drops a tall box, as the reference desktop app does on a click.
  const VERTICAL_TEXT_BOX = { x: inches(2), y: inches(1.5), w: inches(1), h: inches(4) };

  function textFormat(format: TextFormat) {
    editor.invoke('setShapeTextFormat', { format, options: { reset: false } });
  }
  function textArt(preset: TextArtPreset) {
    textArtOpen = false;
    doc.transact(t('Text Art Styles'), () => { for (const shape of shapes) if (getShapeKind(shape) === 'shape') applyTextArtPreset(shape, preset); });
  }
  // Clear Text Art removes the run effects and outline and the bevel's 3-D, and keeps the fill.
  function clearTextArt() {
    textArtOpen = false;
    doc.transact(t('Clear Text Art'), () => {
      for (const shape of shapes) {
        if (getShapeKind(shape) !== 'shape') continue;
        setShapeTextFormat(shape, { outline: null, shadow: null, innerShadow: null, glow: null, reflection: null });
        setShapeText3D(shape, null);
      }
    });
  }
  function changeShape(preset: PresetShape) {
    doc.transact(t('Change Shape'), () => { for (const shape of shapes) if (getShapeKind(shape) === 'shape') setShapePreset(shape, preset); });
  }
  function outline(color: Color) { editor.invoke('setShapeStroke', { options: { color } }); }
  function drawShape(preset: GalleryShape) {
    editor.shapeGallery = null;
    editor.drawShape = preset;
  }
  function verticalTextBox() {
    const slide = doc.currentSlide;
    if (!slide) return;
    const index = doc.selection.slideIndex;
    doc.transact(t('Vertical Text Box'), () => {
      const shape = addSlideTextBox(slide, { ...VERTICAL_TEXT_BOX, text: t('Text') });
      setShapeTextDirection(shape, 'eaVert');
      doc.selectShape(index, getShapeId(shape));
    });
  }
</script>

{#snippet textBoxItems()}
  <button role="menuitem" disabled={!editor.canRun('addSlideTextBox')} onclick={() => editor.runOrPrompt('addSlideTextBox', PRESET.textBox)}>{t('Draw Horizontal Text Box')}</button>
  <button role="menuitem" disabled={!doc.currentSlide} onclick={verticalTextBox}>{t('Vertical Text Box')}</button>
{/snippet}

{#snippet editShapeItems()}
  <div class="ctx-heading">{t('Change Shape')}</div>
  {#each CHANGE_SHAPES as [preset, label] (preset)}<button role="menuitem" onclick={() => changeShape(preset)}>{t(label)}</button>{/each}
  <hr />
  <button role="menuitem" title={t('Edit Points is not available in this editor yet.')} disabled>{t('Edit Points')}</button>
{/snippet}

{#snippet mergeItems()}
  <button role="menuitem" disabled>{t('Union')}</button>
{/snippet}

{#snippet shapeEffectItems()}
  <button role="menuitem" disabled={!editor.canRun('setShapeShadow')} onclick={() => editor.runOrPrompt('setShapeShadow')}>{t('Shadow')}</button>
  <button role="menuitem" disabled={!editor.canRun('setShapeGlow')} onclick={() => editor.runOrPrompt('setShapeGlow')}>{t('Glow')}</button>
  <hr />
  <button role="menuitem" disabled={!editor.canRun('clearShapeEffects')} onclick={() => editor.invoke('clearShapeEffects')}>{t('No Effects')}</button>
{/snippet}

{#snippet textEffectItems()}
  <button role="menuitem" onclick={() => textFormat({ shadow: TEXT_SHADOW })}>{t('Shadow')}</button>
  <button role="menuitem" onclick={() => textFormat({ reflection: TEXT_REFLECTION })}>{t('Reflection')}</button>
  <button role="menuitem" onclick={() => textFormat({ glow: TEXT_GLOW })}>{t('Glow')}</button>
  <hr />
  <button role="menuitem" onclick={() => textFormat({ shadow: null, reflection: null, glow: null })}>{t('No Effects')}</button>
{/snippet}

<div class="ctx-ribbon shape-format" bind:this={collapse.node} bind:clientWidth={collapse.width}>
  <section class="ctx-group" aria-label={t('Insert Shapes')}>
    {#if compact}
      <button class="ctx-big" aria-label={t('Shapes')} aria-haspopup="menu" aria-expanded={!!editor.shapeGallery} disabled={!editor.canRun('addSlideShape')} onclick={(event) => editor.openShapeGallery(event.currentTarget)}><span class="ctx-icon-row"><Icon name="shapes" size={32} /><span class="ctx-caret" aria-hidden="true">▾</span></span><span class="ctx-caption">{t('Shapes')}</span></button>
      <div class="ctx-rows">
        <MenuButton look="icon" icon="textbox" label={t('Text Box')}>{@render textBoxItems()}</MenuButton>
        <MenuButton look="icon" icon="shapes" label={t('Edit Shape')} disabled={!texty}>{@render editShapeItems()}</MenuButton>
        <MenuButton look="icon" icon="group" label={t('Merge Shapes')} title={t('Merging shapes needs boolean geometry, which the library does not compute.')} disabled>{@render mergeItems()}</MenuButton>
      </div>
    {:else}
      <div class="ctx-gallery" role="group" aria-label={t('Shapes')}>
        <button class="ctx-gallery-arrow" class:hidden={stripPage === 0} aria-label={t('Previous Shapes gallery')} onclick={() => (stripPage = Math.max(0, stripPage - 1))}>‹</button>
        <div class="ctx-gallery-items shape-strip" style:width="{STRIP_COLUMNS * STRIP_CELL.w}px">
          {#each stripVisible as preset (preset)}
            {@const cell = sprite?.cells.get(preset) ?? 0}
            <button
              aria-label={t(galleryShapeLabel(preset))}
              title={t(galleryShapeLabel(preset))}
              disabled={!editor.canRun('addSlideShape')}
              style:background-image={sprite ? `url('${sprite.url}')` : undefined}
              style:background-size={sprite ? `${sprite.columns * STRIP_CELL.icon}px ${sprite.rows * STRIP_CELL.icon}px` : undefined}
              style:background-position={sprite ? `-${(cell % sprite.columns) * STRIP_CELL.icon}px -${Math.floor(cell / sprite.columns) * STRIP_CELL.icon}px` : undefined}
              onclick={() => drawShape(preset)}
            ></button>
          {/each}
        </div>
        <button class="ctx-gallery-arrow" class:hidden={stripPage === stripPages - 1} aria-label={t('Next Shapes gallery')} onclick={() => (stripPage = Math.min(stripPages - 1, stripPage + 1))}>›</button>
      </div>
      <div class="ctx-rows">
        <MenuButton look="row" icon="textbox" label={t('Text Box')}>{@render textBoxItems()}</MenuButton>
        <MenuButton look="row" icon="shapes" label={t('Edit Shape')} disabled={!texty}>{@render editShapeItems()}</MenuButton>
        <MenuButton look="row" icon="group" label={t('Merge Shapes')} title={t('Merging shapes needs boolean geometry, which the library does not compute.')} disabled>{@render mergeItems()}</MenuButton>
      </div>
    {/if}
  </section>

  <section class="ctx-group ctx-shrink" aria-label={t('Shape Styles')}>
    <ShapeQuickStyles inline />
    <span class="ctx-paint ctx-big" class:disabled={!paintable}><span class="ctx-icon-row"><Icon name="fill" size={32} /></span><span class="ctx-caption">{captionLines(t('Shape Fill'))}</span><ShapeFillPicker disabled={!paintable} /></span>
    <div class="ctx-rows tools">
      <span class="ctx-paint ctx-tool" class:disabled={!paintable}><Icon name="outline" size={18} /><ColorPicker compact label={t('Shape Outline')} disabled={!paintable} choose={outline} /></span>
      <MenuButton look="tool" icon="shadow" label={t('Shape Effects')} disabled={!paintable}>{@render shapeEffectItems()}</MenuButton>
    </div>
  </section>

  <section class="ctx-group ctx-shrink" aria-label={t('Text Art Styles')}>
    <div class="ctx-gallery" role="group" aria-label={t('Text Art Styles')}>
      <span class="ctx-gallery-arrow hidden" aria-hidden="true"></span>
      <div class="ctx-gallery-items text-art-strip">
        {#each TEXT_ART_STRIP as preset (preset.label)}
          <button aria-label={t(preset.label)} title={t(preset.label)} disabled={!texty} onclick={() => textArt(preset)}><span class="letter" aria-hidden="true" style={textArtSwatchStyle(preset, theme)}>A</span></button>
        {/each}
      </div>
      <button class="ctx-gallery-arrow" bind:this={textArtButton} aria-label={t('Text Art Quick Styles')} title={t('Text Art Quick Styles')} aria-haspopup="menu" aria-expanded={textArtOpen} disabled={!texty} onclick={() => (textArtOpen = !textArtOpen)}>›</button>
      {#if textArtOpen && textArtButton}
        <TextArtGallery anchor={textArtButton} label={t('Text Art Quick Styles')} choose={textArt} close={() => (textArtOpen = false)} clear={clearTextArt} />
      {/if}
    </div>
    <span class="ctx-paint ctx-big" class:disabled={!texty}><span class="ctx-icon-row"><Icon name="font-color" size={32} /></span><span class="ctx-caption">{captionLines(t('Text Fill'))}</span><ColorPicker compact label={t('Text Fill')} disabled={!texty} choose={(color) => textFormat({ color })} /></span>
    <div class="ctx-rows tools">
      <span class="ctx-paint ctx-tool" class:disabled={!texty}><Icon name="outline" size={18} /><ColorPicker compact label={t('Text Outline')} disabled={!texty} choose={(color) => textFormat({ outline: { color, widthEmu: THIN_OUTLINE_EMU } })} /></span>
      <MenuButton look="tool" icon="glow" label={t('Text Effects')} disabled={!texty}>{@render textEffectItems()}</MenuButton>
    </div>
  </section>

  <section class="ctx-group" aria-label={t('Accessibility')}>
    <button class="ctx-big" aria-label={t('Alt Text')} title={t('Display the Alt Text Pane')} disabled={!editor.canRun('setShapeDescription')} onclick={() => editor.runOrPrompt('setShapeDescription')}><span class="ctx-icon-row"><Icon name="alt-text" size={32} /></span><span class="ctx-caption">{captionLines(t('Alt Text'))}</span></button>
  </section>

  <ArrangeGroup collapsed={compact} />

  <section class="ctx-group" aria-label={t('Size')}>
    <SizeSpinners />
  </section>

  <section class="ctx-group" aria-label={t('Format Pane')}>
    <button class="ctx-big" aria-label={t('Format Pane')} title={t('Display the Format Pane')} disabled={shapes.length === 0} onclick={() => editor.showShapeFormat()}><span class="ctx-icon-row"><Icon name="format-pane" size={32} /></span><span class="ctx-caption">{captionLines(t('Format Pane'))}</span></button>
  </section>
</div>

<style>
  .ctx-gallery-items.shape-strip { display: grid; grid-template-columns: repeat(6, 20px); grid-auto-rows: 18px; align-content: center; box-sizing: border-box; background: #fff; }
  .shape-strip button { width: 20px; height: 18px; padding: 1px 2px; border-radius: 2px; background-repeat: no-repeat; background-origin: content-box; box-sizing: border-box; }
  .ctx-gallery-items.text-art-strip { width: 174px; justify-content: space-around; background: #fff; }
  .text-art-strip button { display: flex; align-items: center; justify-content: center; width: 56px; height: 54px; padding: 0; overflow: hidden; }
  .text-art-strip button:hover:not(:disabled) { outline: 2px solid var(--ok-accent); outline-offset: -2px; background: none; }
  .letter { font: 34px/1 Calibri, Carlito, Arial, sans-serif; color: transparent; }
  .ctx-gallery { position: relative; }
</style>
