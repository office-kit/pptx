<script lang="ts">
  // Mac PowerPoint's Picture Format tab: Remove Background; Adjust
  // (Corrections, Color, Artistic Effects, Transparency and the Compress /
  // Quality / Change / Reset icons); Picture Styles (style strip, Picture
  // Border, Picture Effects, Picture Layout); Alt Text; Arrange; Size (Crop,
  // Height, Width); Format Pane; Animate as Background. Below 1300 pt the
  // labelled rows become icon menus and Arrange collapses into one button.
  import './contextual.css';
  import { getShapeBounds, getShapeImageArtisticEffect, getShapeImageIntrinsicSize, getShapeKind, getShapeMedia, getShapeStrokeEffective, resetShapeImageColorEffects, setShapeBounds, setShapeImageBrightness, setShapeImageContrast, setShapeImageCrop, setShapeImageOpacity, setShapePreset, type Color, type ImageArtisticEffect, type PresetShape, type SlideShapeData } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import ColorPicker from '../ui/ColorPicker.svelte';
  import Icon from '../ui/Icon.svelte';
  import ArrangeGroup from './ArrangeGroup.svelte';
  import MenuButton from './MenuButton.svelte';
  import PictureStyleGallery from './PictureStyleGallery.svelte';
  import SizeSpinners from './SizeSpinners.svelte';
  import VideoCorrectionsMenu from './VideoCorrectionsMenu.svelte';
  import VideoRecolorMenu from './VideoRecolorMenu.svelte';
  import { captionLines } from './caption.ts';
  import { RibbonCollapse } from './ribbon-collapse.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const collapse = new RibbonCollapse(1);
  const compact = $derived(collapse.step >= 1);
  const picture = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    const shape = shapes.length === 1 ? shapes[0]! : null;
    return shape && getShapeKind(shape) === 'picture' && !getShapeMedia(shape) ? shape : null;
  });
  const editable = $derived(picture !== null && !editor.selectionLocked());
  // PowerPoint's Transparency presets.
  const TRANSPARENCY = [0, 15, 30, 50, 65, 80, 95] as const;
  // PowerPoint's default picture border when a color is chosen first.
  const BORDER_EMU = 12700;
  const CROP_SHAPES: readonly [PresetShape, string][] = [
    ['rect', 'Rectangle'],
    ['roundRect', 'Rounded Rectangle'],
    ['ellipse', 'Oval'],
    ['triangle', 'Isosceles Triangle'],
    ['diamond', 'Diamond'],
    ['pentagon', 'Regular Pentagon'],
    ['hexagon', 'Hexagon'],
    ['star5', '5-Point Star'],
    ['heart', 'Heart'],
  ];

  function transparency(percent: number) {
    editor.invoke('setShapeImageOpacity', { opacity: percent === 0 ? null : 1 - percent / 100 });
  }
  function border(color: Color) {
    const shape = picture;
    if (!shape) return;
    const visible = getShapeStrokeEffective(doc.pres, shape)?.kind === 'solid';
    editor.invoke('setShapeStroke', { options: visible ? { color } : { color, widthEmu: BORDER_EMU } });
  }
  // Mac PowerPoint's Artistic Effects gallery, in order (menus/picture-artistic-effect).
  const ARTISTIC_EFFECTS: readonly [ImageArtisticEffect | null, string][] = [
    [null, 'None'],
    ['marker', 'Marker'],
    ['pencilGrayscale', 'Pencil Grayscale'],
    ['pencilSketch', 'Pencil Sketch'],
    ['lineDrawing', 'Line Drawing'],
    ['chalkSketch', 'Chalk Sketch'],
    ['paintStrokes', 'Paint Strokes'],
    ['paintBrush', 'Paint Brush'],
    ['glowDiffused', 'Glow Diffused'],
    ['blur', 'Blur'],
    ['lightScreen', 'Light Screen'],
    ['watercolorSponge', 'Watercolor Sponge'],
    ['filmGrain', 'Film Grain'],
    ['mosaicBubbles', 'Mosaic Bubbles'],
    ['glass', 'Glass'],
    ['cement', 'Cement'],
    ['texturizer', 'Texturizer'],
    ['crisscrossEtching', 'Crisscross Etching'],
    ['pastelsSmooth', 'Pastels Smooth'],
    ['plasticWrap', 'Plastic Wrap'],
    ['cutout', 'Cutout'],
    ['photocopy', 'Photocopy'],
    ['glowEdges', 'Glow Edges'],
  ];
  const ARTISTIC_UNAVAILABLE = "PowerPoint saves an artistic effect as its own rendering of the picture plus a JPEG XR copy of the original; the editor can show an existing effect but cannot produce either.";
  const artisticEffect = $derived.by(() => {
    doc.version;
    return picture ? getShapeImageArtisticEffect(picture) : null;
  });
  // The picture's natural size at its own resolution, which Reset Picture & Size restores.
  const originalSize = $derived.by(() => {
    doc.version;
    return picture ? getShapeImageIntrinsicSize(picture) : null;
  });

  function resetImage(shape: SlideShapeData) {
    resetShapeImageColorEffects(shape);
    setShapeImageBrightness(shape, null);
    setShapeImageContrast(shape, null);
    setShapeImageOpacity(shape, null);
  }
  function resetPicture() {
    const shape = picture;
    if (!shape || !editable) return;
    doc.transact(t('Reset Picture'), () => resetImage(shape));
  }
  function resetPictureAndSize() {
    const shape = picture;
    const size = originalSize;
    const bounds = shape && getShapeBounds(shape);
    if (!shape || !editable || !size || !bounds) return;
    // The top-left corner stays put, as when the Size group's boxes change.
    doc.transact(t('Reset Picture & Size'), () => {
      resetImage(shape);
      setShapeImageCrop(shape, null);
      setShapeBounds(shape, { ...bounds, w: size.width, h: size.height });
    });
  }
  function cropToShape(preset: PresetShape) {
    const shape = picture;
    if (!shape || !editable) return;
    doc.transact(t('Crop to Shape'), () => setShapePreset(shape, preset));
  }
</script>

{#snippet transparencyItems()}
  {#each TRANSPARENCY as percent (percent)}<button role="menuitem" onclick={() => transparency(percent)}>{percent === 0 ? t('No Transparency') : `${t('Transparency')}: ${percent}%`}</button>{/each}
  <hr />
  <button role="menuitem" onclick={() => editor.showShapeFormat()}>{t('Picture Transparency Options...')}</button>
{/snippet}

{#snippet artisticItems()}
  <div class="artistic-grid" role="group" aria-label={t('Artistic Effect')}>
    {#each ARTISTIC_EFFECTS as [effect, label] (label)}<button role="menuitemradio" aria-checked={artisticEffect === effect} title={t(ARTISTIC_UNAVAILABLE)} disabled>{t(label)}</button>{/each}
  </div>
{/snippet}

{#snippet qualityItems()}
  <button role="menuitem" onclick={() => (editor.activeDialog = 'compressPictures')}>{t('Compress Pictures...')}</button>
  <button role="menuitem" title={t('Upscale Picture enlarges a picture with a Microsoft cloud AI service, which the editor does not have.')} disabled>{t('Upscale Picture')}</button>
{/snippet}

{#snippet changeItems()}
  <button role="menuitem" disabled={!editor.canRun('setShapeImage')} onclick={() => editor.runOrPrompt('setShapeImage')}>{t('From a File...')}</button>
{/snippet}

{#snippet resetItems()}
  <button role="menuitem" onclick={resetPicture}>{t('Reset Picture')}</button>
  <button role="menuitem" title={originalSize ? undefined : t("The picture's format does not record its original size.")} disabled={!originalSize} onclick={resetPictureAndSize}>{t('Reset Picture & Size')}</button>
{/snippet}

{#snippet effectItems()}
  <button role="menuitem" disabled={!editor.canRun('setShapeShadow')} onclick={() => editor.runOrPrompt('setShapeShadow')}>{t('Shadow')}</button>
  <button role="menuitem" disabled={!editor.canRun('setShapeGlow')} onclick={() => editor.runOrPrompt('setShapeGlow')}>{t('Glow')}</button>
  <hr />
  <button role="menuitem" disabled={!editor.canRun('clearShapeEffects')} onclick={() => editor.invoke('clearShapeEffects')}>{t('No Effects')}</button>
{/snippet}

{#snippet cropItems()}
  <button role="menuitem" disabled={!editor.canRun('setShapeImageCrop')} onclick={() => editor.runOrPrompt('setShapeImageCrop')}>{t('Crop')}</button>
  <div class="ctx-heading">{t('Crop to Shape')}</div>
  {#each CROP_SHAPES as [preset, label] (preset)}<button role="menuitem" onclick={() => cropToShape(preset)}>{t(label)}</button>{/each}
{/snippet}

<div class="ctx-ribbon picture-format" bind:this={collapse.node} bind:clientWidth={collapse.width}>
  <section class="ctx-group" aria-label={t('Remove Background')}>
    <button class="ctx-big" aria-label={t('Remove Background')} title={t('Removing a background needs image segmentation, which the editor does not have.')} disabled><span class="ctx-icon-row"><Icon name="remove-background" size={32} /></span><span class="ctx-caption">{captionLines(t('Remove Background'))}</span></button>
  </section>

  <section class="ctx-group" aria-label={t('Adjust')}>
    <VideoCorrectionsMenu target="picture" variant="big" />
    <div class="ctx-rows">
      <VideoRecolorMenu target="picture" variant={compact ? 'icon' : 'row'} />
      <MenuButton look={compact ? 'icon' : 'row'} icon="artistic-effects" label={t('Artistic Effects')} disabled={!picture}>{@render artisticItems()}</MenuButton>
      <MenuButton look={compact ? 'icon' : 'row'} icon="transparency" label={t('Transparency')} disabled={!editable}>{@render transparencyItems()}</MenuButton>
    </div>
    <div class="ctx-rows">
      <button class="ctx-icon" aria-label={t('Compress Pictures')} title={t('Compress Pictures')} disabled={!editable} onclick={() => (editor.activeDialog = 'compressPictures')}><Icon name="compress" size={18} /></button>
      <MenuButton look="icon" icon="picture" label={t('Picture Quality')} disabled={!editable}>{@render qualityItems()}</MenuButton>
      <MenuButton look="icon" icon="replace" label={t('Change Picture')} disabled={!editable}>{@render changeItems()}</MenuButton>
    </div>
    <div class="ctx-rows">
      <MenuButton look="icon" icon="reset" label={t('Reset Picture')} disabled={!editable}>{@render resetItems()}</MenuButton>
    </div>
  </section>

  <section class="ctx-group ctx-shrink" aria-label={t('Picture Styles')}>
    <PictureStyleGallery />
    <div class="ctx-rows">
      <span class="ctx-paint {compact ? 'ctx-icon' : 'ctx-row'}" class:disabled={!editable}><Icon name="outline" size={16} />{#if !compact}<span>{t('Picture Border')}</span>{/if}<ColorPicker compact label={t('Picture Border')} disabled={!editable} choose={border} /></span>
      <MenuButton look={compact ? 'icon' : 'row'} icon="shadow" label={t('Picture Effects')} disabled={!editable}>{@render effectItems()}</MenuButton>
      <MenuButton look={compact ? 'icon' : 'row'} icon="picture-layout" label={t('Picture Layout')} title={t('SmartArt is not supported by the library yet.')} disabled><button role="menuitem" disabled>{t('Picture Layout')}</button></MenuButton>
    </div>
  </section>

  <section class="ctx-group" aria-label={t('Accessibility')}>
    <button class="ctx-big" aria-label={t('Alt Text')} title={t('Display the Alt Text Pane')} disabled={!editor.canRun('setShapeDescription')} onclick={() => editor.runOrPrompt('setShapeDescription')}><span class="ctx-icon-row"><Icon name="alt-text" size={32} /></span><span class="ctx-caption">{captionLines(t('Alt Text'))}</span></button>
  </section>

  <ArrangeGroup collapsed={compact} />

  <section class="ctx-group" aria-label={t('Size')}>
    <MenuButton look="big" icon="crop" label={t('Crop')} disabled={!editable}>{@render cropItems()}</MenuButton>
    <SizeSpinners labels />
  </section>

  <section class="ctx-group" aria-label={t('Format Pane')}>
    <button class="ctx-big" aria-label={t('Format Pane')} title={t('Display the Format Pane')} disabled={!picture} onclick={() => editor.showShapeFormat()}><span class="ctx-icon-row"><Icon name="format-pane" size={32} /></span><span class="ctx-caption">{captionLines(t('Format Pane'))}</span></button>
  </section>

  <section class="ctx-group" aria-label={t('Animate as Background')}>
    <button class="ctx-big" aria-label={t('Animate as Background')} title={t('Animated backgrounds are not supported by the library yet.')} disabled><span class="ctx-icon-row"><Icon name="animate-background" size={32} /></span><span class="ctx-caption">{captionLines(t('Animate as Background'))}</span></button>
  </section>
</div>

<style>
  .artistic-grid { display: grid; grid-template-columns: repeat(5, 76px); }
  .artistic-grid button { height: 40px; white-space: normal; text-align: center; justify-content: center; }
</style>
