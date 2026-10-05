<script lang="ts">
  // PowerPoint's WordArt gallery: twenty "A" swatches, five to a row, shared by
  // Shape Format ▸ WordArt Quick Styles and Insert ▸ WordArt.
  import { onMount } from 'svelte';
  import { asColor, getPresentationTheme, type Color, type ColorTransform, type PatternPreset, type TextFormat } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { resolveColor } from '../core/theme-color.ts';
  import { WORDART_PRESETS, wordArtFormat, type WordArtPreset } from '../core/wordart-presets.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { label, choose, close, clear, anchor }: {
    label: string;
    choose: (preset: WordArtPreset) => void;
    close: () => void;
    /** Clear WordArt, under the swatches; Insert ▸ WordArt has none. */
    clear?: () => void;
    /** The button the gallery opens under. */
    anchor: HTMLElement;
  } = $props();

  const editor = getEditor();
  const theme = $derived.by(() => { editor.doc.version; return getPresentationTheme(editor.doc.pres); });
  const COLUMNS = 5;
  // Swatch scale: one point of the preset draws one CSS pixel.
  const EMU_PER_PX = 12700;
  let menu = $state<HTMLDivElement>();

  const paint = (color: Color, transforms: readonly ColorTransform[] = []) => resolveColor(color, transforms, theme) ?? '#000000';
  const rgba = (hex: string, opacity = 1) => {
    const [r, g, b] = [1, 3, 5].map((at) => parseInt(hex.slice(at, at + 2), 16));
    return `rgba(${r}, ${g}, ${b}, ${opacity})`;
  };
  const px = (emu = 0) => `${emu / EMU_PER_PX}px`;
  // CSS stand-ins for the preset patterns the gallery uses.
  const PATTERNS: Partial<Record<PatternPreset, (fg: string, bg: string) => string>> = {
    dkUpDiag: (fg, bg) => `repeating-linear-gradient(-45deg, ${fg} 0 2px, ${bg} 2px 4px)`,
    ltDnDiag: (fg, bg) => `repeating-linear-gradient(45deg, ${fg} 0 1px, ${bg} 1px 4px)`,
    narHorz: (fg, bg) => `repeating-linear-gradient(0deg, ${fg} 0 1px, ${bg} 1px 3px)`,
    pct50: (fg, bg) => `repeating-conic-gradient(${fg} 0 25%, ${bg} 0 50%) 0 0 / 2px 2px`,
  };

  function fill(format: TextFormat): string {
    const textFill = format.textFill;
    if (textFill?.kind === 'gradient') {
      // DrawingML measures from 3 o'clock, CSS from 12.
      const stops = textFill.stops.map((stop) => `${paint(stop.color, stop.colorTransforms)} ${stop.offset * 100}%`);
      return `linear-gradient(${(textFill.angleDeg ?? 90) + 90}deg, ${stops.join(', ')})`;
    }
    if (textFill?.kind === 'pattern') {
      const fg = paint(asColor(textFill.foreground) ?? 'tx1', textFill.foregroundTransforms);
      const bg = paint(asColor(textFill.background) ?? 'bg1', textFill.backgroundTransforms);
      return PATTERNS[textFill.preset]?.(fg, bg) ?? fg;
    }
    // A preset without a fill keeps the inherited text color, Text 1 here.
    return paint(format.color ?? 'tx1');
  }

  function swatch(preset: WordArtPreset): string {
    const format = wordArtFormat(preset, theme);
    const filters: string[] = [];
    if (format.shadow) {
      const { offsetEmu = 0, angleDeg = 45, blurEmu = 0, color = '#000000', opacity } = format.shadow;
      const radians = (angleDeg * Math.PI) / 180;
      filters.push(`drop-shadow(${px(offsetEmu * Math.cos(radians))} ${px(offsetEmu * Math.sin(radians))} ${px(blurEmu)} ${rgba(paint(color), opacity)})`);
    }
    if (format.glow) {
      const glow = rgba(paint(format.glow.color), format.glow.opacity);
      filters.push(`drop-shadow(0 0 ${px(format.glow.radiusEmu)} ${glow})`, `drop-shadow(0 0 ${px(format.glow.radiusEmu)} ${glow})`);
    }
    // CSS has no inner shadow for glyphs; those swatches show fill and outline only.
    const outline = format.outline?.color && format.outline.widthEmu ? `${px(format.outline.widthEmu)} ${paint(format.outline.color)}` : '0';
    return [
      // The shorthand resets the clip, so the clip follows it here.
      `background: ${fill(format)}`,
      'background-clip: text',
      '-webkit-background-clip: text',
      `font-weight: ${format.bold ? 700 : 400}`,
      `-webkit-text-stroke: ${outline}`,
      `filter: ${filters.join(' ') || 'none'}`,
      format.reflection ? `-webkit-box-reflect: below 0 linear-gradient(transparent 55%, ${rgba('#000000', format.reflection.startOpacity)})` : '',
    ].filter(Boolean).join('; ');
  }

  function keys(event: KeyboardEvent) {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(); return; }
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -COLUMNS, ArrowDown: COLUMNS };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')];
    const enabled = items.filter((item) => !item.disabled);
    if (event.key === 'Home' || event.key === 'End') { (event.key === 'Home' ? enabled[0] : enabled.at(-1))?.focus(); return; }
    // Step over disabled swatches in the direction of travel; Down from the
    // last row reaches Clear WordArt below the grid.
    const step = offsets[event.key]!;
    for (let index = items.indexOf(event.target as HTMLButtonElement) + step; index >= 0 && index < items.length; index += step) {
      if (!items[index]!.disabled) { items[index]!.focus(); return; }
    }
    if (step > 0) enabled.at(-1)?.focus();
  }

  // Fixed, so the ribbon panel's overflow does not clip it.
  function place(node: HTMLElement) {
    const bounds = anchor.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${bounds.bottom + 2}px`;
  }
  onMount(() => menu?.querySelector<HTMLButtonElement>('[role="menuitem"]:not(:disabled)')?.focus());
</script>

<svelte:window onpointerdown={(event) => { if (!menu?.contains(event.target as Node) && !anchor.contains(event.target as Node)) close(); }} />

<div class="wordart-gallery" role="menu" aria-label={label} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
  <div class="grid">
    {#each WORDART_PRESETS as preset (preset.label)}
      <button role="menuitem" aria-label={t(preset.label)} title={preset.unavailable ? t(preset.unavailable) : t(preset.label)} disabled={!!preset.unavailable} onclick={() => choose(preset)}>
        <span class="letter" aria-hidden="true" style={swatch(preset)}>A</span>
      </button>
    {/each}
  </div>
  {#if clear}
    <hr />
    <button role="menuitem" class="clear" onclick={clear}>{t('Clear WordArt')}</button>
  {/if}
</div>

<style>
  .wordart-gallery { position: fixed; z-index: 450; padding: 6px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); box-shadow: var(--ok-shadow-lg); }
  .grid { display: grid; grid-template-columns: repeat(5, 48px); gap: 4px; }
  button { font: inherit; color: var(--ok-text); border: 1px solid transparent; border-radius: 3px; cursor: pointer; }
  .grid button { display: flex; align-items: center; justify-content: center; width: 48px; height: 48px; padding: 0; overflow: hidden; background: #fff; }
  .grid button:hover:not(:disabled), .grid button:focus-visible { outline: 2px solid var(--ok-accent); outline-offset: -2px; }
  .grid button:disabled { opacity: 0.4; cursor: default; }
  .letter { font: 34px/1 Calibri, Carlito, Arial, sans-serif; color: transparent; }
  hr { border: none; border-top: 1px solid var(--ok-border); margin: 6px 0 4px; }
  .clear { width: 100%; padding: 5px 8px; text-align: left; font-size: 12px; white-space: nowrap; background: none; }
  .clear:hover { background: var(--ok-hover); }
</style>
