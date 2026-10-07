<script lang="ts">
  // A slide master or one of its layouts as Mac PowerPoint's Slide Master view
  // draws it: the background, and every placeholder as a dashed box holding
  // its prompt text in the theme fonts at the default master's sizes.
  import { getPresentationFonts, getSlideSize, type PresentationData, type SlideLayoutData } from '@office-kit/pptx';
  import { contrastInk, layoutBoxes, MASTER_BOXES, solidBackground, type MasterBox } from '../core/master-geometry.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { pres, layout, master = false, version = 0, pixelWidth, outlines = true }: {
    pres: PresentationData;
    /** The layout to draw, or (with `master`) any layout of the master to draw. */
    layout: SlideLayoutData;
    master?: boolean;
    /** The document version, so edits redraw. */
    version?: number;
    /** The drawn width in screen px, for 1 px dashed outlines at any size. */
    pixelWidth: number;
    /** The notes master's slide image shows the prompt text without the boxes. */
    outlines?: boolean;
  } = $props();

  const EMU_PER_PT = 12700;
  // The default Office master: 44 pt titles; 28/24/20/20/20 pt body levels,
  // 0.25 in bullet hang and 0.5 in per level; 12 pt footers; 0.1 × 0.05 in insets.
  const TITLE_PT = 44;
  const CENTERED_TITLE_PT = 60;
  const SUBTITLE_PT = 24;
  const LEVEL_PT = [28, 24, 20, 20, 20];
  const FOOTER_PT = 12;
  const BULLET_HANG = 18;
  const LEVEL_STEP = 36;
  const INSET_X = 7.2;
  const INSET_Y = 3.6;
  const LINE_SPACING = 0.9;
  const SPACE_BEFORE_PT = 10;

  const size = $derived.by(() => { version; return getSlideSize(pres) ?? { width: 12192000, height: 6858000 }; });
  // Drawn in points: Chromium does not draw text at EMU-sized font sizes.
  const width = $derived(size.width / EMU_PER_PT);
  const height = $derived(size.height / EMU_PER_PT);
  const scale = $derived(pixelWidth / width);
  const fonts = $derived.by(() => { version; return getPresentationFonts(pres); });
  const background = $derived.by(() => { version; return solidBackground(pres, layout, master); });
  const ink = $derived(contrastInk(background));
  const boxes = $derived.by<readonly MasterBox[]>(() => { version; return master ? MASTER_BOXES : layoutBoxes(layout, size); });
  const majorFont = $derived(`'${fonts?.majorLatin ?? 'Aptos Display'}', system-ui, sans-serif`);
  const minorFont = $derived(`'${fonts?.minorLatin ?? 'Aptos'}', system-ui, sans-serif`);
  const today = new Date().toLocaleDateString();
  const levels = $derived([t('Click to edit Master text styles'), t('Second level'), t('Third level'), t('Fourth level'), t('Fifth level')]);

  function rect(box: MasterBox) {
    return { x: box.x * width, y: box.y * height, w: box.w * width, h: box.h * height };
  }
  function bodyLines(box: MasterBox) {
    const area = rect(box);
    let y = area.y + INSET_Y;
    return levels.map((text, level) => {
      const pt = LEVEL_PT[level]!;
      y += (level ? SPACE_BEFORE_PT : 0) + pt * LINE_SPACING * 1.2;
      return { text, level, pt, x: area.x + INSET_X + BULLET_HANG + level * LEVEL_STEP, y };
    }).filter((line) => line.y < area.y + area.h);
  }
</script>

<svg class="master-slide" viewBox="0 0 {width} {height}" aria-hidden="true">
  <rect {width} {height} fill={background} />
  {#each boxes as box, i (i)}
    {@const area = rect(box)}
    {#if outlines}<rect class="placeholder" data-role={box.role} x={area.x} y={area.y} width={area.w} height={area.h} fill="none" stroke="#8c8c8c" stroke-width={1 / scale} stroke-dasharray="{4 / scale} {3 / scale}" />{/if}
    {#if box.role === 'title'}
      {@const centered = box.type === 'ctrTitle'}
      {@const pt = centered ? CENTERED_TITLE_PT : TITLE_PT}
      <text x={centered ? area.x + area.w / 2 : area.x + INSET_X} y={centered ? area.y + area.h - INSET_Y - pt  * 0.25 : area.y + area.h / 2} text-anchor={centered ? 'middle' : 'start'} dominant-baseline={centered ? 'alphabetic' : 'central'} font-family={majorFont} font-size={pt } fill={ink}>{t('Click to edit Master title style')}</text>
    {:else if box.role === 'subtitle'}
      <text x={area.x + area.w / 2} y={area.y + INSET_Y + SUBTITLE_PT  * 1.1} text-anchor="middle" font-family={minorFont} font-size={SUBTITLE_PT } fill={ink}>{t('Click to edit Master subtitle style')}</text>
    {:else if box.role === 'body' || box.role === 'object'}
      {#each bodyLines(box) as line (line.level)}
        <text x={line.x} y={line.y} font-family={minorFont} font-size={line.pt } fill={ink}><tspan x={line.x - BULLET_HANG}>•</tspan><tspan x={line.x}>{line.text}</tspan></text>
      {/each}
    {:else if box.role === 'dt' || box.role === 'ftr' || box.role === 'sldNum'}
      {@const anchor = box.role === 'dt' ? 'start' : box.role === 'ftr' ? 'middle' : 'end'}
      {@const x = box.role === 'dt' ? area.x + INSET_X : box.role === 'ftr' ? area.x + area.w / 2 : area.x + area.w - INSET_X}
      <text {x} y={area.y + area.h / 2} text-anchor={anchor} dominant-baseline="central" font-family={minorFont} font-size={FOOTER_PT } fill={ink} fill-opacity="0.75">{box.role === 'dt' ? today : box.role === 'ftr' ? t('Footer') : '‹#›'}</text>
    {/if}
  {/each}
</svg>

<style>
  .master-slide { display: block; width: 100%; height: 100%; }
</style>
