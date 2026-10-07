<script lang="ts">
  // A layout drawn the way PowerPoint's New Slide and Layout galleries show it:
  // the layout (or master) background with its title and content placeholders
  // as boxes of prompt lines. Footer placeholders are left out, as natively.
  import { getSlideMasterPartName, getSlideSize, type PresentationData, type SlideLayoutData } from '@office-kit/pptx';
  import { contrastInk, layoutBoxes, masterBoxes, solidBackground } from '../core/master-geometry.ts';

  let { pres, layout }: { pres: PresentationData; layout: SlideLayoutData } = $props();

  const FOOTER_TYPES = new Set(['dt', 'ftr', 'sldNum', 'hdr']);
  const TITLE_TYPES = new Set(['title', 'ctrTitle']);
  const size = $derived(getSlideSize(pres) ?? { width: 12192000, height: 6858000 });
  const background = $derived(solidBackground(pres, layout));
  // Prompt lines contrast with the background, as PowerPoint's do.
  const ink = $derived(contrastInk(background, '#595959', '#d9d9d9'));
  const boxes = $derived.by(() => {
    const master = getSlideMasterPartName(layout);
    return layoutBoxes(layout, size, master ? masterBoxes(pres, master, size) : []).flatMap((box) =>
      FOOTER_TYPES.has(box.type ?? '')
        ? []
        : [{ x: box.x * size.width, y: box.y * size.height, w: box.w * size.width, h: box.h * size.height, title: TITLE_TYPES.has(box.type ?? '') }],
    );
  });
</script>

<svg class="layout-thumbnail" viewBox="0 0 {size.width} {size.height}" aria-hidden="true">
  <rect width={size.width} height={size.height} fill={background} />
  {#each boxes as box, i (i)}
    <rect x={box.x} y={box.y} width={box.w} height={box.h} fill="none" stroke={ink} stroke-width={size.width / 400} stroke-dasharray="{size.width / 120} {size.width / 160}" />
    {#if box.title}
      <rect x={box.x + box.w * 0.25} y={box.y + box.h * 0.35} width={box.w * 0.5} height={Math.min(box.h * 0.3, size.height / 25)} fill={ink} />
    {:else}
      {#each [0.15, 0.32, 0.49] as offset (offset)}
        {#if offset * box.h + size.height / 60 < box.h}<rect x={box.x + box.w * 0.06} y={box.y + box.h * offset} width={box.w * (offset === 0.49 ? 0.5 : 0.8)} height={size.height / 60} fill={ink} />{/if}
      {/each}
    {/if}
  {/each}
</svg>

<style>
  .layout-thumbnail { display: block; width: 100%; height: auto; border: 1px solid var(--ok-border); }
</style>
