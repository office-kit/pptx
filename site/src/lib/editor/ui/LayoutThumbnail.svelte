<script lang="ts">
  // A layout drawn the way PowerPoint's New Slide and Layout galleries show it:
  // the layout (or master) background with its title and content placeholders
  // as boxes of prompt lines. Footer placeholders are left out, as natively.
  import {
    getSlideLayoutBackground,
    getSlideLayoutPlaceholders,
    getPresentationTheme,
    getSlideMasterBackground,
    getSlideSize,
    type PresentationData,
    type SlideLayoutData,
  } from '@office-kit/pptx';

  let { pres, layout }: { pres: PresentationData; layout: SlideLayoutData } = $props();

  const FOOTER_TYPES = new Set(['dt', 'ftr', 'sldNum', 'hdr']);
  const TITLE_TYPES = new Set(['title', 'ctrTitle']);
  const size = $derived(getSlideSize(pres) ?? { width: 12192000, height: 6858000 });
  // Scheme colors as the default color map (bg1 = lt1, tx1 = dk1) resolves them.
  const SCHEME_SLOTS = { bg1: 'light1', lt1: 'light1', tx1: 'dark1', dk1: 'dark1', bg2: 'light2', lt2: 'light2', tx2: 'dark2', dk2: 'dark2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6' } as const;
  const background = $derived.by(() => {
    const own = getSlideLayoutBackground(layout);
    const effective = own.kind === 'inherit' ? getSlideMasterBackground(pres, layout) : own;
    if (effective.kind !== 'solid') return '#ffffff';
    const token = effective.color.replace(/^scheme:/, '');
    const slot = SCHEME_SLOTS[token as keyof typeof SCHEME_SLOTS];
    const resolved = slot ? getPresentationTheme(pres)?.[slot] : effective.color;
    return resolved && /^#?[0-9a-f]{6}$/i.test(resolved) ? (resolved.startsWith('#') ? resolved : `#${resolved}`) : '#ffffff';
  });
  // Prompt lines contrast with the background, as PowerPoint's do.
  const ink = $derived.by(() => {
    const hex = background.replace('#', '');
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    return 0.299 * r! + 0.587 * g! + 0.114 * b! > 140 ? '#595959' : '#d9d9d9';
  });
  // A placeholder without its own xfrm sits where the master puts it; the
  // library does not expose master placeholders, so use the default Office
  // master's title and body areas, as fractions of the slide.
  const MASTER_TITLE = { x: 0.0688, y: 0.0532, w: 0.8625, h: 0.1933 };
  const MASTER_BODY = { x: 0.0688, y: 0.2662, w: 0.8625, h: 0.6345 };
  const boxes = $derived(getSlideLayoutPlaceholders(layout).flatMap((placeholder) => {
    if (FOOTER_TYPES.has(placeholder.type ?? '')) return [];
    const title = TITLE_TYPES.has(placeholder.type ?? '');
    if (placeholder.bounds) return [{ ...placeholder.bounds, title }];
    const area = title ? MASTER_TITLE : MASTER_BODY;
    return [{ x: area.x * size.width, y: area.y * size.height, w: area.w * size.width, h: area.h * size.height, title }];
  }));
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
