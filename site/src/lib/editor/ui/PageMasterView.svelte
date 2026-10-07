<script lang="ts">
  // View ▸ Handout Master and Notes Master. The library reads and writes
  // neither master yet, so the page shows the default Office master — header,
  // date, footer and page number in the corners; six slide frames on the
  // handout, the slide image and five body levels on the notes page — as Mac
  // PowerPoint draws it, without editing.
  import { getSlideLayouts } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { HANDOUT_AREAS, HANDOUT_SLIDE_FRAMES, NOTES_AREAS, PAGE_WIDTH_PT, type Area } from '../core/master-geometry.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import MasterSlide from './MasterSlide.svelte';
  import PageFrame from './PageFrame.svelte';

  let { kind }: { kind: 'handoutMaster' | 'notesMaster' } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const place = (area: Area) => `left:${area.x * 100}%;top:${area.y * 100}%;width:${area.w * 100}%;height:${area.h * 100}%`;
  const today = new Date().toLocaleDateString();
  const corners = $derived(kind === 'handoutMaster' ? HANDOUT_AREAS : NOTES_AREAS);
  const layout = $derived.by(() => { doc.version; return getSlideLayouts(doc.pres)[0] ?? null; });
  const levels = $derived([t('Click to edit Master text styles'), t('Second level'), t('Third level'), t('Fourth level'), t('Fifth level')]);
</script>

<PageFrame label={t(kind === 'handoutMaster' ? 'Handout Master' : 'Notes Master')}>
  <div class="placeholder corner" data-role="hdr" style={place(corners.hdr)}>{t('Header')}</div>
  <div class="placeholder corner end" data-role="dt" style={place(corners.dt)}>{today}</div>
  {#if kind === 'handoutMaster'}
    {#each HANDOUT_SLIDE_FRAMES as frame, i (i)}<div class="frame" data-role="slide" style={place(frame)}></div>{/each}
  {:else}
    <div class="slide-image" data-role="slide" style={place(NOTES_AREAS.slideImage)}>{#if layout}<MasterSlide pres={doc.pres} {layout} master outlines={false} version={doc.version} pixelWidth={NOTES_AREAS.slideImage.w * PAGE_WIDTH_PT * editor.pageZoom} />{/if}</div>
    <div class="placeholder body" data-role="body" style={place(NOTES_AREAS.body)}>{#each levels as text, level (level)}<p style:padding-left="calc({level} * 36px * var(--page-scale))">{text}</p>{/each}</div>
  {/if}
  <div class="placeholder corner bottom" data-role="ftr" style={place(corners.ftr)}>{t('Footer')}</div>
  <div class="placeholder corner end bottom" data-role="sldNum" style={place(corners.sldNum)}>‹#›</div>
</PageFrame>

<style>
  /* 12 pt text in 1 pt dashed boxes, scaled with the page. */
  .placeholder { position: absolute; box-sizing: border-box; padding: calc(4px * var(--page-scale)) calc(7px * var(--page-scale)); border: 1px dashed #8c8c8c; font: calc(12px * var(--page-scale)) / 1.2 system-ui, sans-serif; overflow: hidden; }
  .end { text-align: right; }
  .bottom { display: flex; align-items: flex-end; }
  .bottom.end { justify-content: flex-end; }
  .frame { position: absolute; box-sizing: border-box; border: 1px dotted #8c8c8c; }
  .slide-image { position: absolute; box-sizing: border-box; border: 1px solid #000; overflow: hidden; }
  .body p { margin: 0 0 calc(2px * var(--page-scale)); }
</style>
