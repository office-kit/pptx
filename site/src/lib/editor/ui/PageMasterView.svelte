<script lang="ts">
  // View ▸ Handout Master and Notes Master: the deck's master (or, until the
  // first edit writes one, the default Office master) as Mac PowerPoint draws
  // it — header, date, footer and page number in the corners; the slide
  // frames for the chosen slides per page on the handout, the slide image and
  // five body levels on the notes page. The ribbon edits them.
  import { getHandoutSlidesPerPage, getSlideLayouts, getSlideSize } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { handoutSlideFrames, pageMasterBoxes, pageSizePt, type Area } from '../core/master-geometry.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import MasterSlide from './MasterSlide.svelte';
  import PageFrame from './PageFrame.svelte';

  let { kind }: { kind: 'handoutMaster' | 'notesMaster' } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  const place = (area: Area) => `left:${area.x * 100}%;top:${area.y * 100}%;width:${area.w * 100}%;height:${area.h * 100}%`;
  const today = new Date().toLocaleDateString();
  const boxes = $derived.by(() => { doc.version; return pageMasterBoxes(doc.pres, kind); });
  const box = (role: string) => boxes.find((item) => item.role === role) ?? null;
  const page = $derived.by(() => { doc.version; return pageSizePt(doc.pres); });
  const frames = $derived.by(() => {
    doc.version;
    const size = getSlideSize(doc.pres) ?? { width: 16, height: 9 };
    return handoutSlideFrames(getHandoutSlidesPerPage(doc.pres), page, size.width / size.height);
  });
  const layout = $derived.by(() => { doc.version; return getSlideLayouts(doc.pres)[0] ?? null; });
  const levels = $derived([t('Click to edit Master text styles'), t('Second level'), t('Third level'), t('Fourth level'), t('Fifth level')]);
</script>

<PageFrame label={t(kind === 'handoutMaster' ? 'Handout Master' : 'Notes Master')}>
  {@const hdr = box('hdr')}
  {@const dt = box('dt')}
  {@const ftr = box('ftr')}
  {@const sldNum = box('sldNum')}
  {@const slideImage = box('sldImg')}
  {@const body = box('body')}
  {#if hdr}<div class="placeholder corner" data-role="hdr" style={place(hdr)}>{t('Header')}</div>{/if}
  {#if dt}<div class="placeholder corner end" data-role="dt" style={place(dt)}>{today}</div>{/if}
  {#if kind === 'handoutMaster'}
    {#each frames as frame, i (i)}<div class="frame" data-role="slide" style={place(frame)}></div>{/each}
  {:else}
    {#if slideImage}<div class="slide-image" data-role="slide" style={place(slideImage)}>{#if layout}<MasterSlide pres={doc.pres} {layout} master outlines={false} version={doc.version} pixelWidth={slideImage.w * page.width * editor.pageZoom} />{/if}</div>{/if}
    {#if body}<div class="placeholder body" data-role="body" style={place(body)}>{#each levels as text, level (level)}<p style:padding-left="calc({level} * 36px * var(--page-scale))">{text}</p>{/each}</div>{/if}
  {/if}
  {#if ftr}<div class="placeholder corner bottom" data-role="ftr" style={place(ftr)}>{t('Footer')}</div>{/if}
  {#if sldNum}<div class="placeholder corner end bottom" data-role="sldNum" style={place(sldNum)}>‹#›</div>{/if}
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
