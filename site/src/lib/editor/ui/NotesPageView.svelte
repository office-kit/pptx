<script lang="ts">
  // View ▸ Notes Page: the printed notes page — the slide image above its
  // speaker notes, where the deck's notes master (or the default one) puts
  // them — with the notes editable in place.
  import { getEditor } from '../core/context.ts';
  import { NOTES_AREAS, pageMasterBoxes, type Area } from '../core/master-geometry.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import NotesPane from './NotesPane.svelte';
  import PageFrame from './PageFrame.svelte';

  const editor = getEditor();
  const doc = editor.doc;
  const boxes = $derived.by(() => { doc.version; return pageMasterBoxes(doc.pres, 'notesMaster'); });
  const slideImage = $derived(boxes.find((box) => box.role === 'sldImg') ?? NOTES_AREAS.sldImg);
  const body = $derived(boxes.find((box) => box.role === 'body') ?? NOTES_AREAS.body);
  const place = (area: Area) => `left:${area.x * 100}%;top:${area.y * 100}%;width:${area.w * 100}%;height:${area.h * 100}%`;
</script>

<PageFrame label={t('Notes Page')}>
  <button class="slide" style={place(slideImage)} aria-label={t('Edit slide')} ondblclick={() => editor.setViewMode('normal')}>{@html doc.currentSvg}</button>
  {#if doc.currentSlide}{#key doc.currentSlide}<div class="notes" style={place(body)}><NotesPane /></div>{/key}{/if}
</PageFrame>

<style>
  .slide { position: absolute; box-sizing: border-box; padding: 0; border: 1px solid #000; background: #fff; cursor: default; overflow: hidden; }
  .slide :global(svg) { display: block; width: 100%; height: 100%; }
  .notes { position: absolute; display: flex; min-height: 0; }
  .notes :global(> *) { flex: 1; }
</style>
