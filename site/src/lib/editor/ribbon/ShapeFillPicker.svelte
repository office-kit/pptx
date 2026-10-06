<script lang="ts">
  // The ribbon's Shape Fill ▾ (Home ▸ Drawing and Shape Format): colors plus
  // the Texture ▸ gallery, whose More Textures... chooses a picture file.
  import type { Color } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { canFillWithPicture, fillSelectionWithPicture } from '../core/picture-fill.ts';
  import { texturePng, type TextureId } from '../core/textures.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import ColorPicker from '../ui/ColorPicker.svelte';

  let { disabled }: { disabled: boolean } = $props();
  const editor = getEditor();
  const doc = editor.doc;
  let input = $state<HTMLInputElement>();
  const pictureFillable = $derived.by(() => { doc.version; doc.selection; return canFillWithPicture(editor); });

  function fill(color: Color) { editor.invoke('setShapeFill', { color: { color } }); }
  async function picture(read: () => Promise<Uint8Array>, kind: 'picture' | 'texture') {
    try { await fillSelectionWithPicture(editor, read, kind); }
    catch (cause) { editor.toast('error', cause instanceof Error ? cause.message : String(cause)); }
  }
  async function upload(event: Event) {
    const element = event.currentTarget;
    if (!(element instanceof HTMLInputElement)) return;
    const file = element.files?.[0];
    if (!file) return;
    try { await picture(async () => new Uint8Array(await file.arrayBuffer()), 'picture'); }
    finally { element.value = ''; }
  }
</script>

<ColorPicker compact label={t('Shape Fill')} {disabled} choose={fill} texture={{
  disabled: !pictureFillable,
  choose: (id: TextureId) => picture(() => texturePng(id), 'texture'),
  more: () => input?.click(),
}} />
<input type="file" accept="image/*" hidden bind:this={input} onchange={upload} aria-label={t('More Textures...')} />
