<script lang="ts">
  import { tick } from 'svelte';
  import { cm, emu, getShapeBoundsResolved, getShapeId, getShapeMedia, setShapeBounds } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { getMediaPreview } from '../core/media-preview.svelte.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import ArrangeMenu from './ArrangeMenu.svelte';
  import PosterFrameMenu from './PosterFrameMenu.svelte';

  const editor = getEditor();
  const doc = editor.doc;
  const preview = getMediaPreview(editor);
  let effectsOpen = $state(false);
  let effectsTrigger = $state<HTMLButtonElement>();
  let effectsMenu = $state<HTMLDivElement>();
  let sizeLock = $state(true);
  const maxDimension = 5963.92;

  const selected = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    if (shapes.length !== 1) return null;
    const shape = shapes[0]!;
    return getShapeMedia(shape)?.kind === 'video' ? shape : null;
  });
  const playLabel = $derived(preview.state.playing ? 'Pause' : 'Play');
  const geometry = $derived.by(() => {
    doc.version;
    return editor.selectedShapes().flatMap(shape => {
      const bounds = getShapeBoundsResolved(doc.pres, shape);
      return bounds ? [{ shape, bounds }] : [];
    });
  });
  const dimensions = $derived.by(() => {
    const common = (field: 'w' | 'h') => {
      const values = geometry.map(item => item.bounds[field]);
      return values.length > 0 && values.every(value => value === values[0]) ? Math.round((values[0]! / cm(1)) * 100) / 100 : null;
    };
    return { w: common('w'), h: common('h') };
  });
  const canLockSize = $derived(geometry.length > 0 && geometry.every(item => item.bounds.w > 0 && item.bounds.h > 0));

  function togglePlayback(): void {
    const shape = selected;
    if (!shape) return;
    preview.command(preview.state.playing ? 'pause' : 'play', getShapeId(shape));
  }

  function showFormatPane(tab: 'paint' | 'effects' | 'size' = 'paint'): void {
    editor.showShapeFormat(tab);
  }

  function changeDimension(field: 'w' | 'h', input: HTMLInputElement): void {
    const restore = () => { input.value = dimensions[field] === null ? '' : String(dimensions[field]); };
    if (!geometry.length || editor.selectionLocked() || !input.reportValidity() || !Number.isFinite(input.valueAsNumber)) { restore(); return; }
    const value = input.valueAsNumber;
    const updates = geometry.map(item => {
      const next = { ...item.bounds, [field]: emu(cm(value)) };
      if (sizeLock && canLockSize) {
        if (field === 'w') next.h = emu(item.bounds.h * next.w / item.bounds.w);
        else next.w = emu(item.bounds.w * next.h / item.bounds.h);
      }
      return { shape: item.shape, bounds: next };
    });
    if (updates.some(item => item.bounds.w > cm(maxDimension) || item.bounds.h > cm(maxDimension))) { restore(); editor.toast('error', t('The proportional size is too large')); return; }
    doc.transact(t('Set bounds'), () => { for (const item of updates) setShapeBounds(item.shape, item.bounds); });
  }

  function closeEffects(restore = true): void {
    effectsOpen = false;
    if (restore) effectsTrigger?.focus();
  }

  async function showEffects(): Promise<void> {
    effectsOpen = !effectsOpen;
    if (effectsOpen) {
      await tick();
      effectsMenu?.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus();
    }
  }

  function placeEffects(node: HTMLElement): void {
    const bounds = effectsTrigger!.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.min(bounds.bottom, innerHeight - node.offsetHeight - 8)}px`;
  }

  function effectsKeys(event: KeyboardEvent): void {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); closeEffects(); return; }
    if (event.key === 'Tab') { closeEffects(false); return; }
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const items = [...effectsMenu!.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length]?.focus();
  }
</script>

{#if selected}
  <div class="group">
    <div class="items">
      <button class="cmd" type="button" aria-label={t(playLabel)} onclick={togglePlayback}>
        <span class="play-icon" aria-hidden="true">{preview.state.playing ? 'Ⅱ' : '▶'}</span>
        <span>{t(playLabel)}</span>
      </button>
    </div>
    <span class="title">{t('Preview')}</span>
  </div>

  <div class="group">
    <div class="items"><PosterFrameMenu /></div>
    <span class="title">{t('Adjust')}</span>
  </div>

  <div class="group">
    <div class="items">
      <button class="cmd" type="button" disabled={!editor.canRun('setShapeStroke')} aria-label={t('Video Border')} onclick={() => editor.runOrPrompt('setShapeStroke')}>
        <Icon name="outline" /><span>{t('Video Border')}</span>
      </button>
      <div class="effects">
        <button class="cmd" type="button" bind:this={effectsTrigger} aria-haspopup="menu" aria-expanded={effectsOpen} aria-label={t('Video Effects')} onclick={showEffects}>
          <Icon name="shadow" /><span>{t('Video Effects')} ▾</span>
        </button>
        {#if effectsOpen}
          <div class="menu" role="menu" aria-label={t('Video Effects')} tabindex="-1" bind:this={effectsMenu} use:placeEffects onkeydown={effectsKeys}>
            <button role="menuitem" disabled={!editor.canRun('setShapeShadow')} onclick={() => { closeEffects(); editor.runOrPrompt('setShapeShadow'); }}>{t('Shadow')}</button>
            <button role="menuitem" disabled={!editor.canRun('setShapeGlow')} onclick={() => { closeEffects(); editor.runOrPrompt('setShapeGlow'); }}>{t('Glow')}</button>
            <button role="menuitem" disabled={!editor.canRun('clearShapeEffects')} onclick={() => { closeEffects(); editor.invoke('clearShapeEffects'); }}>{t('Clear Effects')}</button>
          </div>
        {/if}
      </div>
    </div>
    <span class="title">{t('Video Styles')}</span>
  </div>

  <div class="group">
    <div class="items">
      <button class="cmd" type="button" disabled={!editor.canRun('setShapeDescription')} aria-label={t('Alt Text')} onclick={() => editor.runOrPrompt('setShapeDescription')}>
        <Icon name="text-format" /><span>{t('Alt Text')}</span>
      </button>
    </div>
    <span class="title">{t('Alt Text')}</span>
  </div>

  <div class="group">
    <div class="items"><ArrangeMenu /></div>
    <span class="title">{t('Arrange')}</span>
  </div>

  <div class="group size-group">
    <div class="size-items">
      <label><span>{t('Height')}</span><input class="ok-input" type="number" min="0" max={maxDimension} step="any" disabled={editor.selectionLocked()} value={dimensions.h ?? ''} placeholder={dimensions.h === null ? t('Mixed') : undefined} onchange={event => changeDimension('h', event.currentTarget)} /><span>cm</span></label>
      <label><span>{t('Width')}</span><input class="ok-input" type="number" min="0" max={maxDimension} step="any" disabled={editor.selectionLocked()} value={dimensions.w ?? ''} placeholder={dimensions.w === null ? t('Mixed') : undefined} onchange={event => changeDimension('w', event.currentTarget)} /><span>cm</span></label>
      <label class="check"><input type="checkbox" bind:checked={sizeLock} disabled={editor.selectionLocked() || !canLockSize} /><span>{t('Lock aspect ratio')}</span></label>
    </div>
    <span class="title">{t('Size')}</span>
  </div>

  <div class="group">
    <div class="items">
      <button class="cmd" type="button" aria-label={t('Format Pane')} onclick={() => showFormatPane()}>
        <Icon name="resize" /><span>{t('Format Pane')}</span>
      </button>
    </div>
    <span class="title">{t('Format Pane')}</span>
  </div>
{/if}

<svelte:window
  onpointerdown={(event) => { if (effectsOpen && !(event.target as Element)?.closest('.effects')) closeEffects(false); }}
  onblur={() => { if (effectsOpen) closeEffects(false); }}
  onresize={() => { if (effectsOpen) closeEffects(false); }}
/>

<style>
  .group { flex: none; display: flex; flex-direction: column; justify-content: space-between; padding: 0 8px; border-right: 1px solid var(--ok-border); }
  .items { display: flex; align-items: center; gap: 2px; flex: 1; }
  .title { text-align: center; font-size: 10px; color: var(--ok-text-3); padding-top: 2px; }
  .cmd { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 3px; min-width: 62px; min-height: 48px; padding: 4px 5px; border: 1px solid transparent; border-radius: var(--ok-radius); background: transparent; color: var(--ok-text); font: inherit; font-size: 10px; cursor: pointer; }
  .cmd:hover:not(:disabled) { background: var(--ok-hover); border-color: var(--ok-border); }
  .cmd:disabled { opacity: .4; cursor: default; }
  .play-icon { display: block; height: 20px; line-height: 20px; font-size: 18px; }
  .effects { position: relative; }
  .menu { position: fixed; z-index: 400; min-width: 150px; padding: 5px; border: 1px solid var(--ok-border); border-radius: 6px; background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .menu button { display: block; width: 100%; padding: 6px 10px; border: 0; border-radius: 4px; background: transparent; color: inherit; font: inherit; font-size: 12px; text-align: left; white-space: nowrap; }
  .menu button:hover:not(:disabled), .menu button:focus-visible { background: var(--ok-accent); color: white; }
  .menu button:disabled { opacity: .4; }
  .size-group { min-width: 170px; }
  .size-items { display: grid; grid-template-columns: auto 76px auto; align-items: center; gap: 3px 5px; padding: 3px 0; font-size: 10px; }
  .size-items label { display: contents; }
  .size-items .ok-input { width: 70px; min-width: 0; padding: 2px 3px; font-size: 10px; }
  .size-items .check { grid-column: 1 / -1; display: flex; gap: 4px; align-items: center; }
</style>
