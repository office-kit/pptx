<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  import { imageCorrections } from '../core/image-corrections.ts';
  import { tick } from 'svelte';
  import {
    getShapeImageBrightness,
    getShapeImageContrast,
    getShapeKind,
    getShapeMedia,
    setShapeImageBrightness,
    setShapeImageContrast,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';

  // `target` picks videos (Video Format) or pictures (Picture Format); `big`
  // `row` and `icon` are the Picture Format tab's large button, 22 pt row
  // and 36 × 22 pt icon menu.
  let { variant = 'ribbon', target = 'video' } = $props<{ variant?: 'ribbon' | 'pane' | 'big' | 'row' | 'icon'; target?: 'video' | 'picture' }>();

  const componentId = $props.id();
  const editor = getEditor();
  const doc = editor.doc;
  const values = [-40, -20, 0, 20, 40] as const;
  // PowerPoint lays out contrast by row and brightness by column.
  const presets = values.flatMap((contrast) => values.map((brightness) => ({ brightness, contrast })));
  let open = $state(false);
  let trigger = $state<HTMLButtonElement>();
  let menu = $state<HTMLDivElement>();

  const selected = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    if (shapes.length !== 1) return null;
    const media = getShapeMedia(shapes[0]!);
    if (target === 'video' ? media?.kind !== 'video' : media || getShapeKind(shapes[0]!) !== 'picture') return null;
    return shapes[0]!;
  });
  const locked = $derived(editor.selectionLocked());

  function label(value: number): string {
    if (value === 0) return `0% (${t('Normal')})`;
    return `${value > 0 ? '+' : ''}${value}%`;
  }

  function presetLabel(preset: (typeof presets)[number]): string {
    return `${t('Brightness')}: ${label(preset.brightness)} ${t('Contrast')}: ${label(preset.contrast)}`;
  }

  function isSelected(preset: (typeof presets)[number]): boolean {
    if (!selected) return false;
    return (getShapeImageBrightness(selected) ?? 0) === preset.brightness / 100 && (getShapeImageContrast(selected) ?? 0) === preset.contrast / 100;
  }

  function close(restore = true): void {
    open = false;
    if (restore) trigger?.focus();
  }

  async function show(): Promise<void> {
    if (open) { close(); return; }
    open = true;
    await tick();
    menu?.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus();
  }

  function apply(preset: (typeof presets)[number]): void {
    const shape = selected;
    if (!shape || locked) return;
    doc.transact(t(target === 'video' ? 'Set video corrections' : 'Corrections'), () => {
      setShapeImageBrightness(shape, preset.brightness / 100 || null);
      setShapeImageContrast(shape, preset.contrast / 100 || null);
    });
    close();
  }

  function place(node: HTMLElement): void {
    const bounds = trigger!.getBoundingClientRect();
    node.style.left = `${Math.max(8, Math.min(bounds.left, innerWidth - node.offsetWidth - 8))}px`;
    node.style.top = `${Math.min(bounds.bottom, innerHeight - node.offsetHeight - 8)}px`;
  }

  function keys(event: KeyboardEvent): void {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    if (event.key === 'Tab') { close(false); return; }
    const items = [...menu!.querySelectorAll<HTMLButtonElement>('.grid button:not(:disabled)')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    if (event.key === 'Home' || event.key === 'End') {
      event.preventDefault();
      menu!.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')[event.key === 'Home' ? 0 : menu!.querySelectorAll<HTMLButtonElement>('button:not(:disabled)').length - 1]?.focus();
      return;
    }
    if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return;
    event.preventDefault();
    if (index < 0) return;
    const row = Math.floor(index / values.length);
    const column = index % values.length;
    const nextRow = event.key === 'ArrowUp' ? (row + values.length - 1) % values.length : event.key === 'ArrowDown' ? (row + 1) % values.length : row;
    const nextColumn = event.key === 'ArrowLeft' ? (column + values.length - 1) % values.length : event.key === 'ArrowRight' ? (column + 1) % values.length : column;
    items[nextRow * values.length + nextColumn]?.focus();
  }
</script>

<svelte:window onpointerdown={event => { if (open && !menu?.contains(eventTarget(event) as Node) && !trigger?.contains(eventTarget(event) as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />

{#if variant === 'big' || variant === 'row' || variant === 'icon'}
<button class="ctx-{variant}" type="button" bind:this={trigger} disabled={!selected || locked} aria-label={t('Corrections')} aria-haspopup="menu" aria-expanded={open} onclick={show}>
  {#if variant === 'big'}<span class="ctx-icon-row"><Icon name="corrections" size={32} /><span class="ctx-caret" aria-hidden="true">▾</span></span><span class="ctx-caption">{t('Corrections')}</span>{:else if variant === 'icon'}<Icon name="corrections" size={18} /><span class="ctx-caret" aria-hidden="true">▾</span>{:else}<Icon name="corrections" size={16} /><span>{t('Corrections')}</span><span class="ctx-caret" aria-hidden="true">▾</span>{/if}
</button>
{:else}
<button class="trigger" class:compact={variant === 'pane'} type="button" bind:this={trigger} disabled={!selected || locked} aria-label={t(variant === 'pane' ? 'Corrections presets' : 'Corrections')} aria-haspopup="menu" aria-expanded={open} onclick={show}>
  {#if variant !== 'pane'}<Icon name="gradient" />{/if}<span>{t(variant === 'pane' ? 'Presets' : 'Corrections')} ▾</span>
</button>
{/if}
{#if open}
  <div class="menu" role="menu" aria-label={t('Corrections')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    <div class="heading">{t('Brightness')} / {t('Contrast')}</div>
    <div class="grid">
      {#each presets as preset (preset.brightness + ':' + preset.contrast)}
        {@const filterId = `${componentId}-${preset.brightness}-${preset.contrast}`}
        {@const { slope, intercept } = imageCorrections(preset.brightness / 100, preset.contrast / 100)}
        <button
          type="button"
          role="menuitemradio"
          aria-checked={isSelected(preset)}
          aria-label={presetLabel(preset)}
          title={presetLabel(preset)}
          class:selected={isSelected(preset)}
          onclick={() => apply(preset)}
        >
          <span class="sample" aria-hidden="true">
            <svg viewBox="0 0 80 45" preserveAspectRatio="none">
              <defs><filter id={filterId}><feComponentTransfer color-interpolation-filters="sRGB">
                <feFuncR type="linear" {slope} {intercept} />
                <feFuncG type="linear" {slope} {intercept} />
                <feFuncB type="linear" {slope} {intercept} />
              </feComponentTransfer></filter></defs>
              <g filter={`url(#${filterId})`}><rect width="80" height="45" fill="#a9d9ef" /><circle cx="59" cy="12" r="7" fill="#f8d56a" /><path d="M0 36L20 17l14 12L50 9l30 27Z" fill="#557d55" /><path d="M0 41l16-10 11 6 14-9 20 11 19-6v12H0Z" fill="#2f543d" /><path d="M18 19l7 10-5-2-5 4-8-1Z" fill="#dbe8e9" opacity=".7" /></g></svg>
          </span>
        </button>
      {/each}
    </div>
    <button class="options" type="button" role="menuitem" onclick={() => { close(false); editor.showShapeFormat(target === 'video' ? 'video' : 'paint'); }}>{t(target === 'video' ? 'Movie Correction Options...' : 'Picture Corrections Options...')}</button>
  </div>
{/if}

<style>
  .trigger { display:flex; flex-direction:column; align-items:center; justify-content:center; gap:3px; min-width:68px; min-height:48px; padding:4px 7px; background:transparent; border:1px solid transparent; border-radius:var(--ok-radius); color:var(--ok-text); font:inherit; font-size:10px; cursor:pointer; }
  .trigger.compact { display: inline-flex; flex-direction: row; min-width: 0; min-height: 0; padding: 3px 8px; border: 1px solid var(--ok-border); font-size: 11px; }
  .trigger:hover:not(:disabled) { background:var(--ok-hover); border-color:var(--ok-border); }
  .trigger:disabled { opacity:.4; cursor:default; }
  .menu { position:fixed; z-index:400; width:270px; padding:8px; border:1px solid var(--ok-border); border-radius:6px; background:var(--ok-panel); color:var(--ok-text); box-shadow:var(--ok-shadow-lg); }
  .heading { padding:2px 4px 7px; color:var(--ok-text-2); font-size:11px; }
  .grid { display:grid; grid-template-columns:repeat(5, 1fr); gap:4px; }
  .grid button { min-width:0; padding:3px; border:1px solid transparent; border-radius:4px; background:transparent; color:inherit; font:inherit; font-size:9px; cursor:pointer; }
  .grid button:hover, .grid button:focus-visible, .grid button.selected { border-color:var(--ok-accent); background:var(--ok-hover); outline:none; }
  .sample { display:block; height:32px; overflow:hidden; border:1px solid var(--ok-border); border-radius:2px; }
  .sample svg { display:block; width:100%; height:100%; }
  .options { display:block; width:100%; margin-top:7px; padding:6px 4px 2px; border:0; border-top:1px solid var(--ok-border); background:transparent; color:inherit; font:inherit; font-size:11px; text-align:left; cursor:pointer; }
  .options:hover, .options:focus-visible { color:var(--ok-accent); outline:none; }
</style>
