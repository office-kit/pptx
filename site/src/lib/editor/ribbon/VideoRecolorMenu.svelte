<script lang="ts">
  import { imageCorrections } from '../core/image-corrections.ts';
  import { tick } from 'svelte';
  import {
    getPresentationTheme,
    getShapeImageBiLevelThreshold,
    getShapeImageBrightness,
    getShapeImageContrast,
    getShapeImageDuotone,
    getShapeKind,
    getShapeMedia,
    isShapeImageGrayscale,
    resolveDrawingColor,
    type Color,
    type ColorTransform,
    type ImageRecolor,
  } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import ColorPicker from '../ui/ColorPicker.svelte';

  // `target` picks videos (Video Format) or pictures (Picture Format); `big`
  // `row` and `icon` are the Picture Format tab's large button, 22 pt row
  // and 36 × 22 pt icon menu.
  let { variant = 'ribbon', target = 'video' } = $props<{ variant?: 'ribbon' | 'pane' | 'big' | 'row' | 'icon'; target?: 'video' | 'picture' }>();

  type RecolorPreset = {
    id: string;
    label: string;
    kind: 'none' | 'gray' | 'sepia' | 'washout' | 'bw' | 'theme';
    theme?: 'bg2' | 'tx2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6';
    shade?: 'dark' | 'light';
    threshold?: number;
  };

  const presets: RecolorPreset[] = [
    { id: 'none', label: 'No Recolor', kind: 'none' },
    { id: 'grayscale', label: 'Grayscale', kind: 'gray' },
    { id: 'sepia', label: 'Sepia', kind: 'sepia' },
    { id: 'washout', label: 'Washout', kind: 'washout' },
    { id: 'black-white-25', label: 'Black and White: 25%', kind: 'bw', threshold: 25 },
    { id: 'black-white-50', label: 'Black and White: 50%', kind: 'bw', threshold: 50 },
    { id: 'black-white-75', label: 'Black and White: 75%', kind: 'bw', threshold: 75 },
    { id: 'text2-dark', label: 'Text 2 Dark', kind: 'theme', theme: 'tx2', shade: 'dark' },
    { id: 'accent1-dark', label: 'Accent 1 Dark', kind: 'theme', theme: 'accent1', shade: 'dark' },
    { id: 'accent2-dark', label: 'Accent 2 Dark', kind: 'theme', theme: 'accent2', shade: 'dark' },
    { id: 'accent3-dark', label: 'Accent 3 Dark', kind: 'theme', theme: 'accent3', shade: 'dark' },
    { id: 'accent4-dark', label: 'Accent 4 Dark', kind: 'theme', theme: 'accent4', shade: 'dark' },
    { id: 'accent5-dark', label: 'Accent 5 Dark', kind: 'theme', theme: 'accent5', shade: 'dark' },
    { id: 'accent6-dark', label: 'Accent 6 Dark', kind: 'theme', theme: 'accent6', shade: 'dark' },
    { id: 'bg2-light', label: 'Background 2 Light', kind: 'theme', theme: 'bg2', shade: 'light' },
    { id: 'accent1-light', label: 'Accent 1 Light', kind: 'theme', theme: 'accent1', shade: 'light' },
    { id: 'accent2-light', label: 'Accent 2 Light', kind: 'theme', theme: 'accent2', shade: 'light' },
    { id: 'accent3-light', label: 'Accent 3 Light', kind: 'theme', theme: 'accent3', shade: 'light' },
    { id: 'accent4-light', label: 'Accent 4 Light', kind: 'theme', theme: 'accent4', shade: 'light' },
    { id: 'accent5-light', label: 'Accent 5 Light', kind: 'theme', theme: 'accent5', shade: 'light' },
    { id: 'accent6-light', label: 'Accent 6 Light', kind: 'theme', theme: 'accent6', shade: 'light' },
  ];

  const editor = getEditor();
  const doc = editor.doc;
  let open = $state(false);
  let trigger = $state<HTMLButtonElement>();
  let menu = $state<HTMLDivElement>();
  const componentId = $props.id();
  const theme = $derived.by(() => { doc.version; return getPresentationTheme(doc.pres); });
  const selected = $derived.by(() => {
    doc.version;
    const shapes = editor.selectedShapes();
    if (shapes.length !== 1) return null;
    const media = getShapeMedia(shapes[0]!);
    if (target === 'video' ? media?.kind !== 'video' : media || getShapeKind(shapes[0]!) !== 'picture') return null;
    return shapes[0]!;
  });
  const locked = $derived(editor.selectionLocked());

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

  function apply(preset: RecolorPreset): void {
    if (!selected || locked) return;
    // The command owns the single undo transaction. Recoloring also clears
    // any previous image color correction in the same operation.
    editor.invoke('setShapeImageRecolor', { recolor: recolorFor(preset) });
    close();
  }

  function chooseVariation(color: Color, transforms: readonly ColorTransform[] = []): void {
    if (!selected || locked) return;
    editor.invoke('setShapeImageRecolor', {
      recolor: {
        kind: 'duotone',
        colors: ['#000000', { color, colorTransforms: [...transforms, { kind: 'tint', value: 0.45 }, { kind: 'satMod', value: 4 }] }],
      },
    });
    close(false);
  }

  function recolorFor(preset: RecolorPreset): ImageRecolor {
    if (preset.kind === 'none') return { kind: 'none' };
    if (preset.kind === 'gray') return { kind: 'grayscale' };
    if (preset.kind === 'washout') return { kind: 'washout' };
    if (preset.kind === 'bw') return { kind: 'threshold', threshold: preset.threshold ?? 50 };
    if (preset.kind === 'sepia') {
      return {
        kind: 'duotone',
        colors: [
          '#000000',
          { color: '#D9C3A5', colorTransforms: [{ kind: 'tint', value: 0.5 }, { kind: 'satMod', value: 1.8 }] },
        ],
      };
    }
    const color = preset.theme ?? 'accent1';
    if (preset.shade === 'light') {
      return {
        kind: 'duotone',
        colors: [
          { color, colorTransforms: [{ kind: 'shade', value: 0.45 }, { kind: 'satMod', value: 1.35 }] },
          '#FFFFFF',
        ],
      };
    }
    return {
      kind: 'duotone',
      colors: ['#000000', { color, colorTransforms: [{ kind: 'tint', value: 0.45 }, { kind: 'satMod', value: 4 }] }],
    };
  }

  function sameColor(left: string | null, right: string | null | undefined): boolean {
    return !!left && !!right && left.replace('#', '').toUpperCase() === right.replace('#', '').toUpperCase();
  }

  function transformedColor(base: string, transforms: readonly [number, number] | null): string | null {
    if (!transforms) return base;
    type ColorElement = Parameters<typeof resolveDrawingColor>[0];
    const [tint, satMod] = transforms;
    const element: ColorElement = {
      kind: 'element' as const,
      name: { prefix: 'a', localName: 'srgbClr', namespaceURI: 'http://schemas.openxmlformats.org/drawingml/2006/main' },
      attrs: [
        { name: { prefix: '', localName: 'val', namespaceURI: '' }, value: base.replace(/^#/, '') },
      ],
      prefixDecls: new Map<string, string>(),
      children: [
        {
          kind: 'element' as const,
          name: { prefix: 'a', localName: tint < 0 ? 'shade' : 'tint', namespaceURI: 'http://schemas.openxmlformats.org/drawingml/2006/main' },
          attrs: [{ name: { prefix: '', localName: 'val', namespaceURI: '' }, value: String(Math.abs(tint) * 100000) }],
          prefixDecls: new Map<string, string>(), children: [],
        },
        {
          kind: 'element' as const,
          name: { prefix: 'a', localName: 'satMod', namespaceURI: 'http://schemas.openxmlformats.org/drawingml/2006/main' },
          attrs: [{ name: { prefix: '', localName: 'val', namespaceURI: '' }, value: String(satMod * 100000) }],
          prefixDecls: new Map<string, string>(), children: [],
        },
      ],
    };
    return resolveDrawingColor(element, null);
  }

  function isSelected(preset: RecolorPreset): boolean {
    if (!selected) return false;
    if (preset.kind === 'gray') return isShapeImageGrayscale(selected);
    if (preset.kind === 'bw') return getShapeImageBiLevelThreshold(selected) === preset.threshold;
    if (preset.kind === 'washout') {
      return getShapeImageBrightness(selected) === 0.7 && getShapeImageContrast(selected) === -0.7;
    }
    const duotone = getShapeImageDuotone(doc.pres, selected);
    if (preset.kind === 'none') {
      return !isShapeImageGrayscale(selected) && getShapeImageBiLevelThreshold(selected) === null && !duotone &&
        getShapeImageBrightness(selected) === null && getShapeImageContrast(selected) === null;
    }
    if (!duotone) return false;
    if (preset.kind === 'sepia') return sameColor(duotone.firstColor, '#000000') && sameColor(duotone.secondColor, transformedColor('#D9C3A5', [0.5, 1.8]) ?? undefined);
    const base = themePaint(preset);
    const color = base ? transformedColor(base, preset.shade === 'light' ? [-0.45, 1.35] : [0.45, 4]) : undefined;
    return preset.shade === 'dark'
      ? sameColor(duotone.firstColor, '#000000') && sameColor(duotone.secondColor, color)
      : sameColor(duotone.firstColor, color) && sameColor(duotone.secondColor, '#FFFFFF');
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
      event.preventDefault(); items[event.key === 'Home' ? 0 : items.length - 1]?.focus(); return;
    }
    if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key) || index < 0) return;
    event.preventDefault();
    const row = Math.floor(index / 7);
    const column = index % 7;
    const nextRow = event.key === 'ArrowUp' ? (row + 2) % 3 : event.key === 'ArrowDown' ? (row + 1) % 3 : row;
    const nextColumn = event.key === 'ArrowLeft' ? (column + 6) % 7 : event.key === 'ArrowRight' ? (column + 1) % 7 : column;
    items[nextRow * 7 + nextColumn]?.focus();
  }

  function themePaint(preset: RecolorPreset): string | undefined {
    if (preset.kind !== 'theme' || !preset.theme || !theme) return undefined;
    const slot = preset.theme === 'tx2' ? 'dark2' : preset.theme === 'bg2' ? 'light2' : preset.theme;
    return theme[slot];
  }

  const washout = imageCorrections(0.7, -0.7);
  const luminanceMatrix = '0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0';

  function duotoneColors(preset: RecolorPreset): readonly [string, string] | null {
    if (preset.kind === 'sepia') {
      return ['#000000', transformedColor('#D9C3A5', [0.5, 1.8]) ?? '#D9C3A5'];
    }
    if (preset.kind !== 'theme') return null;
    const base = themePaint(preset);
    if (!base) return null;
    const color = transformedColor(base, preset.shade === 'light' ? [-0.45, 1.35] : [0.45, 4]);
    if (!color) return null;
    return preset.shade === 'light' ? [color, '#FFFFFF'] : ['#000000', color];
  }

  function channelTable(first: string, second: string, channel: number): string {
    const value = (color: string) => {
      const parsed = Number.parseInt(color.slice(1 + channel * 2, 3 + channel * 2), 16);
      return Number.isFinite(parsed) ? parsed / 255 : 0;
    };
    return `${value(first)} ${value(second)}`;
  }

  function hasFilter(preset: RecolorPreset): boolean {
    return preset.kind === 'washout' || preset.kind === 'gray' || preset.kind === 'sepia' || preset.kind === 'bw' || preset.kind === 'theme';
  }
</script>

<svelte:window onpointerdown={event => { if (open && !menu?.contains(event.target as Node) && !trigger?.contains(event.target as Node)) close(false); }} onblur={() => { if (open) close(false); }} onresize={() => { if (open) close(false); }} />

{#if variant === 'big' || variant === 'row' || variant === 'icon'}
<button class="ctx-{variant}" type="button" bind:this={trigger} disabled={!selected || locked} aria-label={t('Color')} aria-haspopup="menu" aria-expanded={open} onclick={show}>
  {#if variant === 'big'}<span class="ctx-icon-row"><Icon name="gradient" size={32} /><span class="ctx-caret" aria-hidden="true">▾</span></span><span class="ctx-caption">{t('Color')}</span>{:else if variant === 'icon'}<Icon name="gradient" size={18} /><span class="ctx-caret" aria-hidden="true">▾</span>{:else}<Icon name="gradient" size={16} /><span>{t('Color')}</span><span class="ctx-caret" aria-hidden="true">▾</span>{/if}
</button>
{:else}
<button class="trigger" class:compact={variant === 'pane'} type="button" bind:this={trigger} disabled={!selected || locked} aria-label={t(variant === 'pane' ? 'Recolor presets' : 'Color')} aria-haspopup="menu" aria-expanded={open} onclick={show}>
  {#if variant !== 'pane'}<Icon name="gradient" />{/if}<span>{t(variant === 'pane' ? 'Presets' : 'Color')} ▾</span>
</button>
{/if}
{#if open}
  <div class="menu" role="menu" aria-label={t('Recolor')} tabindex="-1" bind:this={menu} use:place onkeydown={keys}>
    <div class="heading">{t('Recolor')}</div>
    <div class="grid">
      {#each presets as preset (preset.id)}
        {@const filterId = `${componentId}-${preset.id}`}
        {@const colors = duotoneColors(preset)}
        <button type="button" role="menuitemradio" aria-checked={isSelected(preset)} aria-label={t(preset.label)} title={t(preset.label)} onclick={() => apply(preset)}>
          <span class="sample" class:none={preset.kind === 'none'}>
            <svg viewBox="0 0 80 45" preserveAspectRatio="none" aria-hidden="true">
              {#if hasFilter(preset)}
                <defs><filter id={filterId}>
                  {#if preset.kind === 'washout'}
                    <feComponentTransfer color-interpolation-filters="sRGB">
                      <feFuncR type="linear" slope={washout.slope} intercept={washout.intercept} />
                      <feFuncG type="linear" slope={washout.slope} intercept={washout.intercept} />
                      <feFuncB type="linear" slope={washout.slope} intercept={washout.intercept} />
                    </feComponentTransfer>
                  {:else if preset.kind === 'gray'}
                    <feColorMatrix type="matrix" values={luminanceMatrix} />
                  {:else if colors}
                    <feColorMatrix type="matrix" values={luminanceMatrix} />
                    <feComponentTransfer>
                      <feFuncR type="table" tableValues={channelTable(colors[0], colors[1], 0)} />
                      <feFuncG type="table" tableValues={channelTable(colors[0], colors[1], 1)} />
                      <feFuncB type="table" tableValues={channelTable(colors[0], colors[1], 2)} />
                    </feComponentTransfer>
                  {:else if preset.kind === 'bw'}
                    <feColorMatrix type="matrix" values={luminanceMatrix} />
                    <feComponentTransfer>
                      <feFuncR type="linear" slope="1" intercept={0.5 - (preset.threshold ?? 50) / 100} />
                      <feFuncG type="linear" slope="1" intercept={0.5 - (preset.threshold ?? 50) / 100} />
                      <feFuncB type="linear" slope="1" intercept={0.5 - (preset.threshold ?? 50) / 100} />
                    </feComponentTransfer>
                    <feComponentTransfer>
                      <feFuncR type="discrete" tableValues="0 1" />
                      <feFuncG type="discrete" tableValues="0 1" />
                      <feFuncB type="discrete" tableValues="0 1" />
                    </feComponentTransfer>
                  {/if}
                </filter></defs>
              {/if}
              <g filter={hasFilter(preset) ? `url(#${filterId})` : undefined}>
                <rect width="80" height="45" fill="#a9d9ef" /><circle cx="59" cy="12" r="7" fill="#f8d56a" /><path d="M0 36L20 17l14 12L50 9l30 27Z" fill="#557d55" /><path d="M0 41l16-10 11 6 14-9 20 11 19-6v12H0Z" fill="#2f543d" /><path d="M18 19l7 10-5-2-5 4-8-1Z" fill="#dbe8e9" opacity=".7" />
              </g>
            </svg>
          </span>
        </button>
      {/each}
    </div>
    <div class="variation" role="none">
      <ColorPicker label={t('More Variations...')} showThemeShades disabled={!selected || locked} choose={chooseVariation} />
      <span>{t('More Variations...')}</span>
    </div>
    <button class="options" type="button" role="menuitem" onclick={() => { close(false); editor.showShapeFormat(target === 'video' ? 'video' : 'paint'); }}>{t(target === 'video' ? 'Movie Color Options...' : 'Picture Color Options...')}</button>
  </div>
{/if}

<style>
  .trigger { display:flex; flex-direction:column; align-items:center; justify-content:center; gap:3px; min-width:68px; min-height:48px; padding:4px 7px; background:transparent; border:1px solid transparent; border-radius:var(--ok-radius); color:var(--ok-text); font:inherit; font-size:10px; cursor:pointer; }
  .trigger.compact { display:inline-flex; flex-direction:row; min-width:0; min-height:0; padding:3px 8px; border:1px solid var(--ok-border); font-size:11px; }
  .trigger:hover:not(:disabled) { background:var(--ok-hover); border-color:var(--ok-border); }
  .trigger:disabled { opacity:.4; cursor:default; }
  .menu { position:fixed; z-index:400; width:300px; padding:8px; border:1px solid var(--ok-border); border-radius:6px; background:var(--ok-panel); color:var(--ok-text); box-shadow:var(--ok-shadow-lg); }
  .heading { padding:2px 4px 7px; color:var(--ok-text-2); font-size:11px; }
  .grid { display:grid; grid-template-columns:repeat(7, 1fr); gap:4px; }
  .grid button { min-width:0; padding:3px; border:1px solid transparent; border-radius:4px; background:transparent; color:inherit; font:inherit; cursor:pointer; }
  .grid button:hover, .grid button:focus-visible { border-color:var(--ok-accent); background:var(--ok-hover); outline:none; }
  .sample { position:relative; display:block; height:32px; overflow:hidden; border:1px solid var(--ok-border); border-radius:2px; background:var(--ok-panel); }
  .sample svg { display:block; width:100%; height:100%; }
  .sample.none::after { content:'×'; position:absolute; inset:0; display:grid; place-items:center; color:#b3261e; font-size:26px; font-weight:600; background:#fff8; }
  .options { display:block; width:100%; margin-top:7px; padding:6px 4px 2px; border:0; border-top:1px solid var(--ok-border); background:transparent; color:inherit; font:inherit; font-size:11px; text-align:left; cursor:pointer; }
  .variation { display:flex; align-items:center; gap:7px; margin-top:7px; padding:6px 4px 2px; border-top:1px solid var(--ok-border); font-size:11px; }
  .variation :global(.trigger) { flex:none; }
  .options:hover, .options:focus-visible { color:var(--ok-accent); outline:none; }
</style>
