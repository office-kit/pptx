<script lang="ts">
  import { getShapePatternFill, setShapePatternFill, type PatternFillOptions } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { patterns, patternSwatches } from './patterns.ts';

  const editor = getEditor();
  const swatches = patternSwatches();
  const shapes = $derived.by(() => { editor.doc.version; return editor.selectedShapes(); });
  const fills = $derived(shapes.map(shape => getShapePatternFill(editor.doc.pres, shape)));
  const locked = $derived(editor.selectionLocked());
  function common(field: keyof PatternFillOptions): string | undefined {
    const values = new Set(fills.map(fill => fill?.[field]));
    return values.size === 1 ? [...values][0] : undefined;
  }
  function apply(patch: Partial<PatternFillOptions>) {
    if (locked) return;
    editor.doc.transact(t('Pattern fill'), () => {
      shapes.forEach(shape => setShapePatternFill(shape, patch));
    });
  }
  function keys(event: KeyboardEvent) {
    const offsets: Record<string, number> = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -6, ArrowDown: 6 };
    if (!(event.key in offsets) && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault(); event.stopPropagation();
    const items = [...(event.currentTarget as HTMLElement).parentElement!.querySelectorAll<HTMLButtonElement>('button')];
    const index = items.indexOf(event.target as HTMLButtonElement);
    items[event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + offsets[event.key]! + items.length) % items.length]?.focus();
  }
</script>

<fieldset disabled={locked} class="pattern-fields">
  <legend>{t('Pattern')}</legend>
  <div class="patterns" role="group" aria-label={t('Pattern')}>
    {#each patterns as [preset, label]}
      <button type="button" onkeydown={keys} aria-label={t(label)} title={t(label)} aria-pressed={common('preset') === preset} onclick={() => apply({ preset })}><img src={swatches.get(preset)} alt="" /></button>
    {/each}
  </div>
  {#each [['foreground', 'Foreground'], ['background', 'Background']] as [field, label]}
    {@const color = common(field === 'foreground' ? 'foreground' : 'background')}
    <label class="color"><span>{t(label)}</span><span>
      <input type="color" aria-label={t(label)} value={color ?? '#000000'} onchange={event => apply({ [field]: event.currentTarget.value })} />
      {#if color === undefined}<span>{t('Mixed')}</span>{/if}
    </span></label>
  {/each}
</fieldset>
<style>
  .pattern-fields { border: 0; padding: 0; margin: 0; min-width: 0; }
  legend { padding: 0; margin-bottom: 6px; }
  .patterns { display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 3px; margin-bottom: 12px; }
  .patterns button { padding: 2px; border: 1px solid var(--ok-border); border-radius: 2px; background: var(--ok-panel); min-width: 0; }
  .patterns button:hover, .patterns button:focus-visible, .patterns button[aria-pressed=true] { border-color: var(--ok-accent); outline: 1px solid var(--ok-accent); }
  img { display: block; width: 100%; aspect-ratio: 1.5; }
  .color { display: flex; align-items: center; justify-content: space-between; margin-top: 8px; }
  .color > span:last-child { display: flex; align-items: center; gap: 4px; }
</style>
