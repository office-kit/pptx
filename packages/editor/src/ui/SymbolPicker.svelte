<script lang="ts">
  import { eventTarget } from '../core/dom-root.ts';
  // The reference desktop app's (Mac) Symbol opens the system Character Viewer; the browser has
  // no such panel, so this is a small one with the characters decks use most.
  // Buttons never take focus, so the text cursor stays where it was.
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const GROUPS: ReadonlyArray<readonly [string, string]> = [
    ['Punctuation', '•·–—…‘’“”«»¶§†‡'],
    ['Currency', '$€£¥¢₩₹₽'],
    ['Math', '±×÷≈≠≤≥∞√∑∏∫∂∆πµ°‰'],
    ['Arrows', '←↑→↓↔↕⇐⇒⇔↗↘'],
    ['Marks', '©®™✓✗★☆♥♦♣♠●○■□▲△'],
    ['Greek', 'αβγδεθλσφωΩΣΦ'],
  ];
  const anchor = $derived(editor.symbolPicker);

  function insert(symbol: string) {
    editor.inlineTextFormat?.insertText?.(symbol);
  }
  $effect(() => {
    // The picker inserts at a text cursor; leaving text editing closes it.
    if (!editor.inlineTextFormat?.insertText) editor.symbolPicker = null;
  });
</script>

<svelte:window
  onkeydown={(event) => { if (event.key === 'Escape') editor.symbolPicker = null; }}
  onpointerdown={(event) => { if (!(eventTarget(event) as Element).closest?.('.symbol-picker, [aria-label="Symbol"]')) editor.symbolPicker = null; }}
/>

{#if anchor}
  <div class="symbol-picker" role="dialog" aria-label={t('Symbol')} style="left:{Math.max(8, Math.min(anchor.left, innerWidth - 336))}px; top:{anchor.bottom + 4}px">
    {#each GROUPS as [title, symbols] (title)}
      <div class="group" role="group" aria-label={t(title)}>
        <span class="title">{t(title)}</span>
        <div class="grid">
          {#each [...symbols] as symbol (symbol)}
            <button type="button" tabindex="-1" aria-label={symbol} onmousedown={(event) => event.preventDefault()} onclick={() => insert(symbol)}>{symbol}</button>
          {/each}
        </div>
      </div>
    {/each}
  </div>
{/if}

<style>
  .symbol-picker { position: fixed; z-index: 400; width: 320px; max-height: 60vh; overflow-y: auto; padding: 8px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  .group { margin-bottom: 6px; }
  .title { display: block; margin: 2px 2px 4px; font-size: 11px; color: var(--ok-text-2); }
  .grid { display: grid; grid-template-columns: repeat(10, 1fr); gap: 2px; }
  button { height: 28px; font-size: 16px; color: inherit; background: transparent; border: 1px solid transparent; border-radius: 4px; cursor: pointer; }
  button:hover { background: var(--ok-hover); border-color: var(--ok-border); }
</style>
