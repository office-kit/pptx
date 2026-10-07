<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { replaceTextInPresentation, replaceTextInSlide, setShapeText, setTableCellText } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { findText, searchPattern, type TextMatch } from '../core/find-text.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const initialSlide = untrack(() => doc.selection.slideIndex);
  let query = $state('');
  let replacement = $state('');
  let matchCase = $state(false);
  let currentSlideOnly = $state(false);
  let cursor = $state(0);
  let error = $state('');
  let notice = $state('');
  let dialog: HTMLDialogElement;
  let searchInput: HTMLInputElement;
  const matches = $derived.by(() => {
    doc.version;
    return findText(doc.pres, query, matchCase, currentSlideOnly ? initialSlide : undefined);
  });
  const current = $derived(matches[Math.min(cursor, Math.max(0, matches.length - 1))]);
  onMount(() => { dialog.showModal(); searchInput.focus(); });
  function reveal(match: TextMatch | undefined) {
    if (!match) return;
    if (match.target.kind === 'cell') doc.selectCell(match.slideIndex, match.shapeId, match.target.row, match.target.col);
    else doc.selectShape(match.slideIndex, match.shapeId);
  }
  function reset() { cursor = 0; notice = ''; error = ''; reveal(matches[0]); }
  $effect(() => { query; matchCase; currentSlideOnly; untrack(reset); });
  function move(step: number) {
    if (!matches.length) return;
    cursor = (cursor + step + matches.length) % matches.length;
    reveal(current);
  }
  function replace(all: boolean) {
    if (!current) return;
    error = ''; notice = '';
    try {
      const changed = all ? matches.filter(m => m.text.slice(m.start, m.end) !== replacement).length : Number(current.text.slice(current.start, current.end) !== replacement);
      if (!changed) { notice = t('No text changed'); if (!all) move(1); return; }
      const match = current;
      doc.transact(t(all ? 'Replace all' : 'Replace match'), () => {
        if (all) {
          const pattern = searchPattern(query, matchCase);
          // The dialog accepts literal text; escape native replacement-string tokens.
          const value = replacement.replace(/\$/g, () => '$$');
          if (currentSlideOnly) {
            const slide = doc.slides[initialSlide];
            if (slide) replaceTextInSlide(slide, pattern, value);
          } else replaceTextInPresentation(doc.pres, pattern, value);
        } else {
          const options = { range: { start: match.start, end: match.end } };
          if (match.target.kind === 'cell') setTableCellText(match.target.cell, replacement, options);
          else setShapeText(match.target.shape, replacement, options);
        }
      });
      notice = `${t('Replaced')}: ${changed}`;
      if (!all) {
        while (cursor < matches.length) {
          const candidate = matches[cursor]!;
          const sameTarget = candidate.slideIndex === match.slideIndex && candidate.shapeId === match.shapeId &&
            (candidate.target.kind === 'shape' || (match.target.kind === 'cell' && candidate.target.row === match.target.row && candidate.target.col === match.target.col));
          if (!sameTarget || candidate.start >= match.start + replacement.length) break;
          cursor++;
        }
      }
      cursor = matches.length ? cursor % matches.length : 0;
      reveal(current);
    } catch (cause) {
      error = `${t('Text could not be replaced')}: ${cause instanceof Error ? cause.message : String(cause)}`;
    }
  }
  function keydown(event: KeyboardEvent) {
    if (event.isComposing) return;
    if ((event.ctrlKey || event.metaKey) && ['f', 'h'].includes(event.key.toLowerCase())) {
      event.preventDefault(); searchInput.focus(); searchInput.select(); return;
    }
    if (event.key !== 'Enter' || event.target instanceof HTMLButtonElement) return;
    event.preventDefault(); move(event.shiftKey ? -1 : 1);
  }
</script>

<dialog bind:this={dialog} aria-label={t('Find and replace')} onclose={() => editor.closeDialog()} onkeydown={keydown}>
  <header><strong>{t('Find and replace')}</strong><button class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>×</button></header>
  <label>{t('Find text')}<input class="ok-input" bind:this={searchInput} bind:value={query} /></label>
  <label>{t('Replace with')}<input class="ok-input" bind:value={replacement} /></label>
  <div class="options">
    <label><input type="checkbox" bind:checked={matchCase} />{t('Match case')}</label>
    <label><input type="checkbox" bind:checked={currentSlideOnly} />{t('Only the starting slide')}</label>
  </div>
  <div class="navigation">
    <span role="status">{query ? (current ? `${t('Match')} ${Math.min(cursor + 1, matches.length)} / ${matches.length}` : t('No matches')) : t('Enter text to search')}</span>
    <button class="ok-btn" disabled={!matches.length} onclick={() => move(-1)}>{t('Previous match')}</button>
    <button class="ok-btn" disabled={!matches.length} onclick={() => move(1)}>{t('Next match')}</button>
  </div>
  {#if current}
    <section aria-label={t('Current match')}>
      <small>{t('Slide')} {current.slideIndex + 1}{current.target.kind === 'cell' ? ` · ${t('Cell')} ${current.target.row + 1}, ${current.target.col + 1}` : ''}</small>
      <p>{current.start > 60 ? '…' : ''}{current.text.slice(Math.max(0, current.start - 60), current.start)}<mark>{current.text.slice(current.start, current.end)}</mark>{current.text.slice(current.end, current.end + 90)}{current.end + 90 < current.text.length ? '…' : ''}</p>
    </section>
  {/if}
  {#if notice}<p role="status">{notice}</p>{/if}
  {#if error}<p role="alert">{error}</p>{/if}
  <footer><button class="ok-btn" disabled={!current} onclick={() => replace(false)}>{t('Replace match')}</button><button class="ok-btn primary" disabled={!current} onclick={() => replace(true)}>{t('Replace all')}</button></footer>
</dialog>

<style>
  dialog { width: min(540px, 90vw); max-height: 85vh; padding: 20px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0004; }
  header, footer, .navigation, .options { display: flex; align-items: center; gap: 10px; }
  header { justify-content: space-between; margin-bottom: 18px; }
  label { display: grid; gap: 6px; margin-bottom: 14px; }
  .options { flex-wrap: wrap; }
  .options label { display: flex; align-items: center; }
  .navigation { flex-wrap: wrap; margin: 10px 0; }
  .navigation span { margin-right: auto; }
  section { border: 1px solid var(--ok-border); border-radius: 6px; padding: 12px; margin: 14px 0; }
  p { white-space: pre-wrap; overflow-wrap: anywhere; line-height: 1.6; }
  mark { background: #ffdc69; color: #202020; }
  footer { justify-content: flex-end; margin-top: 18px; }
  [role='alert'] { color: #bf3131; }
</style>
