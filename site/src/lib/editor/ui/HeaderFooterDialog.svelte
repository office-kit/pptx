<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getSlides } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { applyHeaderFooter, isTitleSlide, readHeaderFooter } from '../core/header-footer.ts';
  import { selectedSlideIndices } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const editor = getEditor();
  const doc = editor.doc;
  const targets = untrack(() => {
    const all = getSlides(doc.pres);
    return selectedSlideIndices(doc.selection).flatMap((index) => (all[index] ? [all[index]!] : []));
  });
  const initial = targets[0] ? readHeaderFooter(targets[0]) : { date: null, slideNumber: false, footer: null };
  let dialog: HTMLDialogElement;
  let date = $state(initial.date !== null);
  let dateKind = $state<'auto' | 'fixed'>(initial.date?.kind ?? 'auto');
  let format = $state<'datetime1' | 'datetime2'>(initial.date?.kind === 'auto' ? initial.date.format : 'datetime1');
  let fixed = $state(initial.date?.kind === 'fixed' ? initial.date.text : new Date().toLocaleDateString());
  let slideNumber = $state(initial.slideNumber);
  let footer = $state(initial.footer !== null);
  let footerText = $state(initial.footer ?? '');
  let hideOnTitle = $state(untrack(() => getSlides(doc.pres).some(isTitleSlide)) && !initial.slideNumber);
  const sample = new Date();
  onMount(() => dialog.showModal());

  function apply(all: boolean) {
    const settings = {
      date: !date ? null : dateKind === 'auto' ? { kind: 'auto' as const, format } : { kind: 'fixed' as const, text: fixed },
      slideNumber,
      footer: footer ? footerText : null,
      hideOnTitle,
    };
    doc.transact(t('Header & Footer'), () => {
      for (const slide of all ? getSlides(doc.pres) : targets) applyHeaderFooter(doc.pres, slide, settings);
    });
    editor.closeDialog();
  }
</script>

<dialog bind:this={dialog} aria-label={t('Header and Footer')} onclose={() => editor.closeDialog()}>
  <form onsubmit={(event) => { event.preventDefault(); apply(false); }}>
    <header><strong>{t('Header and Footer')}</strong><button type="button" class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    <fieldset>
      <legend>{t('Include on slide')}</legend>
      <label class="check"><input type="checkbox" bind:checked={date} />{t('Date and time')}</label>
      <div class="indent" class:off={!date}>
        <label class="check"><input type="radio" name="date-kind" value="auto" disabled={!date} bind:group={dateKind} />{t('Update automatically')}</label>
        <select class="ok-input" aria-label={t('Date format')} disabled={!date || dateKind !== 'auto'} bind:value={format}>
          <option value="datetime1">{sample.toLocaleDateString()}</option>
          <option value="datetime2">{sample.toLocaleDateString(undefined, { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</option>
        </select>
        <label class="check"><input type="radio" name="date-kind" value="fixed" disabled={!date} bind:group={dateKind} />{t('Fixed')}</label>
        <input class="ok-input" aria-label={t('Fixed date')} disabled={!date || dateKind !== 'fixed'} bind:value={fixed} />
      </div>
      <label class="check"><input type="checkbox" bind:checked={slideNumber} />{t('Slide number')}</label>
      <label class="check"><input type="checkbox" bind:checked={footer} />{t('Footer')}</label>
      <input class="ok-input indent" aria-label={t('Footer text')} disabled={!footer} bind:value={footerText} />
    </fieldset>
    <label class="check"><input type="checkbox" bind:checked={hideOnTitle} />{t("Don't show on title slide")}</label>
    <footer>
      <button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button>
      <button type="button" class="ok-btn" onclick={() => apply(true)}>{t('Apply to All')}</button>
      <button type="submit" class="ok-btn primary" disabled={targets.length === 0}>{t('Apply')}</button>
    </footer>
  </form>
</dialog>

<style>
  dialog { width: min(440px, 90vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form { display: grid; gap: 14px; }
  header, footer { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
  footer { justify-content: flex-end; }
  fieldset { display: grid; gap: 8px; margin: 0; padding: 10px 12px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius); }
  legend { padding: 0 4px; font-size: 12px; color: var(--ok-text-2); }
  .check { display: flex; align-items: center; gap: 6px; }
  .indent { display: grid; gap: 6px; margin-left: 22px; }
  .off { opacity: 0.6; }
</style>
