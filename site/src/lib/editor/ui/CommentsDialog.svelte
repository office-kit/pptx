<script lang="ts">
  import { onMount, untrack } from 'svelte';
  import { getSlides, getSlideTitle, addSlideComment, getSlideComments, getCommentAuthor, getCommentText, getCommentParent, type SlideCommentData, removeSlideComment, setCommentText } from '@office-kit/pptx';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const openedSlide = untrack(() => doc.currentSlide);
  const version = untrack(() => doc.version);
  const slides = untrack(() => getSlides(doc.pres));
  const original = slides.map(slide => getSlideComments(slide));
  type Draft = { id: number; parent: number | null; comment: number | null; author: string; text: string };
  const initialIndex = Math.max(0, slides.indexOf(openedSlide!));
  let reviewIndex = $state(initialIndex);
  let nextId = original.reduce((max, comments) => Math.max(max, comments.length), 0);
  const newDraft = (parent: number | null = null): Draft => ({ id: nextId++, parent, comment: null, author: '', text: '' });
  let slideDrafts = $state<Draft[][]>(original.map(comments => {
    const indices = new Map(comments.map((comment, index) => [comment, index]));
    return comments.map((comment, index) => ({ id: index, parent: indices.get(getCommentParent(comment)!) ?? null, comment: index, author: getCommentAuthor(comment).name, text: getCommentText(comment) }));
  }));
  if (slideDrafts[initialIndex]?.length === 0) slideDrafts[initialIndex]!.push(newDraft());
  const drafts = $derived(slideDrafts[reviewIndex] ?? []);
  // An untouched new row is a placeholder, not an unfinished comment.
  const pending = $derived(slideDrafts.map(comments => comments.filter(draft => draft.comment !== null || draft.author.trim() || draft.text.trim())));
  const parents = $derived(new Set(drafts.flatMap(draft => draft.parent === null ? [] : [draft.parent])));
  const draftById = $derived(new Map(drafts.map(draft => [draft.id, draft])));
  const complete = $derived(pending.every(comments => {
    const ids = new Set(comments.map(draft => draft.id));
    return comments.every(draft => draft.text.trim() && (draft.comment !== null || draft.author.trim()) && (draft.parent === null || ids.has(draft.parent)));
  }));
  const changed = $derived(pending.some((comments, index) => comments.length !== original[index]!.length || comments.some(draft => draft.comment === null || draft.text !== getCommentText(original[index]![draft.comment]!))));
  let error = $state('');
  let dialog: HTMLDialogElement;
  const valid = $derived(slides.length > 0 && complete && changed);
  onMount(() => dialog.showModal());
  function removeDraft(id: number) {
    const children = new Map<number, number[]>();
    for (const draft of drafts) if (draft.parent !== null) {
      const siblings = children.get(draft.parent) ?? [];
      siblings.push(draft.id);
      children.set(draft.parent, siblings);
    }
    const removed = new Set<number>();
    const queue = [id];
    while (queue.length) {
      const current = queue.pop()!;
      if (removed.has(current)) continue;
      removed.add(current);
      for (const child of children.get(current) ?? []) queue.push(child);
    }
    slideDrafts[reviewIndex] = drafts.filter(draft => !removed.has(draft.id));
  }
  function apply(event: SubmitEvent) {
    event.preventDefault();
    if (!valid) return;
    if (doc.version !== version || doc.currentSlide !== openedSlide) { error = t('The slide changed. Reopen this dialog.'); return; }
    try {
      doc.transact(t('Comments'), () => {
        for (const [slideIndex, comments] of pending.entries()) {
          const retained = new Set(comments.map(draft => draft.comment));
          const existing = original[slideIndex]!;
          for (const [index, comment] of existing.entries()) if (!retained.has(index)) removeSlideComment(comment);
          const handles = new Map<number, SlideCommentData>(existing.map((comment, index) => [index, comment]));
          for (const draft of comments) {
            if (draft.comment !== null) {
              const comment = existing[draft.comment]!;
              if (draft.text !== getCommentText(comment)) setCommentText(comment, draft.text);
            } else {
              const replyTo = draft.parent === null ? undefined : handles.get(draft.parent);
              if (draft.parent !== null && !replyTo) throw new Error(t('Reply parent is missing.'));
              handles.set(draft.id, addSlideComment(slides[slideIndex]!, { author: { name: draft.author.trim() }, text: draft.text, ...(replyTo ? { replyTo } : {}) }));
            }
          }
        }
      });
      editor.closeDialog();
    } catch (cause) { error = `${t('Comment update failed')}: ${cause instanceof Error ? cause.message : String(cause)}`; }
  }
</script>
<dialog bind:this={dialog} aria-label={t('Comments')} onclose={() => editor.closeDialog()}>
  <form onsubmit={apply}>
    <header><strong>{t('Comments')}</strong><button type="button" class="ok-btn" aria-label={t('Close')} onclick={() => editor.closeDialog()}>✕</button></header>
    <p>{t('Review comments across slides. Apply saves all your changes; Cancel discards them.')}</p>
    <label>{t('Review slide')}
      <select class="ok-input" aria-label={t('Review slide')} bind:value={reviewIndex}>
        {#each slides as slide, index}
          <option value={index}>{index + 1}. {getSlideTitle(slide) || t('Untitled slide')} — {pending[index]!.length} {t('Comments')}</option>
        {/each}
      </select>
    </label>
    <div class="comments">
      {#each drafts as draft, i (draft.id)}
        <section class:reply={draft.parent !== null} aria-label={`${t('Comment')} ${i + 1}`}>
          <header>{#if draft.comment !== null}<strong>{draft.author}</strong>{:else}<label>{t('Author name')}<input class="ok-input" aria-label={t('Author name')} required={!!draft.text.trim()} bind:value={draft.author} /></label>{/if}<button type="button" class="ok-btn" onclick={() => removeDraft(draft.id)}>{t(parents.has(draft.id) ? 'Delete thread' : 'Delete comment')}</button></header>
          {#if draft.parent !== null}<p>{t('Reply to')}: {draftById.get(draft.parent)?.author} — {draftById.get(draft.parent)?.text}</p>{/if}
          <label>{t('Comment text')}<textarea class="ok-input" aria-label={t('Comment text')} rows="3" required={draft.comment !== null || !!draft.author.trim()} bind:value={draft.text}></textarea></label>
          <button type="button" class="ok-btn reply-button" disabled={!draft.text.trim() || (draft.comment === null && !draft.author.trim())} onclick={() => drafts.push(newDraft(draft.id))}>{t('Reply')}</button>
        </section>
      {:else}<p>{t('No comments on this slide.')}</p>{/each}
    </div>
    <button type="button" class="ok-btn" onclick={() => drafts.push(newDraft())}>{t('Add comment')}</button>
    {#if !complete}<p role="status">{t('Complete or delete unfinished comments on all slides before applying.')}</p>{/if}
    {#if error}<p role="alert">{error}</p>{/if}
    <footer><button type="button" class="ok-btn" onclick={() => editor.closeDialog()}>{t('Cancel')}</button><button type="submit" class="ok-btn primary" disabled={!valid}>{t('Apply')}</button></footer>
  </form>
</dialog>
<style>
  dialog { width: min(640px, 90vw); max-height: 85vh; padding: 18px; border: 1px solid var(--ok-border); border-radius: var(--ok-radius-lg); background: var(--ok-panel); color: var(--ok-text); box-shadow: var(--ok-shadow-lg); }
  dialog::backdrop { background: #0006; }
  form, label, .comments, section { display: grid; gap: 12px; }
  .comments { max-height: 50vh; overflow: auto; }
  section { padding: 12px; border: 1px solid var(--ok-border); border-radius: 6px; }
  section.reply { margin-left: 20px; border-left: 3px solid var(--ok-border); }
  .reply-button { justify-self: start; }
  section p { overflow-wrap: anywhere; }
  header, footer { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
  footer { justify-content: flex-end; }
  label { flex: 1; font-size: 12px; }
  textarea { resize: vertical; width: 100%; line-height: 1.5; }
  p { margin: 0; font-size: 12px; color: var(--ok-text-2); }
  [role='alert'] { color: #bf3131; }
</style>
