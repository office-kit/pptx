<script lang="ts">
  import { onMount, untrack, tick } from 'svelte';
  import { getSlides, getSlideTitle, addSlideComment, getSlideComments, getCommentAuthor, getCommentText, getCommentParent, getCommentStatus, type CommentStatus, type SlideCommentData, removeSlideComment, setCommentStatus, setCommentText } from '@office-kit/pptx';
  import { orderCommentThreads } from '../core/comment-threads.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  const editor = getEditor();
  const doc = editor.doc;
  const openedSlide = untrack(() => doc.currentSlide);
  const version = untrack(() => doc.version);
  const slides = untrack(() => getSlides(doc.pres));
  const original = slides.map(slide => getSlideComments(slide));
  // `status` is `null` for a comment that has nowhere to keep one — an
  // ECMA-376 `<p:cm>` — and for a draft that is not in the file yet.
  type Draft = { id: number; parent: number | null; comment: number | null; author: string; text: string; status: CommentStatus | null };
  const initialIndex = Math.max(0, slides.indexOf(openedSlide!));
  let reviewIndex = $state(initialIndex);
  let showResolved = $state(false);
  let nextId = original.reduce((max, comments) => Math.max(max, comments.length), 0);
  const newDraft = (parent: number | null = null): Draft => ({ id: nextId++, parent, comment: null, author: '', text: '', status: null });
  let slideDrafts = $state<Draft[][]>(original.map(comments => {
    const indices = new Map(comments.map((comment, index) => [comment, index]));
    return comments.map((comment, index) => ({ id: index, parent: indices.get(getCommentParent(comment)!) ?? null, comment: index, author: getCommentAuthor(comment).name, text: getCommentText(comment), status: getCommentStatus(comment) }));
  }));
  if (slideDrafts[initialIndex]?.length === 0) slideDrafts[initialIndex]!.push(newDraft());
  const drafts = $derived(slideDrafts[reviewIndex] ?? []);
  const ordered = $derived(orderCommentThreads(drafts));
  // An untouched new row is a placeholder, not an unfinished comment.
  const pending = $derived(slideDrafts.map(comments => comments.filter(draft => draft.comment !== null || draft.author.trim() || draft.text.trim())));
  const parents = $derived(new Set(drafts.flatMap(draft => draft.parent === null ? [] : [draft.parent])));
  const draftById = $derived(new Map(drafts.map(draft => [draft.id, draft])));
  const complete = $derived(pending.every(comments => {
    const ids = new Set(comments.map(draft => draft.id));
    return comments.every(draft => draft.text.trim() && (draft.comment !== null || draft.author.trim()) && (draft.parent === null || ids.has(draft.parent)));
  }));
  // Resolving a thread and changing nothing else is a change: without the
  // status here, Apply would stay disabled and the click would be lost.
  const changed = $derived(pending.some((comments, index) => comments.length !== original[index]!.length || comments.some(draft => draft.comment === null || draft.text !== getCommentText(original[index]![draft.comment]!) || draft.status !== getCommentStatus(original[index]![draft.comment]!))));
  // A resolved thread is out of the way by default, and its replies with it,
  // but it is still here to be brought back and reopened.
  const resolvedRoot = (draft: Draft): boolean => (draft.parent === null ? draft : draftById.get(draft.parent) ?? draft).status === 'resolved';
  const visible = $derived(showResolved ? ordered : ordered.filter(draft => !resolvedRoot(draft)));
  const hidden = $derived(ordered.length - visible.length);
  let error = $state('');
  let dialog: HTMLDialogElement;
  const valid = $derived(slides.length > 0 && complete && changed);
  onMount(() => dialog.showModal());
  async function addDraft(parent: number | null = null) {
    const draft = newDraft(parent);
    drafts.push(draft);
    await tick();
    dialog.querySelector<HTMLInputElement>(`[data-comment-id="${draft.id}"] input`)?.focus();
  }
  async function removeDraft(id: number) {
    const position = ordered.findIndex(draft => draft.id === id);
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
    const remaining = ordered.filter(draft => !removed.has(draft.id));
    const next = remaining[Math.min(position, remaining.length - 1)];
    slideDrafts[reviewIndex] = drafts.filter(draft => !removed.has(draft.id));
    await tick();
    const selector = next ? `[data-comment-id="${next.id}"] textarea` : '[data-add-comment]';
    dialog.querySelector<HTMLElement>(selector)?.focus();
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
              if (draft.status !== null && draft.status !== getCommentStatus(comment)) setCommentStatus(comment, draft.status);
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
    <label class="inline"><input type="checkbox" aria-label={t('Show resolved threads')} bind:checked={showResolved} /> {t('Show resolved threads')}{#if hidden > 0} ({hidden}){/if}</label>
    <div class="comments">
      {#each visible as draft, i (draft.id)}
        <section data-comment-id={draft.id} class:reply={draft.parent !== null} class:resolved={draft.status === 'resolved'} aria-label={`${t('Comment')} ${i + 1}`}>
          <header>{#if draft.comment !== null}<strong>{draft.author}</strong>{:else}<label>{t('Author name')}<input class="ok-input" aria-label={t('Author name')} required={!!draft.text.trim()} bind:value={draft.author} /></label>{/if}{#if draft.status === 'resolved'}<span class="badge">{t('Resolved')}</span>{/if}<button type="button" class="ok-btn" onclick={() => removeDraft(draft.id)}>{t(parents.has(draft.id) ? 'Delete thread' : 'Delete comment')}</button></header>
          {#if draft.parent !== null}<p>{t('Reply to')}: {draftById.get(draft.parent)?.author} — {draftById.get(draft.parent)?.text}</p>{/if}
          <label>{t('Comment text')}<textarea class="ok-input" aria-label={t('Comment text')} rows="3" required={draft.comment !== null || !!draft.author.trim()} bind:value={draft.text}></textarea></label>
          <div class="row">
            <button type="button" class="ok-btn" disabled={!draft.text.trim() || (draft.comment === null && !draft.author.trim())} onclick={() => addDraft(draft.id)}>{t('Reply')}</button>
            {#if draft.parent === null && draft.status !== null}
              <button type="button" class="ok-btn" aria-label={`${t(draft.status === 'resolved' ? 'Reopen' : 'Resolve')} ${i + 1}`} onclick={() => draft.status = draft.status === 'resolved' ? 'active' : 'resolved'}>{t(draft.status === 'resolved' ? 'Reopen' : 'Resolve')}</button>
            {/if}
          </div>
        </section>
      {:else}<p>{ordered.length > 0 ? t('Every thread on this slide is resolved.') : t('No comments on this slide.')}</p>{/each}
    </div>
    <button type="button" class="ok-btn" data-add-comment onclick={() => addDraft()}>{t('Add comment')}</button>
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
  section.resolved { opacity: 0.75; }
  .row { display: flex; gap: 8px; }
  .inline { display: flex; align-items: center; gap: 6px; font-size: 12px; }
  .badge { padding: 1px 8px; border: 1px solid var(--ok-border); border-radius: 999px; font-size: 11px; color: var(--ok-text-2); }
  section p { overflow-wrap: anywhere; }
  header, footer { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
  footer { justify-content: flex-end; }
  label { flex: 1; font-size: 12px; }
  textarea { resize: vertical; width: 100%; line-height: 1.5; }
  p { margin: 0; font-size: 12px; color: var(--ok-text-2); }
  [role='alert'] { color: #bf3131; }
</style>
