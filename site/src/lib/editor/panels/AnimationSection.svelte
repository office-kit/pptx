<script lang="ts">
  // The slide's animation pane: what plays, in what order, and what this
  // library cannot play. Effects are addressed by the shape they animate and
  // their place in the click order — the `<p:cTn id>` the library uses as a
  // handle is an implementation detail and never asked of the user.
  import {
    getGroupChildren,
    getShapeId,
    getShapeKind,
    getShapeName,
    getShapeText,
    getSlideAnimations,
    getSlideShapes,
    moveSlideAnimation,
    removeSlideAnimation,
    setShapeAnimation,
    updateSlideAnimation,
    type AnimationDirection,
    type AnimationEffect,
    type AnimationStartCondition,
    type SlideAnimationStep,
    type SlideData,
    type SlideShapeData,
  } from '@office-kit/pptx';
  import { untrack } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { selectedShapeId } from '../core/selection.ts';
  import AnimationPlayback from '../ui/AnimationPlayback.svelte';

  const editor = getEditor();
  const doc = editor.doc;

  const EFFECTS: { value: AnimationEffect; label: string }[] = [
    { value: 'fadeIn', label: 'Fade in' },
    { value: 'appear', label: 'Appear' },
    { value: 'flyIn', label: 'Fly in' },
    { value: 'zoomIn', label: 'Zoom in' },
    { value: 'spin', label: 'Spin' },
    { value: 'fadeOut', label: 'Fade out' },
    { value: 'disappear', label: 'Disappear' },
    { value: 'flyOut', label: 'Fly out' },
    { value: 'zoomOut', label: 'Zoom out' },
  ];
  // What the file records is an edge of the slide, and the same edge means
  // "from there" for an entrance and "out through there" for an exit. The
  // labels name the edge, so the effect beside them says which of the two it
  // is rather than the list having to say it twice.
  const DIRECTIONS: { value: AnimationDirection; label: string }[] = [
    { value: 'bottom', label: 'Bottom edge' },
    { value: 'top', label: 'Top edge' },
    { value: 'left', label: 'Left edge' },
    { value: 'right', label: 'Right edge' },
  ];
  const FLYING: AnimationEffect[] = ['flyIn', 'flyOut'];
  const STARTS: { value: AnimationStartCondition; label: string }[] = [
    { value: 'click', label: 'On click' },
    { value: 'withPrevious', label: 'With previous' },
    { value: 'afterPrevious', label: 'After previous' },
  ];

  const slide = $derived.by<SlideData | null>(() => {
    doc.version;
    return doc.currentSlide;
  });

  /** Every shape the slide draws, including the ones inside groups. */
  const shapes = $derived.by<SlideShapeData[]>(() => {
    doc.version;
    const out: SlideShapeData[] = [];
    const walk = (list: readonly SlideShapeData[]): void => {
      for (const shape of list) {
        out.push(shape);
        if (getShapeKind(shape) === 'group') walk(getGroupChildren(shape));
      }
    };
    if (slide) walk(getSlideShapes(slide));
    return out;
  });

  /** "Title 1 — Quarterly review": the name the deck gives it, plus a hint. */
  function shapeLabel(shape: SlideShapeData): string {
    const name = getShapeName(shape);
    const text = getShapeText(shape).replace(/\s+/g, ' ').trim();
    if (text === '') return name;
    const excerpt = text.length > 28 ? `${text.slice(0, 28)}…` : text;
    return `${name} — ${excerpt}`;
  }

  const labelById = $derived(new Map(shapes.map((shape) => [getShapeId(shape), shapeLabel(shape)])));

  function targetLabel(step: SlideAnimationStep): string {
    const ids = step.targetShapeIds;
    if (ids.length === 0) return t('Unknown object');
    const names = ids.map((id) => labelById.get(id) ?? `#${id}`);
    return names.join(', ');
  }

  const steps = $derived.by<SlideAnimationStep[]>(() => {
    doc.version;
    return slide ? [...getSlideAnimations(slide)] : [];
  });
  /** Each effect's place in the click order `moveSlideAnimation` counts in. */
  const places = $derived.by(() => {
    const map = new Map<SlideAnimationStep, number>();
    for (const step of steps) if (step.sequence === 'mainSeq') map.set(step, map.size);
    return map;
  });
  const mainSeqCount = $derived(places.size);

  function reasonFor(step: SlideAnimationStep): string {
    if (step.sequence === 'interactiveSeq')
      return t('Runs when its own object is clicked, not in the slide’s click order.');
    if (step.sequence !== 'mainSeq') return t('Belongs to a sequence this library only reads.');
    if (!step.playable)
      return t('Uses an effect, target or start condition this library only reads.');
    return t('Has no unique handle in the file, so it can be read but not changed.');
  }

  let error = $state('');
  function apply(label: string, operation: () => void): void {
    try {
      doc.transact(t(label), operation);
      error = '';
    } catch (cause) {
      error = cause instanceof Error ? cause.message : String(cause);
    }
  }

  function positionOf(step: SlideAnimationStep): number {
    return places.get(step) ?? -1;
  }
  function move(step: SlideAnimationStep, by: number): void {
    const target = slide;
    const id = step.id;
    const place = positionOf(step);
    const next = place + by;
    if (!target || id === null || place < 0 || next < 0 || next >= mainSeqCount) return;
    apply('Reorder animation', () => moveSlideAnimation(target, id, next));
  }

  function patch(step: SlideAnimationStep, change: Parameters<typeof updateSlideAnimation>[2]): void {
    const target = slide;
    const id = step.id;
    if (!target || id === null) return;
    apply('Change animation', () => updateSlideAnimation(target, id, change));
  }

  function remove(step: SlideAnimationStep): void {
    const target = slide;
    const id = step.id;
    if (!target || id === null) return;
    apply('Delete animation', () => removeSlideAnimation(target, id));
  }

  // --- Adding ------------------------------------------------------------
  const selected = $derived.by<number | null>(() => {
    doc.version;
    const sel = doc.selection;
    return sel.kind === 'shape' ? selectedShapeId(sel) : null;
  });
  // A shape id is only a handle within its own slide, and the object it named
  // may have been deleted, so the choice is re-checked against the slide in
  // hand rather than carried over to whatever now holds that id.
  const slideIndex = $derived(doc.selection.slideIndex);
  const shapeIds = $derived(new Set(shapes.map(getShapeId)));
  let chosen = $state<{ slide: number; id: number } | null>(null);
  const addTarget = $derived.by<number | null>(() => {
    if (chosen !== null && chosen.slide === slideIndex && shapeIds.has(chosen.id)) return chosen.id;
    if (selected !== null && shapeIds.has(selected)) return selected;
    const first = shapes[0];
    return first ? getShapeId(first) : null;
  });
  let addEffect = $state<AnimationEffect>('fadeIn');
  let addDirection = $state<AnimationDirection>('bottom');
  let addStart = $state<AnimationStartCondition>('click');
  let addDuration = $state(500);
  let addDelay = $state(0);
  let addByParagraph = $state(false);

  function add(): void {
    const shape = shapes.find((item) => getShapeId(item) === addTarget);
    if (!shape) return;
    apply('Add animation', () =>
      setShapeAnimation(shape, {
        effect: addEffect,
        ...(FLYING.includes(addEffect) ? { direction: addDirection } : {}),
        durationMs: addDuration,
        start: addStart,
        delayMs: addDelay,
        byParagraph: addByParagraph,
      }),
    );
  }

  // --- Playback ----------------------------------------------------------
  // Play reads the document in hand, not the saved deck: an effect changed a
  // moment ago plays as it is now, without waiting for a save to land.
  let playing = $state(false);
  let playingAt = { version: -1, slide: -1 };
  const playbackSvg = $derived.by(() => {
    doc.version;
    return doc.currentSvg;
  });
  function startPlaying(): void {
    playingAt = { version: doc.version, slide: slideIndex };
    playing = true;
  }
  // An edit, an undo or a move to another slide while the show is open leaves
  // it playing a deck that no longer exists, so it closes instead.
  $effect(() => {
    const version = doc.version;
    const at = doc.selection.slideIndex;
    untrack(() => {
      if (playing && (version !== playingAt.version || at !== playingAt.slide)) playing = false;
    });
  });
</script>

{#if slide}
  <section aria-label={t('Animations')}>
    <strong>{t('Animations')}</strong>

    {#if steps.length === 0}
      <p class="empty">{t('This slide has no animations yet.')}</p>
    {:else}
      <ol class="steps">
        {#each steps as step, position (position)}
          {@const place = positionOf(step)}
          <li>
            <div class="row head">
              <span class="order" aria-hidden="true">{place >= 0 ? place + 1 : '–'}</span>
              <span class="target">{targetLabel(step)}</span>
            </div>
            {#if step.editable}
              <div class="row">
                <label
                  >{t('Effect')}
                  <select
                    class="ok-input"
                    aria-label="{t('Effect')} {place + 1}"
                    value={step.effect}
                    onchange={(event) => {
                      const next = event.currentTarget.value as AnimationEffect;
                      // A preset that flies needs an edge, and one that does
                      // not refuses to be given one.
                      patch(step, {
                        effect: next,
                        ...(FLYING.includes(next) ? { direction: step.direction ?? 'bottom' } : {}),
                      });
                    }}
                  >
                    {#each EFFECTS as item}<option value={item.value}>{t(item.label)}</option>{/each}
                  </select>
                </label>
                {#if step.direction !== null}
                  <label
                    >{t('Direction')}
                    <select
                      class="ok-input"
                      aria-label="{t('Direction')} {place + 1}"
                      value={step.direction}
                      onchange={(event) =>
                        patch(step, {
                          direction: event.currentTarget.value as AnimationDirection,
                        })}
                    >
                      {#each DIRECTIONS as item}<option value={item.value}>{t(item.label)}</option
                        >{/each}
                    </select>
                  </label>
                {/if}
                <label
                  >{t('Start')}
                  <select
                    class="ok-input"
                    aria-label="{t('Start')} {place + 1}"
                    value={step.start}
                    onchange={(event) =>
                      patch(step, {
                        start: event.currentTarget.value as AnimationStartCondition,
                      })}
                  >
                    {#each STARTS as item}<option value={item.value}>{t(item.label)}</option>{/each}
                  </select>
                </label>
              </div>
              <div class="row">
                <label
                  >{t('Duration (ms)')}
                  <input
                    class="ok-input"
                    type="number"
                    min="0"
                    step="100"
                    aria-label="{t('Duration (ms)')} {place + 1}"
                    value={step.durationMs ?? 500}
                    onchange={(event) =>
                      patch(step, { durationMs: Number(event.currentTarget.value) })}
                  />
                </label>
                <label
                  >{t('Delay (ms)')}
                  <input
                    class="ok-input"
                    type="number"
                    min="0"
                    step="100"
                    aria-label="{t('Delay (ms)')} {place + 1}"
                    value={step.delayMs ?? 0}
                    onchange={(event) => patch(step, { delayMs: Number(event.currentTarget.value) })}
                  />
                </label>
              </div>
              <div class="row">
                <label class="check"
                  ><input
                    type="checkbox"
                    aria-label="{t('By paragraph')} {place + 1}"
                    checked={step.buildByParagraph}
                    onchange={(event) =>
                      patch(step, { byParagraph: event.currentTarget.checked })} />{t(
                    'By paragraph',
                  )}</label
                >
                <span class="spacer"></span>
                <button
                  class="ok-btn"
                  aria-label="{t('Move earlier')} {place + 1}"
                  disabled={place <= 0}
                  onclick={() => move(step, -1)}>↑</button
                >
                <button
                  class="ok-btn"
                  aria-label="{t('Move later')} {place + 1}"
                  disabled={place < 0 || place >= mainSeqCount - 1}
                  onclick={() => move(step, 1)}>↓</button
                >
                <button
                  class="ok-btn"
                  aria-label="{t('Delete animation')} {place + 1}"
                  onclick={() => remove(step)}>{t('Delete')}</button
                >
              </div>
            {:else}
              <p class="readonly">
                {t('Kept as it is')}: {reasonFor(step)}
              </p>
            {/if}
          </li>
        {/each}
      </ol>
    {/if}

    <div class="add">
      <strong>{t('Add animation')}</strong>
      <label
        >{t('Object')}
        <select
          class="ok-input"
          aria-label={t('Object')}
          value={addTarget ?? ''}
          onchange={(event) =>
            (chosen = { slide: slideIndex, id: Number(event.currentTarget.value) })}
        >
          {#each shapes as shape}
            <option value={getShapeId(shape)}>{shapeLabel(shape)}</option>
          {/each}
        </select>
      </label>
      <div class="row">
        <label
          >{t('Effect')}
          <select class="ok-input" aria-label={t('New effect')} bind:value={addEffect}>
            {#each EFFECTS as item}<option value={item.value}>{t(item.label)}</option>{/each}
          </select>
          {#if FLYING.includes(addEffect)}
            <select class="ok-input" aria-label={t('New direction')} bind:value={addDirection}>
              {#each DIRECTIONS as item}<option value={item.value}>{t(item.label)}</option>{/each}
            </select>
          {/if}
        </label>
        <label
          >{t('Start')}
          <select class="ok-input" aria-label={t('New start')} bind:value={addStart}>
            {#each STARTS as item}<option value={item.value}>{t(item.label)}</option>{/each}
          </select>
        </label>
      </div>
      <div class="row">
        <label
          >{t('Duration (ms)')}
          <input
            class="ok-input"
            type="number"
            min="0"
            step="100"
            aria-label={t('New duration (ms)')}
            bind:value={addDuration}
          />
        </label>
        <label
          >{t('Delay (ms)')}
          <input
            class="ok-input"
            type="number"
            min="0"
            step="100"
            aria-label={t('New delay (ms)')}
            bind:value={addDelay}
          />
        </label>
      </div>
      <label class="check"
        ><input type="checkbox" aria-label={t('New by paragraph')} bind:checked={addByParagraph} />{t(
          'By paragraph',
        )}</label
      >
      <button class="ok-btn" disabled={addTarget === null} onclick={add}>{t('Add animation')}</button
      >
    </div>

    <button class="ok-btn" disabled={steps.length === 0} onclick={startPlaying}
      >{t('Play animations')}</button
    >
    {#if error}<p role="alert">{t('Animation update failed')}: {error}</p>{/if}
  </section>
{/if}

{#if playing && slide}
  <AnimationPlayback svg={playbackSvg} steps={steps} onclose={() => (playing = false)} />
{/if}

<style>
  section {
    display: grid;
    gap: 10px;
    padding: 12px;
    border-bottom: 1px solid var(--ok-border);
  }
  strong {
    font-size: 12px;
  }
  .empty,
  .readonly {
    font-size: 11px;
    color: var(--ok-muted);
    margin: 0;
  }
  .steps {
    display: grid;
    gap: 8px;
    margin: 0;
    padding: 0;
    list-style: none;
  }
  .steps li {
    display: grid;
    gap: 6px;
    padding: 8px;
    border: 1px solid var(--ok-border);
    border-radius: 6px;
  }
  .row {
    display: flex;
    gap: 8px;
    align-items: end;
  }
  .row > label {
    flex: 1;
    min-width: 0;
  }
  .head {
    align-items: center;
    font-size: 11px;
  }
  .order {
    min-width: 18px;
    height: 18px;
    display: grid;
    place-items: center;
    border-radius: 9px;
    background: var(--ok-border);
    font-size: 10px;
  }
  .target {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .spacer {
    flex: 1;
  }
  label {
    display: grid;
    gap: 6px;
    font-size: 11px;
  }
  .check {
    display: flex;
    align-items: center;
    gap: 4px;
  }
  .add {
    display: grid;
    gap: 8px;
    padding-top: 8px;
    border-top: 1px dashed var(--ok-border);
  }
  [role='alert'] {
    color: #bf3131;
    font-size: 11px;
  }
</style>
