/**
 * Object animation playback against a rendered slide.
 *
 * One implementation, three callers: the development preview's presentation
 * mode, its presenter view, and the editor's animation panel. They share it so
 * they cannot disagree about what a timing tree means — the preview bundles it
 * over HTTP (`/animation-player.js`), the panel imports it directly.
 *
 * The rule it is built around: **play only what the deck states.** A step this
 * library reads but cannot reproduce is never approximated — not by guessing a
 * moment for it, not by giving it a click of its own, and not by deciding for
 * itself whether the shape should be on the slide before its turn. Such a step
 * is reported as unsupported and everything it touches is left exactly as the
 * renderer drew it, which is the only state that cannot lose content.
 *
 * What that means in detail:
 *
 *   - Only the main sequence. An `interactiveSeq` runs off its own trigger —
 *     a click on some shape — and is not part of the slide's click order, so
 *     it neither plays nor hides anything here.
 *   - A `click` step opens a stop, wherever it appears: a click stop is stated
 *     by the tree, so a group of known effects after an unsupported one still
 *     plays. `withPrevious` runs from the start of the group in hand and
 *     `afterPrevious` from its end. The first step of a slide has no
 *     predecessor, so a leading `withPrevious` / `afterPrevious` starts as the
 *     slide appears rather than waiting for a click.
 *   - A step whose start, length, delay, target or preset we cannot read makes
 *     the rest of *its own group* unplaceable: an `afterPrevious` chained onto
 *     an effect of unknown length has no stated moment, and the `withPrevious`
 *     effects sharing that start have none either. That span is unsupported.
 *     The next click stop starts from a stated moment again.
 *   - A shape id that is drawn more than once is not a handle. The schema puts
 *     no uniqueness constraint on `<p:cNvPr id>`, and nothing says a timing
 *     tree that names it means every one of them, so the step is unsupported
 *     rather than animating all of them together.
 */

import type { SlideAnimationStep } from '@office-kit/pptx';

const INSTANT_EFFECTS = ['appear', 'disappear'];
const ENTRANCE_EFFECTS = ['fadeIn', 'appear', 'flyIn', 'zoomIn'];
const EXIT_EFFECTS = ['fadeOut', 'disappear', 'flyOut', 'zoomOut'];
const EMPHASIS_EFFECTS = ['spin'];

/**
 * What an effect does to its target beyond putting it on the slide or taking it
 * off. Each one is reproduced here as the browser's own equivalent of the
 * behaviour the deck states: `style.opacity` as `opacity`, `ppt_x` / `ppt_y` as
 * a translation, `ppt_w` / `ppt_h` as a scale about the shape's centre, and `r`
 * as a rotation about it.
 */
type Motion = 'none' | 'fade' | 'fly' | 'zoom' | 'spin';

const MOTIONS: Readonly<Record<string, Motion>> = {
  appear: 'none',
  disappear: 'none',
  fadeIn: 'fade',
  fadeOut: 'fade',
  flyIn: 'fly',
  flyOut: 'fly',
  zoomIn: 'zoom',
  zoomOut: 'zoom',
  spin: 'spin',
};

/** One clockwise turn, which is what the `spin` preset this library reads is. */
const SPIN_DEGREES = 360;

/** Why a step is not played. */
export type UnsupportedReason =
  /** It belongs to a sequence with its own trigger, not the slide's clicks. */
  | 'notInMainSeq'
  /** Its effect, target or delay is one this library only reads. */
  | 'notModelled'
  /** Its moment depends on an effect whose length or place is not stated. */
  | 'unknownTiming'
  /** More than one shape on the slide carries the id it names. */
  | 'ambiguousTarget';

export interface UnsupportedAnimation {
  readonly step: SlideAnimationStep;
  readonly reason: UnsupportedReason;
}

/** What one step does to its target, and when, within its stop. */
export interface AnimationItem {
  readonly step: SlideAnimationStep;
  /**
   * `'emphasis'` animates a shape that is already on the slide and leaves it
   * there, so unlike the other two it never decides whether the shape is shown.
   */
  readonly kind: 'entrance' | 'exit' | 'emphasis';
  /** Milliseconds from the start of the stop. */
  readonly begin: number;
  /** Null when the tree states no length. */
  readonly duration: number | null;
}

/** One click's worth of animation. */
export interface AnimationStop {
  readonly items: AnimationItem[];
  /** True when the slide starts it itself, with no click — its first step
   * runs with or after a predecessor it does not have. */
  readonly auto: boolean;
}

export interface AnimationPlan {
  readonly stops: AnimationStop[];
  readonly unsupported: UnsupportedAnimation[];
}

export interface AnimationProgress {
  readonly cursor: number;
  readonly stops: number;
  /** How long the stop in hand has been running; null once it has settled. */
  readonly elapsed: number | null;
  /** Bumped by every call that sets where the slide stands. */
  readonly generation: number;
}

export interface AnimationPlayerOptions {
  /**
   * Where the slide was drawn (a shadow root is fine), or a function returning
   * it. A function is asked again on every query, because a transition draws
   * the slide into a layer of its own and the layer beside it holds a
   * different slide whose shapes may carry the very same ids.
   */
  readonly root: ParentNode | (() => ParentNode);
  readonly steps: readonly SlideAnimationStep[] | null | undefined;
  /**
   * Asked afresh on every advance, so a viewer who turns reduced motion on
   * mid-show gets the rest of the deck without motion — in the same order.
   */
  readonly reducedMotion?: () => boolean;
  /**
   * Called whenever the cursor moves and whenever the slide changes by itself
   * — a delayed effect arriving, an effect finishing — so a view mirroring the
   * slide never lags behind what the audience can see.
   */
  readonly onChange?: (player: AnimationPlayer) => void;
}

export interface AnimationPlayer {
  readonly cursor: number;
  readonly stopCount: number;
  /** True while a click still has an effect to play on this slide. */
  readonly pending: boolean;
  /** True while an effect already started is still on its way. */
  readonly running: boolean;
  readonly progress: AnimationProgress;
  /**
   * The steps this player does not run, so a caller can say so rather than let
   * a slide look as though it has no animation.
   */
  readonly unsupported: UnsupportedAnimation[];
  reset(): void;
  /** Plays the next stop. False when the slide has none left. */
  advance(): boolean;
  /** Takes one stop back, showing everything before it. False at the start. */
  back(): boolean;
  /** Moves straight to a point in the click order, without motion. */
  jumpTo(next: number): void;
  /**
   * Shows the moment another view of the same slide is showing: the stops
   * before `next` played out, and the one it is on started `elapsed` ago — so
   * an effect waiting on a delay over there has not arrived here either, and
   * one half-way through a fade is half-way through it here.
   */
  resume(next: number, elapsed: number | null): void;
  /** Everything the slide ends with — what a still preview shows. */
  finish(): void;
  dispose(): void;
}

/** An element whose inline style we can set. */
type StyledElement = Element & ElementCSSInlineStyle;

const styled = (el: Element): StyledElement | null =>
  el instanceof HTMLElement || el instanceof SVGElement ? el : null;

/** Whether a step reveals its target, hides it, or neither; null when we cannot tell. */
const kindOf = (step: SlideAnimationStep): AnimationItem['kind'] | null => {
  if (step.effect === null) return null;
  if (ENTRANCE_EFFECTS.includes(step.effect)) return 'entrance';
  if (EXIT_EFFECTS.includes(step.effect)) return 'exit';
  if (EMPHASIS_EFFECTS.includes(step.effect)) return 'emphasis';
  return null;
};

/**
 * How long an effect runs, or null when the deck does not say. Null is not
 * zero: it is the reason nothing can be chained onto the end of this effect.
 */
const durationOf = (step: SlideAnimationStep): number | null => {
  if (step.durationMs !== null && Number.isFinite(step.durationMs) && step.durationMs >= 0)
    return step.durationMs;
  // 'appear' and 'disappear' are instantaneous by definition of the preset, so
  // a tree that states no duration for one is not leaving anything unsaid.
  if (step.effect !== null && INSTANT_EFFECTS.includes(step.effect)) return 0;
  return null;
};

/**
 * Groups the steps into the stops the viewer clicks through, with each step's
 * start measured from the beginning of its stop, and lists the steps that are
 * left out with the reason why.
 */
export const buildAnimationStops = (
  steps: readonly SlideAnimationStep[] | null | undefined,
): AnimationPlan => {
  const stops: AnimationStop[] = [];
  const unsupported: UnsupportedAnimation[] = [];
  const skip = (step: SlideAnimationStep, reason: UnsupportedReason): void => {
    unsupported.push({ step, reason });
  };
  let stop: AnimationStop | null = null;
  let groupStart: number | null = 0;
  let groupEnd: number | null = 0;
  let seen = false;
  for (const step of steps ?? []) {
    if (step.sequence !== 'mainSeq') {
      skip(step, 'notInMainSeq');
      continue;
    }
    const leading = !seen;
    seen = true;
    if (step.start === 'unknown') {
      // Nothing states whether this opens a click stop or joins the one in
      // hand, so it is neither played nor allowed to place what follows it.
      groupStart = null;
      groupEnd = null;
      skip(step, 'notModelled');
      continue;
    }
    if (step.start === 'click' || leading) {
      stop = { items: [], auto: leading && step.start !== 'click' };
      stops.push(stop);
      groupStart = 0;
      groupEnd = 0;
    } else if (step.start === 'afterPrevious') {
      groupStart = groupEnd;
      groupEnd = groupStart;
    }
    // 'withPrevious' joins the group in hand: same start, same end so far.
    const kind = step.playable ? kindOf(step) : null;
    const duration = kind === null ? null : durationOf(step);
    const delay = step.delayMs;
    if (kind === null || delay === null || groupStart === null) {
      // Its place in the click order is still known — it keeps its stop — but
      // its moment, or what it does, is not. Nothing may be chained onto it.
      skip(step, kind === null || delay === null ? 'notModelled' : 'unknownTiming');
      groupEnd = null;
      continue;
    }
    // Annotated because `groupEnd` below feeds `groupStart` on the next pass,
    // which would make the inference circular.
    const begin: number = groupStart + delay;
    stop?.items.push({ step, kind, begin, duration });
    groupEnd = duration === null || groupEnd === null ? null : Math.max(groupEnd, begin + duration);
  }
  return { stops, unsupported };
};

const shapeSelector = (id: string): string => `[data-pptx-shape-id="${id}"]`;

/**
 * The elements one step animates: a whole object, or paragraphs of its text.
 * Empty when the step's target is one we do not act on.
 */
const targetsOf = (
  root: ParentNode,
  step: SlideAnimationStep,
  blocked: ReadonlySet<string>,
): StyledElement[] => {
  const target = step.target;
  // 'unsupported' names a shape only as a hint about what an effect touches,
  // not as the thing to animate, so it is never resolved to an element.
  if (target.kind !== 'shape' && target.kind !== 'paragraphs') return [];
  if (!Number.isInteger(target.shapeId)) return [];
  const id = String(target.shapeId);
  if (blocked.has(id)) return [];
  const shapes = root.querySelectorAll(shapeSelector(id));
  const shape = shapes.length === 1 ? styled(shapes[0]!) : null;
  if (shape === null) return [];
  if (target.kind !== 'paragraphs') return [shape];
  const out: StyledElement[] = [];
  // One pass over the markers that were actually drawn: a paragraph range may
  // name far more paragraphs than the shape has, and a group nested inside this
  // shape brings paragraphs of its own that belong to a different target. The
  // nearest shape of *any* id is the one a paragraph belongs to.
  for (const marker of shape.querySelectorAll('[data-pptx-paragraph]')) {
    if (marker.closest('[data-pptx-shape-id]') !== shape) continue;
    const n = Number(marker.getAttribute('data-pptx-paragraph'));
    if (n < target.firstParagraph || n > target.lastParagraph) continue;
    const element = styled(marker);
    if (element !== null) out.push(element);
  }
  return out;
};

const nowMs = (): number => (typeof performance === 'object' ? performance.now() : Date.now());

/**
 * The slide's own box on screen. A fly is stated against the slide's edges
 * (`1+#ppt_h/2` puts the shape's centre one half-height below the bottom of it),
 * so that is what the distance has to be measured from — never the viewport,
 * which is not where the slide ends.
 */
const slideBoxOf = (el: Element, root: ParentNode): DOMRect | null => {
  const svg = el.closest('svg');
  if (svg !== null) return svg.getBoundingClientRect();
  return root instanceof Element ? root.getBoundingClientRect() : null;
};

/**
 * The matrix taking the element's own coordinates to the screen, which is also
 * the space a CSS transform written on it is read in.
 *
 * The shape and paragraph markers carry no transform of their own — the
 * renderer wraps the drawn shape rather than transforming the wrapper — so the
 * matrix does not have to be un-done first, and a shape the deck rotates or a
 * group it scales is accounted for by the ancestors it comes through. Text laid
 * out as HTML sits in a `<foreignObject>`, whose own matrix is the one its
 * pixels are scaled by.
 */
const screenMatrixOf = (el: StyledElement): DOMMatrix | null => {
  if (el instanceof SVGGraphicsElement) return el.getScreenCTM();
  const host = el.closest('foreignObject');
  return host instanceof SVGGraphicsElement ? host.getScreenCTM() : null;
};

/**
 * How far the element has to move to sit just outside the named edge, in the
 * units its own CSS transform is written in. `null` when the slide has not been
 * laid out yet, or its matrix cannot be inverted.
 */
const flyOffset = (
  el: StyledElement,
  direction: 'top' | 'right' | 'bottom' | 'left',
  root: ParentNode,
): { readonly dx: number; readonly dy: number } | null => {
  const slide = slideBoxOf(el, root);
  if (slide === null || slide.width <= 0 || slide.height <= 0) return null;
  const box = el.getBoundingClientRect();
  // Clear of the edge by the element's own extent: the far side of the box has
  // to reach the edge before the near side goes past it.
  const dx =
    direction === 'right'
      ? slide.right - box.left
      : direction === 'left'
        ? slide.left - box.right
        : 0;
  const dy =
    direction === 'bottom'
      ? slide.bottom - box.top
      : direction === 'top'
        ? slide.top - box.bottom
        : 0;
  const matrix = screenMatrixOf(el);
  // Plain HTML with no SVG above it: its pixels are the screen's.
  if (matrix === null) return { dx, dy };
  const determinant = matrix.a * matrix.d - matrix.b * matrix.c;
  if (determinant === 0) return null;
  return {
    dx: (matrix.d * dx - matrix.c * dy) / determinant,
    dy: (matrix.a * dy - matrix.b * dx) / determinant,
  };
};

/** What one effect runs on one element, and whether it needs the element's own box. */
interface Movement {
  readonly frames: Keyframe[];
  /**
   * Scaling and turning happen about the shape's centre, which is the shape's
   * own bounding box rather than the slide's — the default reference box for an
   * SVG element is the viewport, so it has to be said.
   */
  readonly aboutOwnCentre: boolean;
}

/**
 * Each motion animates a property of its own — `translate`, `scale`, `rotate`
 * — rather than all of them sharing `transform`.
 *
 * Two reasons, both of them things the deck states and a single `transform`
 * would lose. The first is that PresentationML animates a shape's position, its
 * size and its rotation as separate attributes (`ppt_x`/`ppt_y`, `ppt_w`/
 * `ppt_h`, `r`), so a slide may well run a fly and a spin over one shape at the
 * same moment; written to one property, the animation that begins second
 * replaces the first outright and the shape would only turn. The second is that
 * the drawn slide uses `transform` itself, for the rotation and the group
 * scaling the deck gives a shape — and the individual properties apply on top
 * of it rather than in place of it, so nothing the renderer did is overwritten.
 */

/**
 * Whether a step's motion has to be measured against the drawn slide, so that a
 * movement of `null` means "not read yet" rather than "nothing to run".
 */
const needsGeometry = (item: AnimationItem): boolean =>
  item.step.effect !== null && MOTIONS[item.step.effect] === 'fly';

/**
 * The keyframes one step runs on one element, or `null` when there is nothing
 * to run: an instant effect, or a fly whose distance cannot be measured. Either
 * way the element is put where the step leaves it without motion, which for
 * every effect modelled here is the element's own untransformed place.
 */
const movementFor = (item: AnimationItem, el: StyledElement, root: ParentNode): Movement | null => {
  const effect = item.step.effect;
  const entering = item.kind === 'entrance';
  switch (effect === null ? 'none' : (MOTIONS[effect] ?? 'none')) {
    case 'none':
      return null;
    case 'fade':
      return {
        frames: entering ? [{ opacity: 0 }, { opacity: 1 }] : [{ opacity: 1 }, { opacity: 0 }],
        aboutOwnCentre: false,
      };
    case 'zoom': {
      const none = { scale: '0' };
      const full = { scale: '1' };
      return { frames: entering ? [none, full] : [full, none], aboutOwnCentre: true };
    }
    case 'spin':
      return {
        frames: [{ rotate: '0deg' }, { rotate: `${SPIN_DEGREES}deg` }],
        aboutOwnCentre: true,
      };
    case 'fly': {
      const direction = item.step.direction;
      if (direction === null) return null;
      const offset = flyOffset(el, direction, root);
      if (offset === null) return null;
      const away = { translate: `${offset.dx}px ${offset.dy}px` };
      const place = { translate: '0px 0px' };
      return { frames: entering ? [away, place] : [place, away], aboutOwnCentre: false };
    }
  }
};

/** Plays a slide's animations against a rendered SVG. */
export const createAnimationPlayer = (options: AnimationPlayerOptions): AnimationPlayer => {
  const root = (): ParentNode =>
    typeof options.root === 'function' ? options.root() : options.root;
  const parsed = buildAnimationStops(options.steps);
  const stops = parsed.stops;
  const unsupported = [...parsed.unsupported];
  const reduced = (): boolean => options.reducedMotion?.() === true;
  const onChange = options.onChange ?? ((): void => {});
  const animations: Animation[] = [];
  const timers: ReturnType<typeof setTimeout>[] = [];
  const touched = new Set<StyledElement>();
  // Which elements an effect is currently driving, and which effects of the
  // stop in hand have run. Two effects of one stop can animate the same object
  // — a fade-in the deck fades straight back out, say — and the order they
  // happen to finish in is not the order the deck lists them in. Settling from
  // the list rather than from whichever animation ended last is what makes
  // playing forward land where a jump to the same point lands.
  const inFlight = new Map<StyledElement, number>();
  // Animations that have ended on an element another effect is still animating.
  // SMIL leaves such an animation in the sandwich, frozen at its last value and
  // still above everything that began before it, so an effect that ends early
  // does not hand the object back to a longer one it outranks — unless the tree
  // says its value is taken away when it ends, and then it does. The step's
  // `valueAfterEnd` is what decides, never an assumption; a step that leaves it
  // unstated never reaches here beside another effect on the same element.
  // They are dropped — and the element's own style put back in charge — once
  // nothing is animating it.
  const frozen = new Map<StyledElement, Animation[]>();
  const outstanding = new Map<AnimationItem, number>();
  const finished = new Set<AnimationItem>();
  let cursor = 0;
  let disposed = false;
  let startedAt: number | null = null;
  // Bumped by every call that sets where the slide stands, so a second view can
  // tell an explicit seek from a passing report — including a seek that lands
  // on the cursor it is already on and turns a running stop into a settled one.
  let generation = 0;

  // Every shape an unsupported step names — `targetShapeIds` covers a composite
  // effect that drives several — is left the way the renderer drew it. We know
  // something animates it and cannot reproduce that, and of the two states we
  // could choose, only 'shown' is unable to hide content the deck does show.
  const blocked = new Set<string>();
  for (const entry of unsupported)
    for (const shapeId of entry.step.targetShapeIds) blocked.add(String(shapeId));
  const drawn = new Map<string, number>();
  for (const stop of stops)
    for (const item of stop.items) {
      const id = String(item.step.target.shapeId);
      if (!drawn.has(id)) drawn.set(id, root().querySelectorAll(shapeSelector(id)).length);
    }
  for (const stop of stops)
    for (const item of stop.items)
      if ((drawn.get(String(item.step.target.shapeId)) ?? 0) > 1)
        unsupported.push({ step: item.step, reason: 'ambiguousTarget' });
  for (const [id, count] of drawn) if (count > 1) blocked.add(id);

  /**
   * Where two effects animate one element at the same time, what the slide
   * shows when the first of them ends turns on whether its value stays above
   * the other or gives way to it. `valueAfterEnd` says which; where the tree
   * does not say, there is nothing to play, so the step is reported and what it
   * names is left as the renderer drew it. On its own such a step is no
   * question at all — nothing is underneath it, and its own visibility decides
   * where the shape is left — so it plays as usual.
   *
   * Read per element, which is where "the same thing" is actually settled: a
   * build's paragraphs are separate elements. Each element's effects are sorted
   * once and swept twice, for what begins before an earlier one has ended and
   * what is still running when a later one begins — the pairs themselves are
   * never enumerated.
   *
   * Run again whenever the slide is drawn again, because which elements a step
   * reaches is a fact about the drawing: a transition hands over a layer, and a
   * re-render replaces the nodes underneath us.
   */
  const silence = (current: ParentNode): void => {
    const meets = new Set<AnimationItem>();
    for (const stop of stops) {
      const spans = new Map<StyledElement, { item: AnimationItem; begin: number; end: number }[]>();
      for (const item of stop.items) {
        const end = item.duration === null ? Number.POSITIVE_INFINITY : item.begin + item.duration;
        for (const el of targetsOf(current, item.step, blocked)) {
          const span = { item, begin: item.begin, end };
          const list = spans.get(el);
          if (list === undefined) spans.set(el, [span]);
          else list.push(span);
        }
      }
      for (const list of spans.values()) {
        if (list.length < 2) continue;
        list.sort((a, b) => a.begin - b.begin);
        let furthest = Number.NEGATIVE_INFINITY;
        for (const span of list) {
          if (span.begin < furthest) meets.add(span.item);
          furthest = Math.max(furthest, span.end);
        }
        let earliest = Number.POSITIVE_INFINITY;
        for (let at = list.length - 1; at >= 0; at -= 1) {
          const span = list[at];
          if (span === undefined) continue;
          if (span.end > earliest) meets.add(span.item);
          earliest = Math.min(earliest, span.begin);
        }
      }
    }
    // Removed in one pass, writing the ones that stay over the ones that go,
    // rather than a splice per removal.
    for (const stop of stops) {
      let kept = 0;
      for (const item of stop.items) {
        if (item.step.valueAfterEnd === 'unstated' && meets.has(item)) {
          unsupported.push({ step: item.step, reason: 'notModelled' });
          for (const shapeId of item.step.targetShapeIds) blocked.add(String(shapeId));
          continue;
        }
        stop.items[kept] = item;
        kept += 1;
      }
      stop.items.length = kept;
    }
  };

  const clearPending = (): void => {
    for (const timer of timers) clearTimeout(timer);
    timers.length = 0;
    for (const animation of animations) animation.cancel();
    animations.length = 0;
    for (const list of frozen.values()) for (const animation of list) animation.cancel();
    frozen.clear();
    inFlight.clear();
    outstanding.clear();
    finished.clear();
  };
  const drop = <T>(list: T[], entry: T): void => {
    const at = list.indexOf(entry);
    if (at >= 0) list.splice(at, 1);
  };
  const notifyChanged = (): void => {
    if (disposed) return;
    // Once the last effect of a stop has run, the slide is settled: saying so
    // is what stops a second view from resuming into a stop that has ended.
    if (timers.length === 0 && animations.length === 0) startedAt = null;
    onChange(player);
  };
  // The reference box a scale or a turn is measured against. Set while such an
  // effect runs and taken off once nothing is animating the element, so the
  // player leaves no style of its own on a settled slide. It is safe to set
  // because the markers it is set on carry no transform themselves — the
  // renderer wraps the drawn shape rather than transforming the wrapper — so
  // there is nothing already on the element for a changed origin to move.
  const centreOn = (el: StyledElement): void => {
    touched.add(el);
    el.style.transformBox = 'fill-box';
    el.style.transformOrigin = '50% 50%';
  };
  const uncentre = (el: StyledElement): void => {
    el.style.transformBox = '';
    el.style.transformOrigin = '';
  };
  const show = (el: StyledElement): void => {
    touched.add(el);
    el.style.visibility = '';
    el.style.opacity = '';
    uncentre(el);
  };
  const hide = (el: StyledElement): void => {
    touched.add(el);
    el.style.visibility = 'hidden';
    el.style.opacity = '';
    uncentre(el);
  };

  // Which effects touch each element and which elements each effect touches,
  // resolved once per drawing rather than per effect that finishes. A slide
  // drawn again — a transition handing its layer over, a re-render — gives a
  // different node or leaves the cached ones detached, and is indexed afresh.
  interface Entry {
    readonly item: AnimationItem;
    readonly stopIndex: number;
    /** Its rank among the effects on its element; the higher one decides. */
    order: number;
  }
  /**
   * The effects that touch one element, lowest priority first, and how far
   * through them the element stands.
   *
   * Priority is the order the effects *begin* in, not the order they are written
   * in. PresentationML timing is SMIL's timing and animation model (ECMA-376
   * Part 1 §19.5), and SMIL composites the animations of one attribute so that
   * "the animation first begun has lowest priority and the most recently begun
   * animation has highest priority" (SMIL 3.0 §12.4.3, the animation sandwich).
   * The document decides only between animations that begin together, and then
   * a time dependent "is considered to activate after the syncbase element" —
   * which in a timing tree always follows what it waits on — so sorting on the
   * begin time alone, stably, leaves document order to break exactly the ties
   * the spec leaves to it.
   *
   * A step written first but delayed past a later one therefore outranks it,
   * and reading the standing from this list rather than from whichever
   * animation happened to end last is what makes playing forward, jumping,
   * stepping back and resuming agree.
   */
  interface ElementPlan {
    readonly list: Entry[];
    /**
     * The effects that decide whether the element is on the slide at all, which
     * is every one of them but the emphasis effects: those animate a shape that
     * is already there and hand it back unchanged, so one that has run says
     * nothing about whether the shape is shown.
     */
    readonly gates: Entry[];
    /** The highest-priority gate that has run: the one that decides. */
    top: Entry | null;
    /** Whether the first gate to begin — the lowest priority — has run. */
    firstPlayed: boolean;
  }
  /** One effect's hold on one element, so finishing it costs no search. */
  interface Placement {
    readonly el: StyledElement;
    readonly plan: ElementPlan;
    readonly entry: Entry;
    /**
     * What this effect runs on this element, measured against the slide as it
     * was drawn. Read once per drawing so that playing forward, seeking and
     * resuming all work from the same geometry; `null` when the effect has
     * nothing to run here and the element simply takes its settled place.
     */
    readonly movement: Movement | null;
  }
  let indexedRoot: ParentNode | null = null;
  // Whether the last index found the geometry the effects in it need. An effect
  // that has to know where the slide's edges are cannot be measured before the
  // browser has laid the slide out, and a zero read then is not an answer.
  let laidOut = true;
  let plans = new Map<StyledElement, ElementPlan>();
  let placements = new Map<AnimationItem, Placement[]>();
  /** Every stop below this has run; within the stop in hand, `finished` says. */
  let settledBelow = 0;
  const hasPlayed = (entry: Entry): boolean =>
    entry.stopIndex < settledBelow || finished.has(entry.item);

  /** Reads every element's standing from its effects, for a redraw or a seek. */
  const recompute = (): void => {
    for (const plan of plans.values()) {
      plan.top = null;
      for (const entry of plan.gates) if (hasPlayed(entry)) plan.top = entry;
      const first = plan.gates[0];
      plan.firstPlayed = first !== undefined && hasPlayed(first);
    }
  };

  const index = (): void => {
    const current = root();
    const sample = plans.keys().next();
    // An index that found nothing is not an index: the slide may simply not
    // have been drawn yet, and a deck that really animates nothing costs one
    // empty pass to say so again. Nor is one taken before the browser laid the
    // slide out, where every distance an effect has to travel reads as zero.
    if (current === indexedRoot && sample.done !== true && sample.value.isConnected && laidOut)
      return;
    indexedRoot = current;
    laidOut = true;
    silence(current);
    plans = new Map();
    placements = new Map();
    stops.forEach((stop, stopIndex) => {
      for (const item of stop.items) {
        const spots: Placement[] = [];
        for (const el of targetsOf(current, item.step, blocked)) {
          const plan = plans.get(el) ?? { list: [], gates: [], top: null, firstPlayed: false };
          plans.set(el, plan);
          const entry: Entry = { item, stopIndex, order: 0 };
          plan.list.push(entry);
          if (item.kind !== 'emphasis') plan.gates.push(entry);
          const movement = movementFor(item, el, current);
          if (movement === null && needsGeometry(item)) laidOut = false;
          spots.push({ el, plan, entry, movement });
        }
        placements.set(item, spots);
      }
    });
    for (const plan of plans.values()) {
      // Stable, so effects that begin together keep the document's order.
      const byBegin = (a: Entry, b: Entry): number =>
        a.stopIndex - b.stopIndex || a.item.begin - b.item.begin;
      plan.list.sort(byBegin);
      plan.gates.sort(byBegin);
      plan.list.forEach((entry, at) => {
        entry.order = at;
      });
    }
    recompute();
  };

  /** Where an element stands: the highest-priority effect that has run wins. */
  const visibilityOf = (plan: ElementPlan): boolean => {
    const first = plan.gates[0];
    if (first === undefined) return true;
    // An entrance that has not run yet keeps its target off the slide, even
    // when a later step would have shown it.
    if (first.item.kind === 'entrance' && !plan.firstPlayed) return false;
    return plan.top === null || plan.top.item.kind !== 'exit';
  };

  /** An element an effect is animating is left to that effect until it ends. */
  const render = (el: StyledElement, plan: ElementPlan): void => {
    if (inFlight.has(el)) return;
    if (visibilityOf(plan)) show(el);
    else hide(el);
  };

  /**
   * Puts every animated element where it stands once `played` stops have run,
   * with no motion. Idempotent, so back, jump and a cancelled advance all land
   * on the same picture as playing forward to that point.
   */
  const settle = (played: number): void => {
    finished.clear();
    settledBelow = played;
    index();
    recompute();
    for (const [el, plan] of plans) render(el, plan);
  };

  /**
   * One effect has had its say on one element. Only that element can have
   * changed, and only against the effect deciding it so far — so a stop of many
   * effects costs one comparison each, not a re-reading of the slide each.
   */
  const markPlayed = (spot: Placement): void => {
    const { plan, entry } = spot;
    if (entry.item.kind !== 'emphasis') {
      if (entry === plan.gates[0]) plan.firstPlayed = true;
      if (plan.top === null || entry.order > plan.top.order) plan.top = entry;
    }
    render(spot.el, plan);
  };

  /** One effect has had its say everywhere it reaches. */
  const complete = (item: AnimationItem): void => {
    finished.add(item);
    index();
    for (const spot of placements.get(item) ?? []) markPlayed(spot);
  };

  // `into` is how far through the effect already is — 0 for one starting now,
  // more for one a second view is joining part-way through.
  const play = (item: AnimationItem, into: number): void => {
    index();
    const spots = placements.get(item) ?? [];
    if (spots.length === 0) {
      complete(item);
      return;
    }
    // The elements this effect actually animates. One it reaches but has
    // nothing to run on — an instant preset, or a fly whose distance the slide
    // did not give us — simply takes the place the effect leaves it in, which
    // for every effect modelled here is the element's own untransformed one.
    const moving = spots.filter((spot) => spot.movement !== null);
    const duration = item.duration;
    const from = Math.max(0, into);
    // Over before it could be seen: instantaneous, or joined after its end.
    const over = duration === null || duration === 0 || from >= duration;
    // Such an effect normally just sets the element's own style. It cannot
    // while another effect is animating that element, because an animation's
    // value is applied above the element's style — so it goes into the sandwich
    // instead, frozen at the value it ends on. Unless the tree says that value
    // is taken away when it ends: then there is nothing to hold, and the effect
    // underneath is the one that shows.
    const holds = item.step.valueAfterEnd !== 'removed';
    if (
      reduced() ||
      moving.length === 0 ||
      (over && (!holds || !moving.some((spot) => inFlight.has(spot.el))))
    ) {
      complete(item);
      return;
    }
    // Nothing is going to animate these, so they are settled the moment the
    // effect has its turn.
    for (const spot of spots) if (spot.movement === null) markPlayed(spot);
    const length = over ? 0 : duration;
    outstanding.set(item, moving.length);
    for (const spot of moving) {
      const el = spot.el;
      const movement = spot.movement!;
      touched.add(el);
      // Only an effect that decides whether the shape is on the slide may put
      // it there. An emphasis effect animates whatever is already drawn, so a
      // shape an entrance is still holding back — or one an exit has taken away
      // — stays off the slide while it turns, rather than showing for as long
      // as the turn lasts. `render` cannot undo it either: it leaves an element
      // an effect is animating to that effect until it ends.
      if (item.kind !== 'emphasis') el.style.visibility = '';
      if (movement.aboutOwnCentre) centreOn(el);
      inFlight.set(el, (inFlight.get(el) ?? 0) + 1);
      const animation = el.animate(movement.frames, { duration: length, fill: 'both' });
      animation.currentTime = over ? length : from;
      animations.push(animation);
      void animation.finished.then(
        () => {
          if (disposed) return;
          drop(animations, animation);
          const left = (inFlight.get(el) ?? 1) - 1;
          if (left <= 0 || !holds) {
            // Nothing is animating the element any more, so the sandwich comes
            // off and its own style says where it stands — or this effect's
            // value is one the tree takes away, and it leaves the sandwich now.
            if (left <= 0) {
              inFlight.delete(el);
              for (const other of frozen.get(el) ?? []) other.cancel();
              frozen.delete(el);
            } else inFlight.set(el, left);
            animation.cancel();
          } else {
            inFlight.set(el, left);
            const held = frozen.get(el);
            if (held === undefined) frozen.set(el, [animation]);
            else held.push(animation);
          }
          markPlayed(spot);
          const remaining = (outstanding.get(item) ?? 1) - 1;
          outstanding.set(item, remaining);
          if (remaining <= 0) complete(item);
          notifyChanged();
        },
        () => {},
      );
    }
  };

  const runStop = (stopIndex: number, elapsed: number): void => {
    const stop = stops[stopIndex];
    if (stop === undefined) return;
    const since = Math.max(0, elapsed);
    startedAt = nowMs() - since;
    // In the order they begin, stably, so the animations are built in the order
    // SMIL ranks them: the last one built is the one on top of the sandwich.
    // Joining a stop part-way through starts several at once, and the tree's
    // order is not theirs.
    const ordered = [...stop.items].sort((a, b) => a.begin - b.begin);
    for (const item of ordered) {
      const wait = item.begin - since;
      if (reduced() || wait <= 0) {
        play(item, since - item.begin);
        continue;
      }
      const timer = setTimeout(() => {
        drop(timers, timer);
        play(item, 0);
        notifyChanged();
      }, wait);
      timers.push(timer);
    }
  };

  const clamp = (next: number): number =>
    Math.max(0, Math.min(Number.isFinite(next) ? Math.round(next) : 0, stops.length));

  const player: AnimationPlayer = {
    get cursor() {
      return cursor;
    },
    get stopCount() {
      return stops.length;
    },
    get pending() {
      return cursor < stops.length;
    },
    get running() {
      return timers.length > 0 || animations.length > 0;
    },
    get progress() {
      return {
        cursor,
        stops: stops.length,
        elapsed: startedAt === null ? null : Math.round(nowMs() - startedAt),
        generation,
      };
    },
    get unsupported() {
      return [...unsupported];
    },
    reset() {
      generation += 1;
      clearPending();
      startedAt = null;
      cursor = 0;
      settle(0);
      onChange(player);
      // A slide whose first effect runs with or after 'the previous one' has no
      // previous one: it starts as the slide appears, without a click.
      if (stops[0]?.auto === true) player.advance();
    },
    advance() {
      if (disposed || cursor >= stops.length) return false;
      generation += 1;
      clearPending();
      settle(cursor);
      const stopIndex = cursor;
      cursor += 1;
      runStop(stopIndex, 0);
      onChange(player);
      return true;
    },
    back() {
      if (disposed || cursor <= 0) return false;
      generation += 1;
      clearPending();
      startedAt = null;
      cursor -= 1;
      settle(cursor);
      onChange(player);
      return true;
    },
    jumpTo(next) {
      if (disposed) return;
      generation += 1;
      clearPending();
      startedAt = null;
      cursor = clamp(next);
      settle(cursor);
      onChange(player);
    },
    resume(next, elapsed) {
      if (disposed) return;
      generation += 1;
      clearPending();
      startedAt = null;
      cursor = clamp(next);
      if (cursor > 0 && elapsed !== null && Number.isFinite(elapsed)) {
        settle(cursor - 1);
        runStop(cursor - 1, elapsed);
      } else settle(cursor);
      onChange(player);
    },
    finish() {
      player.jumpTo(stops.length);
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      clearPending();
      for (const el of touched) {
        el.style.visibility = '';
        el.style.opacity = '';
        uncentre(el);
      }
      touched.clear();
    },
  };
  // So that a caller asking what will not play gets an answer before the first
  // click, rather than only once the slide has been read for a seek.
  index();
  return player;
};
