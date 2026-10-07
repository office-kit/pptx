import {
  getParagraphPropertiesEffective,
  setParagraphIndent,
  setParagraphTabs,
  type PresentationData,
  type SlideShapeData,
  type TableCellData,
} from '@office-kit/pptx';
import { editTabStops, type TabStopEdit } from './paragraph-tabs.ts';

export interface Point {
  readonly x: number;
  readonly y: number;
}

/** The direction lines advance in the text's unrotated box. */
export type TextFlow = 'horizontal' | 'vertical' | 'vertical-reversed';

/** Where text-local distances land on one of the two screen-aligned rulers. */
export interface RulerAxis {
  readonly axis: 'x' | 'y';
  /** Screen point of the paragraph start (indent zero). */
  readonly origin: Point;
  /** Signed screen pixels per text-layout pixel along `axis`. */
  readonly scale: number;
}

export type IndentHandle = 'first' | 'hanging' | 'left';

/** One ruler gesture, shared by the live preview and the committed change. */
export type RulerChange =
  | { readonly kind: 'indent'; readonly handle: IndentHandle; readonly delta: number }
  | { readonly kind: 'tabs'; readonly edits: readonly TabStopEdit[] };

// ST_TextMargin / ST_TextIndent bound marL and indent to ±51206400 EMU.
export const MAX_INDENT_EMU = 51206400;

const NATURAL: Record<TextFlow, Point> = {
  horizontal: { x: 1, y: 0 },
  vertical: { x: 0, y: 1 },
  'vertical-reversed': { x: 0, y: -1 },
};

/**
 * Maps a text body's inline axis onto the axis-aligned rulers. Rotated text
 * is measured as if its shape were unrotated about `center`, so indents and
 * tabs keep their own lengths; vertical writing measures on the vertical
 * ruler. Both match Mac PowerPoint (native capture 2026-10-07: a 30° box
 * keeps an unrotated horizontal ruler at its unrotated left edge, and `vert`
 * text draws its markers on the vertical ruler). `origin` and `direction` are
 * the screen point and per-pixel vector of the text's own inline axis.
 */
export function rulerAxis(input: {
  readonly origin: Point;
  readonly direction: Point;
  readonly center: Point;
  readonly flow: TextFlow;
}): RulerAxis {
  const natural = NATURAL[input.flow];
  const turn = Math.atan2(natural.y, natural.x) - Math.atan2(input.direction.y, input.direction.x);
  const cos = Math.cos(turn);
  const sin = Math.sin(turn);
  const dx = input.origin.x - input.center.x;
  const dy = input.origin.y - input.center.y;
  const axis = input.flow === 'horizontal' ? 'x' : 'y';
  const length = Math.hypot(input.direction.x, input.direction.y);
  return {
    axis,
    origin: { x: input.center.x + dx * cos - dy * sin, y: input.center.y + dx * sin + dy * cos },
    scale: (axis === 'x' ? natural.x : natural.y) * length,
  };
}

/** The indentation a handle drag writes for a paragraph currently at `left`/`first`. */
export function indentAfterDrag(
  current: { readonly left: number; readonly first: number },
  handle: IndentHandle,
  delta: number,
): { leftEmu?: number; firstLineEmu?: number } {
  const clamp = (value: number, min: number) => Math.max(min, Math.min(MAX_INDENT_EMU, value));
  if (handle === 'first') return { firstLineEmu: clamp(current.first + delta, -MAX_INDENT_EMU) };
  const left = clamp(current.left + delta, 0);
  // The hanging handle moves the margin but keeps the first line in place.
  return handle === 'hanging'
    ? { leftEmu: left, firstLineEmu: clamp(current.first + current.left - left, -MAX_INDENT_EMU) }
    : { leftEmu: left };
}

/** Paragraph indentation as the ruler displays it (`marL` defaults by outline level). */
export function rulerIndent(props: {
  readonly marL: number | null;
  readonly indent: number | null;
  readonly level: number;
}): { left: number; first: number } {
  return { left: props.marL ?? props.level * 32 * 9525, first: props.indent ?? 0 };
}

/** Applies one ruler gesture to each paragraph, relative to that paragraph's own values. */
export function applyRulerChange(
  pres: PresentationData,
  target: SlideShapeData | TableCellData,
  indices: readonly number[],
  change: RulerChange,
  options: { inheritanceSource?: SlideShapeData } = {},
): void {
  for (const index of indices) {
    const props = getParagraphPropertiesEffective(pres, target, index, options);
    if (change.kind === 'indent')
      setParagraphIndent(
        target,
        index,
        indentAfterDrag(rulerIndent(props), change.handle, change.delta),
      );
    else
      setParagraphTabs(target, index, {
        tabStops: editTabStops(props.tabStops ?? [], change.edits),
      });
  }
}
