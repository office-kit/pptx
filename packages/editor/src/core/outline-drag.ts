import type { OutlineSelectionField } from './outline-selection.ts';

/** Outline indentation per level in CSS px; RichTextInput's outline styles use the same step. */
export const OUTLINE_LEVEL_STEP_PX = 11.5;
/** Pointer travel before a bullet press becomes a drag. */
export const OUTLINE_DRAG_THRESHOLD_PX = 4;

const PARAGRAPH = '[data-outline-paragraph]';

/**
 * The body paragraph whose bullet gutter (the indent before its text) is under
 * `x`. Titles have no bullet; their slide icon is the drag handle instead.
 */
export function bulletParagraphAt(
  field: OutlineSelectionField,
  target: Element,
  x: number,
): number | null {
  const section = target.closest<HTMLElement>(PARAGRAPH);
  if (!section || !field.root.contains(section) || section.hasAttribute('data-outline-title'))
    return null;
  const gutter = parseFloat(getComputedStyle(section).paddingLeft);
  if (x >= section.getBoundingClientRect().left + gutter) return null;
  return [...field.root.querySelectorAll(PARAGRAPH)].indexOf(section);
}

export type OutlineParagraphBox = {
  field: OutlineSelectionField;
  index: number;
  top: number;
  bottom: number;
  left: number;
};

/** Every rendered outline paragraph in display order. */
export function outlineParagraphBoxes(fields: OutlineSelectionField[]): OutlineParagraphBox[] {
  return fields.flatMap((field) => {
    const left = field.root.getBoundingClientRect().left;
    return [...field.root.querySelectorAll<HTMLElement>(PARAGRAPH)].map((section, index) => {
      const rect = section.getBoundingClientRect();
      return { field, index, top: rect.top, bottom: rect.bottom, left };
    });
  });
}

/**
 * The paragraph a drop at `y` lands after. A drop above the first paragraph
 * (the first slide's title) has no outline position.
 */
export function boundaryBefore(
  boxes: OutlineParagraphBox[],
  y: number,
): OutlineParagraphBox | null {
  let best: OutlineParagraphBox | null = null;
  let distance = Infinity;
  for (const [position, box] of boxes.entries()) {
    // The boundary after `box` sits between its bottom and the next top.
    const next = boxes[position + 1];
    const line = next ? (box.bottom + next.top) / 2 : box.bottom;
    const gap = Math.abs(y - line);
    if (gap < distance) {
      best = box;
      distance = gap;
    }
  }
  const first = boxes[0];
  return first && y < (first.top + first.bottom) / 2 ? null : best;
}
