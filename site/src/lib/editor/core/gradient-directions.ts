import type { ReadGradientFill } from '@office-kit/pptx';

// Mac PowerPoint's Radial and Rectangular galleries order and saved edge insets.
export const pathDirections = [
  { label: 'From Bottom Right Corner', x: 1, y: 1 },
  { label: 'From Bottom Left Corner', x: 0, y: 1 },
  { label: 'From Center', x: 0.5, y: 0.5 },
  { label: 'From Top Right Corner', x: 1, y: 0 },
  { label: 'From Top Left Corner', x: 0, y: 0 },
].map(({ label, x, y }) => ({
  label,
  x,
  y,
  focus: { left: x, top: y, right: 1 - x, bottom: 1 - y },
  tileRect: {
    left: x === 0 ? -1 : 0,
    top: y === 0 ? -1 : 0,
    right: x === 1 ? -1 : 0,
    bottom: y === 1 ? -1 : 0,
  },
}));

export function pathDirectionIndex(gradient: ReadGradientFill | null): number | undefined {
  if (!gradient || !['circle', 'rect', 'shape'].includes(gradient.path ?? '')) return undefined;
  const focus = gradient.focus;
  const index = pathDirections.findIndex(
    (direction) =>
      focus &&
      direction.focus.left === focus.left &&
      direction.focus.top === focus.top &&
      direction.focus.right === focus.right &&
      direction.focus.bottom === focus.bottom,
  );
  return index < 0 ? undefined : index;
}
