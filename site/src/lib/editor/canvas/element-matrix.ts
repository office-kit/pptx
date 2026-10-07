/**
 * The matrix taking `element`'s border-box pixels to viewport pixels,
 * including CSS transforms on it and its positioned ancestors.
 * `getBoundingClientRect` only reports axis-aligned bounds, which loses the
 * direction a rotated text box's lines run in.
 */
export function elementScreenMatrix(element: HTMLElement): DOMMatrix {
  let outermost: HTMLElement | null = null;
  for (let node: HTMLElement | null = element; node; node = node.parentElement)
    if (getComputedStyle(node).transform !== 'none') outermost = node;
  let matrix = new DOMMatrix();
  let node = element;
  // Inside the outermost transform, offsets are untransformed layout positions;
  // above it, the viewport rectangle is exact. The editor's transformed layers
  // are all absolutely positioned, so the offsetParent chain visits each one.
  while (outermost) {
    matrix = localTransform(node).multiply(matrix);
    const parent = node.offsetParent as HTMLElement;
    matrix = new DOMMatrix()
      .translate(
        node.offsetLeft + parent.clientLeft - parent.scrollLeft,
        node.offsetTop + parent.clientTop - parent.scrollTop,
      )
      .multiply(matrix);
    node = parent;
    if (node !== outermost && node.contains(outermost)) break;
  }
  const rect = node.getBoundingClientRect();
  return new DOMMatrix().translate(rect.left, rect.top).multiply(matrix);
}

function localTransform(node: HTMLElement): DOMMatrix {
  const style = getComputedStyle(node);
  if (style.transform === 'none') return new DOMMatrix();
  const [x = 0, y = 0] = style.transformOrigin.split(' ').map(parseFloat);
  return new DOMMatrix().translate(x, y).multiply(new DOMMatrix(style.transform)).translate(-x, -y);
}
