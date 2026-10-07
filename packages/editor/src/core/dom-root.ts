// `mountEditor` renders the editor inside a shadow root. From there, window and
// document listeners see events retargeted to the shadow host, and
// `document.activeElement` / `document.getSelection()` stop at the host (Chromium
// and WebKit report the host, not the node inside). These helpers resolve the
// real node whether the editor is in a shadow root or in the page's own tree.

/** The innermost target of an event, including one inside a shadow root. */
export function eventTarget(event: Event): EventTarget | null {
  return event.composedPath()[0] ?? event.target;
}

// Duck-typed rather than `instanceof ShadowRoot`: the editor's Node tests run
// these modules without DOM globals.
function isShadowRoot(node: Node): node is ShadowRoot {
  return 'host' in node && 'mode' in node;
}

/** The document or shadow root that contains `node`. */
export function rootOf(node: Node): Document | ShadowRoot {
  const root = node.getRootNode();
  return isShadowRoot(root) ? root : (node.ownerDocument ?? document);
}

/** The focused element in the tree that contains `node`. */
export function activeElementOf(node: Node): Element | null {
  return rootOf(node).activeElement;
}

export interface SelectionPoints {
  anchorNode: Node;
  anchorOffset: number;
  focusNode: Node;
  focusOffset: number;
}

/**
 * The document selection's anchor and focus as nodes in `node`'s tree. Inside a
 * shadow root, `Selection.anchorNode` is retargeted to the host, so the range is
 * read with `getComposedRanges`, and `direction` restores which end is the focus.
 */
export function selectionPoints(node: Node): SelectionPoints | null {
  const selection = (node.ownerDocument ?? document).getSelection();
  if (!selection?.rangeCount) return null;
  const root = rootOf(node);
  if (!isShadowRoot(root)) {
    if (!selection.anchorNode || !selection.focusNode) return null;
    return {
      anchorNode: selection.anchorNode,
      anchorOffset: selection.anchorOffset,
      focusNode: selection.focusNode,
      focusOffset: selection.focusOffset,
    };
  }
  const [range] = selection.getComposedRanges({ shadowRoots: [root] });
  if (!range) return null;
  const start = { node: range.startContainer, offset: range.startOffset };
  const end = { node: range.endContainer, offset: range.endOffset };
  const [anchor, focus] = selection.direction === 'backward' ? [end, start] : [start, end];
  return {
    anchorNode: anchor.node,
    anchorOffset: anchor.offset,
    focusNode: focus.node,
    focusOffset: focus.offset,
  };
}
