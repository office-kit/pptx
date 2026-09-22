export type TextSelection = { start: number; end: number };

/** Contenteditable can create DIV/BR nodes while typing; count them as text newlines. */
export function richTextValue(root: HTMLElement, stop?: { node: Node; offset: number }): string {
  let text = '';
  let stopped = false;
  function visit(node: Node) {
    if (stopped) return;
    if (node.nodeType === Node.TEXT_NODE) {
      const value = node.textContent ?? '';
      text += node === stop?.node ? value.slice(0, stop.offset) : value;
      if (node === stop?.node) stopped = true;
      return;
    }
    if (node instanceof HTMLElement && node.tagName === 'BR') {
      if (!node.hasAttribute('data-caret-end')) text += '\n';
      return;
    }
    const block = node !== root && node instanceof HTMLElement && /^(DIV|P)$/.test(node.tagName);
    if (block && text && !text.endsWith('\n')) text += '\n';
    for (let i = 0; i <= node.childNodes.length; i++) {
      if (node === stop?.node && i === stop.offset) {
        stopped = true;
        return;
      }
      if (i < node.childNodes.length) visit(node.childNodes[i]!);
      if (stopped) return;
    }
    if (block && node.nextSibling && !text.endsWith('\n')) text += '\n';
  }
  visit(root);
  return text;
}

export function richTextSelection(root: HTMLElement): TextSelection | null {
  const selection = root.ownerDocument.getSelection();
  if (!selection?.rangeCount) return null;
  const range = selection.getRangeAt(0);
  if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) return null;
  return {
    start: richTextValue(root, { node: range.startContainer, offset: range.startOffset }).length,
    end: richTextValue(root, { node: range.endContainer, offset: range.endOffset }).length,
  };
}

/** Rendered runs contain literal newlines, so selection uses UTF-16 text offsets. */
export function selectRichText(root: HTMLElement, start: number, end = start): void {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const points: { node: Node; offset: number }[] = [];
  let offset = 0;
  let node = walker.nextNode();
  for (const target of [start, end]) {
    while (node && offset + (node.textContent?.length ?? 0) < target) {
      offset += node.textContent?.length ?? 0;
      node = walker.nextNode();
    }
    points.push(
      node
        ? { node, offset: Math.max(0, target - offset) }
        : { node: root, offset: root.childNodes.length },
    );
  }
  const range = document.createRange();
  range.setStart(points[0]!.node, points[0]!.offset);
  range.setEnd(points[1]!.node, points[1]!.offset);
  const selection = document.getSelection();
  selection?.removeAllRanges();
  selection?.addRange(range);
}
