export type TextSelection = { start: number; end: number };

/** Contenteditable can create DIV/BR nodes while typing; count them as text newlines. */
export function richTextValue(root: HTMLElement, stop?: { node: Node; offset: number }): string {
  // Chromium leaves one BR to hold the caret after deleting all content.
  // It is an empty editing host, not an authored line break.
  if (root.childNodes.length === 1 && root.firstChild instanceof HTMLBRElement) return '';
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

/** Resolve an editor UTF-16 offset using the same block and BR rules as richTextValue. */
export function richTextPoint(root: HTMLElement, target: number): { node: Node; offset: number } {
  // A lone BR is Chromium's empty-host caret placeholder, so it contributes no
  // value and the canonical empty position is before that placeholder.
  if (root.childNodes.length === 1 && root.firstChild instanceof HTMLBRElement) {
    return { node: root, offset: 0 };
  }
  let candidate: { node: Node; offset: number } = { node: root, offset: root.childNodes.length };
  let total = 0;
  let last = '';
  const visit = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const length = node.textContent?.length ?? 0;
      if (total + length >= target) {
        candidate = { node, offset: Math.max(0, target - total) };
        return true;
      }
      total += length;
      last = node.textContent?.at(-1) ?? last;
      return false;
    }
    if (node instanceof HTMLElement && node.tagName === 'BR') {
      if (!node.hasAttribute('data-caret-end')) {
        const parent = node.parentNode ?? root;
        if (target === total) {
          const index = [...parent.childNodes].indexOf(node);
          candidate = { node: parent, offset: index };
          return true;
        }
        if (target === total + 1) {
          const index = [...parent.childNodes].indexOf(node);
          candidate = { node: parent, offset: index + 1 };
          return true;
        }
        total++;
        last = '\n';
      }
      return false;
    }
    const block = node !== root && node instanceof HTMLElement && /^(DIV|P)$/.test(node.tagName);
    if (block && total && last !== '\n') {
      const parent = node.parentNode ?? root;
      if (target === total) {
        const index = [...parent.childNodes].indexOf(node);
        candidate = { node: parent, offset: index };
        return true;
      }
      if (target === total + 1) {
        candidate = { node, offset: 0 };
        return true;
      }
      total++;
      last = '\n';
    }
    for (const child of [...node.childNodes]) if (visit(child)) return true;
    if (block && node.nextSibling && last !== '\n') {
      if (target === total) {
        candidate = { node, offset: node.childNodes.length };
        return true;
      }
      if (target === total + 1) {
        const parent = node.parentNode ?? root;
        candidate = { node: parent, offset: [...parent.childNodes].indexOf(node) + 1 };
        return true;
      }
      total++;
      last = '\n';
    }
    if (target === total) {
      candidate = { node, offset: node.childNodes.length };
      return true;
    }
    return false;
  };
  visit(root);
  return candidate;
}

/** Restore a selection using the same UTF-16 offsets as text extraction. */
export function selectRichText(root: HTMLElement, start: number, end = start): void {
  const from = richTextPoint(root, start);
  const to = richTextPoint(root, end);
  const range = root.ownerDocument.createRange();
  range.setStart(from.node, from.offset);
  range.setEnd(to.node, to.offset);
  const selection = root.ownerDocument.getSelection();
  selection?.removeAllRanges();
  selection?.addRange(range);
}
