/** Size tab characters without changing the UTF-16 offsets used by editing and clipboard. */
export function layoutEditingTabs(root: HTMLElement, zoom: number): void {
  const paragraphs = root.querySelectorAll<HTMLElement>('[data-tab-stops]');
  if (!paragraphs.length) return;
  const context = document.createElement('canvas').getContext('2d')!;
  for (const paragraph of paragraphs) {
    const stops = paragraph.dataset.tabStops!.split(';').map((entry) => {
      const [position, alignment] = entry.split(':');
      return { position: Number(position) * zoom, alignment };
    });
    paragraph.style.position = 'relative';
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    const nodes: Text[] = [];
    let node;
    while ((node = walker.nextNode())) nodes.push(node as Text);
    const parts: { text: string; width: number; decimal: number; tab?: HTMLElement }[] = [];
    for (const node of nodes) {
      const style = getComputedStyle(node.parentElement!);
      context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
      const spacing = parseFloat(style.letterSpacing) || 0;
      const measure = (text: string) =>
        context.measureText(text).width + Math.max(0, [...text].length - 1) * spacing;
      const fragment = document.createDocumentFragment();
      for (const text of node.data.split(/(\t|\n)/)) {
        if (text === '\t') {
          const tab = document.createElement('span');
          tab.textContent = text;
          tab.style.display = 'inline-block';
          tab.style.whiteSpace = 'pre';
          tab.style.width = '0px';
          tab.style.tabSize = '0';
          fragment.append(tab);
          parts.push({ text, width: 0, decimal: 0, tab });
        } else {
          fragment.append(text);
          const decimal = text.indexOf('.');
          parts.push({
            text,
            width: measure(text),
            decimal: decimal < 0 ? -1 : measure(text.slice(0, decimal)),
          });
        }
      }
      node.replaceWith(fragment);
    }
    let fieldWidth = 0;
    let decimalWidth = 0;
    const fields = new Map<HTMLElement, { width: number; decimal: number }>();
    for (let index = parts.length - 1; index >= 0; index--) {
      const part = parts[index]!;
      if (part.tab || part.text === '\n') {
        if (part.tab) fields.set(part.tab, { width: fieldWidth, decimal: decimalWidth });
        fieldWidth = decimalWidth = 0;
      } else {
        fieldWidth += part.width;
        decimalWidth = part.decimal < 0 ? part.width + decimalWidth : part.decimal;
      }
    }
    const style = getComputedStyle(paragraph);
    const margin = parseFloat(style.paddingLeft) || 0;
    const interval = parseFloat(style.tabSize);
    for (const part of parts) {
      if (!part.tab) continue;
      // offsetLeft is in layout coordinates, so rotated shapes need no screen-space correction.
      const position = part.tab.offsetLeft - margin;
      let low = 0;
      let high = stops.length;
      while (low < high) {
        const middle = (low + high) >>> 1;
        if (stops[middle]!.position <= position + 0.01) low = middle + 1;
        else high = middle;
      }
      const stop = stops[low];
      const next =
        stop?.position ??
        (interval > 0 ? (Math.floor(position / interval) + 1) * interval : position);
      const field = fields.get(part.tab)!;
      const shift =
        stop?.alignment === 'center'
          ? field.width / 2
          : stop?.alignment === 'right'
            ? field.width
            : stop?.alignment === 'decimal'
              ? field.decimal
              : 0;
      part.tab.style.width = `${Math.max(0, next - position - shift)}px`;
    }
  }
}
