/** Size tab characters without changing the UTF-16 offsets used by editing and clipboard. */
export function layoutEditingTabs(root: HTMLElement, zoom: number): void {
  const paragraphs = root.querySelectorAll<HTMLElement>('[data-tab-stops]');
  if (!paragraphs.length) return;
  const context = document.createElement('canvas').getContext('2d')!;
  for (const paragraph of paragraphs) {
    const paragraphStyle = getComputedStyle(paragraph);
    const vertical = paragraphStyle.writingMode.startsWith('vertical');
    // Upright glyph advances differ from horizontal canvas widths. Measure in
    // an untransformed vertical box so shape rotation and zoom aren't applied twice.
    const verticalMeasure = vertical ? document.createElement('span') : null;
    if (verticalMeasure) {
      verticalMeasure.style.cssText =
        'position:fixed;visibility:hidden;white-space:pre;width:max-content;height:max-content;';
      document.body.append(verticalMeasure);
    }
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
      const variant = style.fontVariantCaps === 'small-caps' ? 'small-caps' : 'normal';
      context.font = `${style.fontStyle} ${variant} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
      context.fontKerning =
        style.fontKerning === 'normal' || style.fontKerning === 'none' ? style.fontKerning : 'auto';
      // Canvas applies tracking to shaped glyphs, including the trailing spacing in CSS layout.
      context.letterSpacing = `${parseFloat(style.letterSpacing) || 0}px`;
      if (verticalMeasure) {
        verticalMeasure.style.font = context.font;
        verticalMeasure.style.fontKerning = style.fontKerning;
        verticalMeasure.style.letterSpacing = style.letterSpacing;
        verticalMeasure.style.writingMode = style.writingMode;
        verticalMeasure.style.textOrientation = style.textOrientation;
        verticalMeasure.style.textTransform = style.textTransform;
      }
      const measure = (text: string) => {
        if (verticalMeasure) {
          verticalMeasure.textContent = text;
          return verticalMeasure.getBoundingClientRect().height;
        }
        // The model retains original case, but tab alignment follows painted glyphs.
        const displayed = style.textTransform === 'uppercase' ? text.toUpperCase() : text;
        return context.measureText(displayed).width;
      };
      const fragment = document.createDocumentFragment();
      for (const text of node.data.split(/(\t|\n)/)) {
        if (text === '\t') {
          const tab = document.createElement('span');
          tab.textContent = text;
          tab.style.display = 'inline-block';
          tab.style.whiteSpace = 'pre';
          tab.style.inlineSize = '0px';
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
    verticalMeasure?.remove();
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
    const margin = parseFloat(vertical ? style.paddingTop : style.paddingLeft) || 0;
    const interval = parseFloat(style.tabSize);
    for (const part of parts) {
      if (!part.tab) continue;
      // Offsets are layout coordinates, so rotated shapes need no screen-space correction.
      const position = (vertical ? part.tab.offsetTop : part.tab.offsetLeft) - margin;
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
      part.tab.style.inlineSize = `${Math.max(0, next - position - shift)}px`;
    }
  }
}
