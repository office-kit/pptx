import type { TextFormat } from '@office-kit/pptx';
import type { TextEdit } from './text-edit-preview.ts';

type FormattedText = { text: string; formats: NonNullable<TextEdit['formats']> };
const maxHtmlLength = 4_000_000;
const maxNodes = 50_000;
const maxDepth = 128;
const blocks = new Set(['P', 'DIV', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE']);
const excluded = new Set([
  'SCRIPT',
  'STYLE',
  'HEAD',
  'TITLE',
  'META',
  'LINK',
  'IFRAME',
  'OBJECT',
  'EMBED',
  'SVG',
  'MATH',
  'IMG',
  'TEMPLATE',
]);

/** Parse in an inert template; never attach clipboard markup or load its resources. */
export function parseHtmlTextClipboard(html: string, plain: string): FormattedText | null {
  if (!html || html.length > maxHtmlLength) return null;
  const template = document.createElement('template');
  template.innerHTML = html;
  // Spreadsheet paste has its own cell-aware path.
  if (template.content.querySelector('table')) return null;
  const colorContext = document.createElement('canvas').getContext('2d');
  function color(value: string): string | undefined {
    if (
      !value ||
      !colorContext ||
      !CSS.supports('color', value) ||
      /^(inherit|initial|unset|currentcolor|var\()/i.test(value)
    )
      return undefined;
    colorContext.fillStyle = '#010203';
    colorContext.fillStyle = value;
    const resolved = colorContext.fillStyle;
    return /^#[\da-f]{6}$/i.test(resolved) ? resolved : undefined;
  }
  function formatFor(element: HTMLElement, parent: TextFormat): TextFormat {
    const format = { ...parent };
    const tag = element.tagName;
    if (tag === 'B' || tag === 'STRONG') format.bold = true;
    if (tag === 'I' || tag === 'EM') format.italic = true;
    if (tag === 'U') format.underline = true;
    if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') format.strike = true;
    if (tag === 'SUP') format.baseline = 0.3;
    if (tag === 'SUB') format.baseline = -0.25;
    const style = element.style;
    if (/^(bold|bolder|normal|lighter|[0-9]+)$/.test(style.fontWeight))
      format.bold = /^(bold|bolder)$/.test(style.fontWeight) || Number(style.fontWeight) >= 600;
    if (/^(normal|italic|oblique)/.test(style.fontStyle))
      format.italic = /^(italic|oblique)/.test(style.fontStyle);
    const size = /^(\d+(?:\.\d+)?)(pt|px|em|%)$/.exec(style.fontSize);
    if (size) {
      const n = Number(size[1]);
      const points =
        size[2] === 'pt'
          ? n
          : size[2] === 'px'
            ? n * 0.75
            : parent.size
              ? (parent.size * n) / (size[2] === '%' ? 100 : 1)
              : undefined;
      if (points && points <= 4000) format.size = points;
    }
    const family = (style.fontFamily || element.getAttribute('face') || '')
      .split(',')[0]!
      .trim()
      .replace(/^['"]|['"]$/g, '');
    if (family && family.length <= 256 && !/^(inherit|initial|unset|revert)$/.test(family)) {
      format.font = family;
      format.fontEastAsian = family;
    }
    const foreground = color(style.color || element.getAttribute('color') || '');
    if (foreground) format.color = foreground;
    const background = color(style.backgroundColor);
    if (background) format.highlight = background;
    const decoration = style.textDecorationLine || style.textDecoration;
    if (decoration && !/^(inherit|initial|unset|revert)$/.test(decoration)) {
      format.underline = decoration.includes('underline');
      format.strike = decoration.includes('line-through');
    }
    if (style.verticalAlign === 'super') format.baseline = 0.3;
    if (style.verticalAlign === 'sub') format.baseline = -0.25;
    if (style.verticalAlign === 'baseline') format.baseline = 0;
    if (/^-?\d+(?:\.\d+)?%$/.test(style.verticalAlign))
      format.baseline = Number.parseFloat(style.verticalAlign) / 100;
    return format;
  }
  const formats: FormattedText['formats'] = [];
  let text = '';
  function append(value: string, format: TextFormat) {
    if (!value) return;
    const start = text.length;
    text += value;
    formats.push({ start, end: text.length, format });
  }
  type Visit = {
    node: Node;
    format: TextFormat;
    preserve: boolean;
    depth: number;
    exit?: boolean;
    start?: number;
  };
  const stack: Visit[] = [{ node: template.content, format: {}, preserve: false, depth: 0 }];
  let visited = 0;
  while (stack.length) {
    const item = stack.pop()!;
    if (++visited > maxNodes || item.depth > maxDepth) return null;
    if (item.exit) {
      if (text.length === item.start || !text.endsWith('\n')) append('\n', item.format);
      continue;
    }
    const { node } = item;
    if (node.nodeType === Node.TEXT_NODE) {
      let value = node.textContent ?? '';
      if (!item.preserve) {
        value = value.replace(/[\t\r\n ]+/g, ' ');
        if (!text || /[\n ]$/.test(text)) value = value.replace(/^ /, '');
      }
      append(value, item.format);
      continue;
    }
    let { format, preserve } = item;
    if (node instanceof HTMLElement) {
      if (excluded.has(node.tagName) || node.hidden || node.style.display === 'none') continue;
      if (node.tagName === 'BR') {
        append('\n', format);
        continue;
      }
      format = formatFor(node, format);
      preserve =
        node.tagName === 'PRE' ||
        /^(pre|pre-wrap|break-spaces)$/.test(node.style.whiteSpace) ||
        (preserve && !node.style.whiteSpace);
      if (blocks.has(node.tagName)) {
        if (text && !text.endsWith('\n')) append('\n', format);
        stack.push({ ...item, format, exit: true, start: text.length });
      }
    } else if (node.nodeType !== Node.DOCUMENT_FRAGMENT_NODE) continue;
    for (let i = node.childNodes.length - 1; i >= 0; i--)
      stack.push({ node: node.childNodes[i]!, format, preserve, depth: item.depth + 1 });
  }
  // Plain text is authoritative: unfamiliar HTML layout must not change the copied words.
  const normalizedPlain = plain.replace(/\r\n?/g, '\n');
  if (text.endsWith('\n') && !normalizedPlain.endsWith('\n')) text = text.slice(0, -1);
  if (text !== normalizedPlain && text.replace(/\u00a0/g, ' ') !== normalizedPlain) return null;
  const clipped = formats
    .filter((span) => span.start < text.length)
    .map((span) => ({ ...span, end: Math.min(span.end, text.length) }));
  return { text: normalizedPlain, formats: clipped };
}

export function textClipboardHtml(copied: FormattedText): string {
  const container = document.createElement('div');
  container.style.whiteSpace = 'pre-wrap';
  const cssColor = (value: string) => (/^[\da-f]{6}$/i.test(value) ? `#${value}` : value);
  for (const { start, end, format } of copied.formats) {
    const span = document.createElement('span');
    span.textContent = copied.text.slice(start, end);
    const style = span.style;
    if (format.bold !== undefined) style.fontWeight = format.bold ? 'bold' : 'normal';
    if (format.italic !== undefined) style.fontStyle = format.italic ? 'italic' : 'normal';
    if (format.size !== undefined) style.fontSize = `${format.size}pt`;
    const families = [format.font, format.fontEastAsian].filter((font): font is string => !!font);
    if (families.length) style.fontFamily = families.map((font) => JSON.stringify(font)).join(', ');
    if (format.color) style.color = cssColor(format.color);
    if (format.highlight) style.backgroundColor = cssColor(format.highlight);
    const decorations = [];
    if (format.underline && format.underline !== 'none') decorations.push('underline');
    if (format.strike && format.strike !== 'noStrike') decorations.push('line-through');
    if (decorations.length) style.textDecorationLine = decorations.join(' ');
    if (format.baseline) style.verticalAlign = `${format.baseline * 100}%`;
    container.append(span);
  }
  return container.outerHTML;
}
