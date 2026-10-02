import type { Color, TextFormat } from '@office-kit/pptx';
import { textUnderlineStyle } from '@office-kit/pptx-preview';
import type { TextEdit } from './text-edit-preview.ts';

type FormattedText = { text: string; formats: NonNullable<TextEdit['formats']> };
type CapsState = {
  textTransform?: 'uppercase' | 'none';
  fontVariantCaps?: 'small-caps' | 'normal';
};
const cssUnderlineStyles = {
  double: 'dbl',
  dotted: 'dotted',
  dashed: 'dash',
  wavy: 'wavy',
} as const;
type CssUnderlineStyle = keyof typeof cssUnderlineStyles;
const maxHtmlLength = 4_000_000;
const maxNodes = 50_000;
const maxDepth = 128;
// Match the preview's script-size approximation; authored sizes stay in the model.
const scriptSizeRatio = 0.65;
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
  function color(value: string): Color | undefined {
    if (
      !value ||
      !colorContext ||
      !CSS.supports('color', value) ||
      /^(inherit|initial|unset|currentcolor|var\()/i.test(value)
    )
      return undefined;
    colorContext.fillStyle = '#010203';
    colorContext.fillStyle = value;
    // The template literal preserves the checked HexColor type without a cast.
    const hex = /^#([\da-f]{6})$/i.exec(colorContext.fillStyle);
    return hex ? `#${hex[1]!}` : undefined;
  }
  function formatFor(
    element: HTMLElement,
    parent: TextFormat,
    inheritedCaps: CapsState,
  ): { format: TextFormat; caps: CapsState } {
    const format = { ...parent };
    const caps = { ...inheritedCaps };
    const tag = element.tagName;
    const style = element.style;
    if (tag === 'B' || tag === 'STRONG') format.bold = true;
    if (tag === 'I' || tag === 'EM') format.italic = true;
    if (tag === 'SUP') format.baseline = 0.3;
    if (tag === 'SUB') format.baseline = -0.25;
    // PowerPoint's capitalization is represented by one DrawingML `cap`
    // value, while CSS exposes two properties. Resolve their representable
    // combination without letting `font-variant-caps: normal` erase an
    // explicit `text-transform: uppercase`.
    if (style.textTransform && style.textTransform !== 'inherit') {
      caps.textTransform = style.textTransform.toLowerCase().includes('uppercase')
        ? 'uppercase'
        : 'none';
    }
    if (style.fontVariantCaps && style.fontVariantCaps !== 'inherit') {
      caps.fontVariantCaps = style.fontVariantCaps.toLowerCase().includes('small-caps')
        ? 'small-caps'
        : 'normal';
    }
    if (caps.textTransform === 'uppercase') format.cap = 'all';
    else if (caps.fontVariantCaps === 'small-caps') format.cap = 'small';
    else if (caps.textTransform === 'none' || caps.fontVariantCaps === 'normal')
      format.cap = 'none';
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
      if (points !== undefined && points >= 1 && points <= 4000) format.size = points;
    }
    const letterSpacing = /^(-?\d+(?:\.\d+)?)(pt|px|em)$/.exec(style.letterSpacing);
    if (letterSpacing) {
      const n = Number(letterSpacing[1]);
      // OOXML `spc` is hundredths of a point; CSS uses 96 pixels per inch and 72 points per inch.
      // Resolve em against this element's size, including a relative font-size.
      const points =
        letterSpacing[2] === 'pt'
          ? n
          : letterSpacing[2] === 'px'
            ? n * 0.75
            : format.size !== undefined
              ? n * format.size
              : undefined;
      if (points !== undefined) {
        const spc = points * 100;
        if (Number.isFinite(spc) && Math.abs(spc) <= 400_000) format.spc = Math.round(spc);
      }
    } else if (style.letterSpacing === 'normal') {
      format.spc = 0;
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
    // Explicit `none` overrides the tag default; otherwise semantic tags still
    // use authored decoration style/color even without a decoration-line value.
    const semanticDecoration =
      tag === 'U'
        ? 'underline'
        : tag === 'S' || tag === 'STRIKE' || tag === 'DEL'
          ? 'line-through'
          : '';
    const decoration = style.textDecorationLine || style.textDecoration || semanticDecoration;
    if (decoration && !/^(inherit|initial|unset|revert)$/.test(decoration)) {
      if (decoration.includes('underline')) {
        const decorationStyle = style.textDecorationStyle.toLowerCase() as CssUnderlineStyle;
        format.underline = cssUnderlineStyles[decorationStyle] ?? true;
        const decorationColor = color(style.textDecorationColor);
        if (decorationColor) format.underlineColor = decorationColor;
      }
      if (decoration.includes('line-through'))
        format.strike = style.textDecorationStyle === 'double' ? 'dblStrike' : true;
    }
    if (style.verticalAlign === 'super') format.baseline = 0.3;
    if (style.verticalAlign === 'sub') format.baseline = -0.25;
    if (style.verticalAlign === 'baseline') format.baseline = 0;
    if (/^-?\d+(?:\.\d+)?%$/.test(style.verticalAlign))
      format.baseline = Number.parseFloat(style.verticalAlign) / 100;
    return { format, caps };
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
    caps: CapsState;
    preserve: boolean;
    depth: number;
    exit?: boolean;
    start?: number;
  };
  const stack: Visit[] = [
    { node: template.content, format: {}, caps: {}, preserve: false, depth: 0 },
  ];
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
    let { format, caps, preserve } = item;
    if (node instanceof HTMLElement) {
      if (excluded.has(node.tagName) || node.hidden || node.style.display === 'none') continue;
      if (node.tagName === 'BR') {
        append('\n', format);
        continue;
      }
      ({ format, caps } = formatFor(node, format, caps));
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
      stack.push({ node: node.childNodes[i]!, format, caps, preserve, depth: item.depth + 1 });
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

export type TextClipboardHtmlOptions = { editing?: boolean };

export function textClipboardHtml(
  copied: FormattedText,
  options: TextClipboardHtmlOptions = {},
): string {
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
    if (format.spc !== undefined) style.letterSpacing = `${(format.spc / 100) * (96 / 72)}px`;
    const families = [format.font, format.fontEastAsian].filter((font): font is string => !!font);
    if (families.length) style.fontFamily = families.map((font) => JSON.stringify(font)).join(', ');
    if (format.color) style.color = cssColor(format.color);
    if (format.highlight) style.backgroundColor = cssColor(format.highlight);
    if (format.cap === 'all') style.textTransform = 'uppercase';
    else if (format.cap === 'small') style.fontVariantCaps = 'small-caps';
    else if (format.cap === 'none') {
      style.textTransform = 'none';
      style.fontVariantCaps = 'normal';
    }
    if (format.kern !== undefined && format.size !== undefined)
      style.fontKerning = format.kern > 0 && format.size >= format.kern / 100 ? 'normal' : 'none';
    const decorations = [];
    const underline = format.underline;
    const strikeStyle = format.strike === 'dblStrike' ? 'double' : 'solid';
    const editing = options.editing === true;
    if (format.baseline) {
      if (editing && format.size !== undefined) {
        style.fontSize = `${format.size * scriptSizeRatio}pt`;
        // em follows canvas zoom and uses the reduced font size, whereas CSS
        // percentages use line height rather than OOXML's authored font size.
        style.verticalAlign = `${format.baseline / scriptSizeRatio}em`;
      } else style.verticalAlign = `${format.baseline * 100}%`;
    }
    const explicitUnderlineColor =
      format.underlineColor !== undefined && format.underlineColor !== null;
    const underlineColor = cssColor(format.underlineColor ?? format.color ?? '#000000');
    if (editing && underline && underline !== 'none') {
      const strike = !!format.strike && format.strike !== 'noStrike';
      style.textDecorationLine = strike ? 'line-through' : 'none';
      const underlineStyle = textUnderlineStyle(underline, underlineColor, explicitUnderlineColor);
      const words = underline === 'words' ? span.textContent!.split(/(\s+)/) : [span.textContent!];
      const underlineNodes = document.createDocumentFragment();
      for (const word of words) {
        if (!word) continue;
        if (underline === 'words' && /^\s+$/.test(word))
          underlineNodes.append(document.createTextNode(word));
        else {
          const underlineElement = document.createElement('u');
          underlineElement.textContent = word;
          underlineElement.style.cssText += underlineStyle;
          if (explicitUnderlineColor) underlineElement.style.textDecorationColor = underlineColor;
          underlineNodes.append(underlineElement);
        }
      }
      span.replaceChildren(underlineNodes);
      if (strike) style.textDecorationStyle = strikeStyle;
      container.append(span);
      continue;
    }
    const patternedUnderline =
      underline && underline !== 'none' && underline !== true && underline !== 'sng';
    const underlineStyle =
      underline === 'dbl'
        ? 'double'
        : underline === 'dotted' || underline === 'dottedHeavy'
          ? 'dotted'
          : underline === 'dash' || underline === 'dashHeavy'
            ? 'dashed'
            : underline === 'wavy' || underline === 'wavyHeavy'
              ? 'wavy'
              : undefined;
    const separateUnderline =
      format.strike &&
      format.strike !== 'noStrike' &&
      (explicitUnderlineColor || format.strike === 'dblStrike') &&
      underline !== undefined &&
      underline !== false &&
      underline !== 'none';
    if (
      (patternedUnderline || separateUnderline) &&
      format.strike &&
      format.strike !== 'noStrike'
    ) {
      // CSS applies text-decoration-style to every line on the element. Keep
      // underline and strike styles on separate inline boxes.
      style.textDecorationLine = 'line-through';
      style.textDecorationStyle = strikeStyle;
      const underlineSpan = document.createElement('u');
      underlineSpan.textContent = span.textContent;
      underlineSpan.style.textDecorationLine = 'underline';
      if (underlineStyle) underlineSpan.style.textDecorationStyle = underlineStyle;
      if (explicitUnderlineColor) underlineSpan.style.textDecorationColor = underlineColor;
      span.replaceChildren(underlineSpan);
    } else {
      if (underline && underline !== 'none') decorations.push('underline');
      if (format.strike && format.strike !== 'noStrike') decorations.push('line-through');
      if (decorations.length) style.textDecorationLine = decorations.join(' ');
      if (underlineStyle) style.textDecorationStyle = underlineStyle;
      else if (format.strike === 'dblStrike') style.textDecorationStyle = strikeStyle;
      if (explicitUnderlineColor && underline && underline !== 'none')
        style.textDecorationColor = underlineColor;
    }
    container.append(span);
  }
  return container.outerHTML;
}
