import {
  getParagraphPropertiesEffective,
  getShapeParagraphCount,
  getShapeParagraphElements,
  getShapeRunFormatEffective,
  getTableCellParagraphs,
  getTableCellRunFormatEffective,
  getTableCells,
  getEffectiveColorMap,
  getPresentationTheme,
  getShapeSlide,
  type PresentationData,
  type SlideShapeData,
} from '@office-kit/pptx';
import { paragraphNumberLabels } from '@office-kit/pptx-preview';
import { defaultTextMetrics, shapeTextDefaults } from './text-layout-defaults.ts';
import { textClipboardHtml } from './html-text-clipboard.ts';

const themeKeyByToken: Record<string, keyof NonNullable<ReturnType<typeof getPresentationTheme>>> =
  {
    dk1: 'dark1',
    tx1: 'dark1',
    lt1: 'light1',
    bg1: 'light1',
    dk2: 'dark2',
    tx2: 'dark2',
    lt2: 'light2',
    bg2: 'light2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hlink: 'hyperlink',
    folHlink: 'followedHyperlink',
  };

/** Resolve a table run's literal scheme color using the slide's effective map. */
export function resolveEditingTextColor(
  pres: PresentationData,
  shape: SlideShapeData,
  value: string | null | undefined,
): string | undefined {
  if (!value) return undefined;
  const token = value.startsWith('scheme:') ? value.slice('scheme:'.length) : value;
  if (!themeKeyByToken[token]) return value;
  const theme = getPresentationTheme(pres);
  if (!theme) return undefined;
  const mapped = getEffectiveColorMap(getShapeSlide(shape))[token] ?? token;
  const key = themeKeyByToken[mapped] ?? themeKeyByToken[token];
  return key ? theme[key] : undefined;
}

/** Keep literal UTF-16 paragraph separators for editing and clipboard offsets. */
export function inlineTextHtml(
  pres: PresentationData,
  shape: SlideShapeData,
  source: SlideShapeData | undefined,
  cell?: { row: number; col: number },
): string {
  const tableCell = cell ? getTableCells(shape)[cell.row]![cell.col]! : undefined;
  const target = tableCell ?? shape;
  const defaults = defaultTextMetrics(pres, shape);
  const paragraphs = tableCell
    ? getTableCellParagraphs(tableCell).map((p) => p.elements)
    : Array.from({ length: getShapeParagraphCount(shape) }, (_, index) =>
        getShapeParagraphElements(shape, index),
      );
  const container = document.createElement('div');
  const scaled = (value: number, unit: string) => `calc(${value}${unit} * var(--text-zoom))`;
  const resolve = tableCell
    ? (
        paragraph: number,
        run: number | null | { readonly fieldIndex: number } | { readonly breakIndex: number },
      ) => getTableCellRunFormatEffective(pres, tableCell, paragraph, run)
    : (
        paragraph: number,
        run: number | null | { readonly fieldIndex: number } | { readonly breakIndex: number },
      ) => getShapeRunFormatEffective(pres, shape, paragraph, run, { inheritanceSource: source });
  const properties = paragraphs.map((_, index) =>
    getParagraphPropertiesEffective(pres, target, index, { inheritanceSource: source }),
  );
  const labels = paragraphNumberLabels(
    properties.map((p) => ({ bulletStyle: p.bullet, level: p.level })),
  );
  paragraphs.forEach((elements, index) => {
    if (index) {
      container.append('\n');
    }
    let text = '';
    let runIndex = 0;
    let fieldIndex = 0;
    let breakIndex = 0;
    const formats = elements.map((element) => {
      const start = text.length;
      text += element.kind === 'br' ? '\n' : element.text;
      const rawFormat =
        element.kind === 'r'
          ? resolve(index, runIndex++)
          : element.kind === 'fld'
            ? resolve(index, { fieldIndex: fieldIndex++ })
            : resolve(index, { breakIndex: breakIndex++ });
      const editingFormat =
        cell && rawFormat
          ? {
              ...rawFormat,
              ...(rawFormat.color
                ? { color: resolveEditingTextColor(pres, shape, rawFormat.color) }
                : {}),
              ...(rawFormat.underlineColor !== undefined && rawFormat.underlineColor !== null
                ? {
                    underlineColor: resolveEditingTextColor(pres, shape, rawFormat.underlineColor),
                  }
                : {}),
            }
          : rawFormat;
      // The reader widens colors to strings; the HTML exporter takes what a
      // writer would.
      return { start, end: text.length, format: toWritableTextFormat(editingFormat ?? {}) };
    });
    const paragraph = document.createElement('section');
    paragraph.setAttribute('data-text-paragraph', '');
    const style = paragraph.style;
    style.display = 'inline-block';
    style.verticalAlign = 'top';
    style.width = '100%';
    style.boxSizing = 'border-box';
    // Keep the paragraph strut at zero. A large inherited fallback here would
    // enlarge line boxes containing smaller explicitly-sized runs; preview
    // lays out each run at its own effective size instead.
    const emptyFormat = elements.length === 0 ? resolve(index, null) : null;
    style.fontSize = text ? '0px' : scaled(emptyFormat?.size ?? defaults.size, 'pt');
    style.fontFamily = emptyFormat?.font ?? defaults.family;
    if (emptyFormat?.bold !== undefined) style.fontWeight = emptyFormat.bold ? '700' : '400';
    if (emptyFormat?.italic !== undefined)
      style.fontStyle = emptyFormat.italic ? 'italic' : 'normal';
    if (emptyFormat?.color)
      style.color = resolveEditingTextColor(pres, shape, emptyFormat.color) ?? '';
    style.lineHeight = text ? '0' : '1.05';
    const props = properties[index]!;
    style.tabSize = scaled((props.defaultTabSizeEmu ?? 914400) / 9525, 'px');
    if (props.tabStops?.length)
      paragraph.dataset.tabStops = props.tabStops
        .map((stop) => `${stop.positionEmu / 9525}:${stop.alignment}`)
        .join(';');
    const bullet = props.bullet;
    const marker =
      labels[index] ??
      (bullet === 'bullet'
        ? props.level <= 0
          ? '•'
          : props.level === 1
            ? '◦'
            : '▪'
        : bullet && typeof bullet === 'object' && 'char' in bullet
          ? bullet.char
          : null);
    if (marker) paragraph.setAttribute('data-list-marker', marker);
    style.textAlign =
      props.align === 'distribute'
        ? 'justify'
        : (props.align ?? (tableCell ? 'left' : shapeTextDefaults(shape).align));
    if (props.align === 'distribute') style.textAlignLast = 'justify';
    if (props.lineSpacing)
      style.lineHeight =
        props.lineSpacing.kind === 'pct'
          ? String(props.lineSpacing.value)
          : scaled(props.lineSpacing.value, 'pt');
    if (props.spcBefPts !== null) style.marginTop = scaled(props.spcBefPts, 'pt');
    if (props.spcAftPts !== null) style.marginBottom = scaled(props.spcAftPts, 'pt');
    if (props.marL !== null || props.level > 0)
      style.paddingLeft = scaled(props.marL !== null ? props.marL / 9525 : props.level * 32, 'px');
    if (props.marR !== null) style.paddingRight = scaled(props.marR / 9525, 'px');
    if (props.indent !== null) style.textIndent = scaled(props.indent / 9525, 'px');
    if (props.rtl !== null) style.direction = props.rtl ? 'rtl' : 'ltr';
    const formatted = document.createElement('div');
    // The exporter only emits escaped text and allowlisted styles.
    formatted.innerHTML = textClipboardHtml(
      {
        text,
        formats,
      },
      { editing: true },
    );
    for (const span of formatted.querySelectorAll('span')) {
      if (!span.style.fontSize) span.style.fontSize = `${defaults.size}pt`;
      if (!span.style.fontFamily) span.style.fontFamily = defaults.family;
      if (!props.lineSpacing) span.style.lineHeight = '1.05';
    }
    const firstRun = formatted.querySelector('span');
    if (marker && firstRun) {
      if (firstRun.style.fontSize)
        style.setProperty('--marker-size', `calc(${firstRun.style.fontSize} * var(--text-zoom))`);
      if (firstRun.style.fontFamily)
        style.setProperty('--marker-font', `${firstRun.style.fontFamily}, var(--ok-font)`);
      if (firstRun.style.color) style.setProperty('--marker-color', firstRun.style.color);
    }
    paragraph.append(...formatted.firstElementChild!.childNodes);
    if (!text || text.endsWith('\n')) {
      const end = document.createElement('br');
      end.setAttribute('data-caret-end', '');
      paragraph.append(end);
    }
    container.append(paragraph);
  });
  return container.outerHTML;
}
