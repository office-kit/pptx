import {
  getParagraphPropertiesEffective,
  getShapeParagraphCount,
  getShapeParagraphElements,
  getShapeRunFormatEffective,
  getTableCellParagraphs,
  getTableCells,
  type PresentationData,
  type SlideShapeData,
} from '@office-kit/pptx';
import { paragraphNumberLabels } from '@office-kit/pptx-preview';
import { textClipboardHtml } from './html-text-clipboard.ts';

/** Keep literal UTF-16 paragraph separators for editing and clipboard offsets. */
export function inlineTextHtml(
  pres: PresentationData,
  shape: SlideShapeData,
  source: SlideShapeData | undefined,
  cell?: { row: number; col: number },
): string {
  const tableCell = cell ? getTableCells(shape)[cell.row]![cell.col]! : undefined;
  const target = tableCell ?? shape;
  const paragraphs = tableCell
    ? getTableCellParagraphs(tableCell).map((p) => p.elements)
    : Array.from({ length: getShapeParagraphCount(shape) }, (_, index) =>
        getShapeParagraphElements(shape, index),
      );
  const container = document.createElement('div');
  const scaled = (value: number, unit: string) => `calc(${value}${unit} * var(--text-zoom))`;
  const resolve = cell
    ? undefined
    : (paragraph: number, run: number) =>
        getShapeRunFormatEffective(pres, shape, paragraph, run, { inheritanceSource: source });
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
    const formats = elements.map((element) => {
      const start = text.length;
      text += element.kind === 'br' ? '\n' : element.text;
      const format =
        element.kind === 'r' ? (resolve?.(index, runIndex++) ?? element.format) : element.format;
      return { start, end: text.length, format: format ?? {} };
    });
    const paragraph = document.createElement('section');
    paragraph.setAttribute('data-text-paragraph', '');
    const style = paragraph.style;
    style.display = 'inline-block';
    style.verticalAlign = 'top';
    style.width = '100%';
    style.boxSizing = 'border-box';
    const props = properties[index]!;
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
    style.textAlign = props.align === 'distribute' ? 'justify' : (props.align ?? 'left');
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
    formatted.innerHTML = textClipboardHtml({ text, formats });
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
