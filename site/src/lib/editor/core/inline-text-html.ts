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
import { copyTextRange } from './text-clipboard.ts';
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
  let offset = 0;
  paragraphs.forEach((elements, index) => {
    if (index) {
      container.append('\n');
      offset++;
    }
    const length = elements.reduce(
      (sum, element) => sum + (element.kind === 'br' ? 1 : element.text.length),
      0,
    );
    const paragraph = document.createElement('section');
    paragraph.setAttribute('data-text-paragraph', '');
    const style = paragraph.style;
    style.display = 'inline-block';
    style.verticalAlign = 'top';
    style.width = '100%';
    style.boxSizing = 'border-box';
    const props = getParagraphPropertiesEffective(pres, target, index, {
      inheritanceSource: source,
    });
    style.textAlign = props.align === 'distribute' ? 'justify' : (props.align ?? 'left');
    if (props.align === 'distribute') style.textAlignLast = 'justify';
    if (props.lineSpacing)
      style.lineHeight =
        props.lineSpacing.kind === 'pct'
          ? String(props.lineSpacing.value)
          : scaled(props.lineSpacing.value, 'pt');
    if (props.spcBefPts !== null) style.marginTop = scaled(props.spcBefPts, 'pt');
    if (props.spcAftPts !== null) style.marginBottom = scaled(props.spcAftPts, 'pt');
    if (props.marL !== null) style.paddingLeft = scaled(props.marL / 9525, 'px');
    if (props.marR !== null) style.paddingRight = scaled(props.marR / 9525, 'px');
    if (props.indent !== null) style.textIndent = scaled(props.indent / 9525, 'px');
    if (props.rtl !== null) style.direction = props.rtl ? 'rtl' : 'ltr';
    const formatted = document.createElement('div');
    // The exporter only emits escaped text and allowlisted styles.
    formatted.innerHTML = textClipboardHtml(
      copyTextRange(shape, offset, offset + length, cell, resolve),
    );
    paragraph.append(...formatted.firstElementChild!.childNodes);
    if (!length) {
      const end = document.createElement('br');
      end.setAttribute('data-caret-end', '');
      paragraph.append(end);
    }
    container.append(paragraph);
    offset += length;
  });
  return container.outerHTML;
}
