import {
  getShapeParagraphElements,
  getShapeRunFormatEffective,
  toWritableTextFormat,
  getParagraphPropertiesEffective,
  type PresentationData,
  type SlideShapeData,
} from '@office-kit/pptx';
import { paragraphNumberLabels } from '@office-kit/pptx-preview';
import { textClipboardHtml } from './html-text-clipboard.ts';

/** Outline indentation stays independent of slide paragraph margins. */
export function outlineTextHtml(
  pres: PresentationData,
  shape: SlideShapeData,
  title: boolean,
  showFormatting = false,
): string {
  const paragraphs = getShapeParagraphElements(shape);
  const properties = paragraphs.map((_, index) =>
    getParagraphPropertiesEffective(pres, shape, index),
  );
  const labels = paragraphNumberLabels(
    properties.map((p) => ({ bulletStyle: p.bullet, level: p.level })),
  );
  const container = document.createElement('div');
  paragraphs.forEach((elements, index) => {
    // Literal newlines let the shared rich-text editor preserve UTF-16 offsets.
    if (index) container.append('\n');
    const paragraph = document.createElement('section');
    paragraph.dataset.outlineParagraph = '';
    paragraph.dataset.textParagraph = '';
    const props = properties[index]!;
    paragraph.style.setProperty('--outline-level', String(props.level));
    if (title) paragraph.dataset.outlineTitle = '';
    else {
      const bullet = props.bullet;
      const marker =
        labels[index] ??
        (bullet === 'bullet'
          ? '•'
          : bullet && typeof bullet === 'object' && 'char' in bullet
            ? bullet.char
            : null);
      if (marker) paragraph.dataset.outlineMarker = marker;
    }
    const text = elements.map((element) => (element.kind === 'br' ? '\n' : element.text)).join('');
    if (showFormatting) {
      let offset = 0;
      let run = 0;
      const formats = elements.map((element) => {
        const start = offset;
        offset += element.kind === 'br' ? 1 : element.text.length;
        const effective =
          element.kind === 'r'
            ? getShapeRunFormatEffective(pres, shape, index, run++)
            : element.format;
        const format = toWritableTextFormat(effective ?? {});
        // Outline text follows the UI foreground, including in dark appearance.
        delete format.color;
        const outlineScale = 0.25;
        if (format.size) format.size *= outlineScale;
        return { start, end: offset, format };
      });
      const formatted = document.createElement('div');
      formatted.innerHTML = textClipboardHtml({ text, formats });
      paragraph.append(...formatted.firstElementChild!.childNodes);
      paragraph.style.lineHeight = 'normal';
    } else paragraph.append(text);
    if (!text || text.endsWith('\n')) {
      const end = document.createElement('br');
      end.dataset.caretEnd = '';
      paragraph.append(end);
    }
    container.append(paragraph);
  });
  return container.outerHTML;
}
