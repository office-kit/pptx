import {
  getShapeParagraphElements,
  getParagraphPropertiesEffective,
  type PresentationData,
  type SlideShapeData,
} from '@office-kit/pptx';
import { paragraphNumberLabels } from '@office-kit/pptx-preview';

/** Outline paragraphs use a fixed editing size and a 10px step per level. */
export function outlineTextHtml(
  pres: PresentationData,
  shape: SlideShapeData,
  title: boolean,
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
    paragraph.append(text);
    if (!text || text.endsWith('\n')) {
      const end = document.createElement('br');
      end.dataset.caretEnd = '';
      paragraph.append(end);
    }
    container.append(paragraph);
  });
  return container.outerHTML;
}
