import {
  getShapeId,
  addSlideAt,
  getShapeText,
  getParagraphLevel,
  getShapeParagraphElements,
  getSlideLayout,
  getSlideLayoutPlaceholders,
  getSlides,
  setParagraphLevel,
  setShapeParagraphs,
  setShapeText,
  type PresentationData,
  type SlideShapeData,
  getShapePlaceholderType,
  getSlideShapes,
  getShapeKind,
  isShapePlaceholder,
  type SlideData,
} from '@office-kit/pptx';

/** PowerPoint's outline excludes ordinary text boxes and footer placeholders. */
export function outlineShapes(slide: SlideData) {
  return getSlideShapes(slide).flatMap((shape) => {
    if (!isShapePlaceholder(shape) || getShapeKind(shape) !== 'shape') return [];
    const type = getShapePlaceholderType(shape) ?? 'obj';
    if (!['title', 'ctrTitle', 'subTitle', 'body', 'obj'].includes(type)) return [];
    return [{ id: getShapeId(shape), title: type === 'title' || type === 'ctrTitle' }];
  });
}

/** Mac PowerPoint promotes root body paragraphs into separate slide titles. */
export function promoteOutlineBody(
  pres: PresentationData,
  slide: SlideData,
  source: SlideShapeData,
  range: { start: number; end: number },
): SlideData[] {
  const text = getShapeText(source);
  const levels = getParagraphLevel(source, { start: 0, end: text.length });
  const elementsByParagraph = getShapeParagraphElements(source);
  if (levels.length < elementsByParagraph.length)
    levels.push(getParagraphLevel(source, elementsByParagraph.length - 1));
  let offset = 0;
  const paragraphs = elementsByParagraph.map((elements, index) => {
    const start = offset;
    const end =
      start +
      elements.reduce(
        (length, element) => length + (element.kind === 'br' ? 1 : element.text.length),
        0,
      );
    offset = end + 1;
    return { start, end, level: levels[index]! };
  });
  // Use the public range reader to validate UTF-16 boundaries before any mutation.
  getParagraphLevel(source, range);
  const selected = paragraphs.filter((paragraph) =>
    range.start === range.end
      ? paragraph.start <= range.start && paragraph.end >= range.start
      : paragraph.start < range.end && paragraph.end >= range.start,
  );
  const roots = selected.filter((paragraph) => paragraph.level === 0);
  if (!roots.length) {
    setParagraphLevel(source, range, { offset: -1 });
    return [];
  }
  const layout = getSlideLayout(slide);
  if (!layout) throw new Error('Outline promotion requires a slide layout');
  const slots = getSlideLayoutPlaceholders(layout);
  if (
    !slots.some((slot) => slot.type === 'title' || slot.type === 'ctrTitle') ||
    !slots.some(
      (slot) =>
        slot.type === null ||
        slot.type === 'obj' ||
        slot.type === 'body' ||
        slot.type === 'subTitle',
    )
  ) {
    throw new Error('Outline promotion requires title and body placeholders');
  }
  const slides = getSlides(pres);
  const added = addSlideAt(
    pres,
    slides.indexOf(slide) + 1,
    roots.map(() => ({ layout })),
  );
  const targets: SlideShapeData[] = [];
  const ranges: Array<{ start: number; end: number }> = [];
  // Selected nested paragraphs move up one level; roots will become titles.
  setParagraphLevel(source, range, { offset: -1 });
  added.forEach((destination, index) => {
    const byId = new Map(getSlideShapes(destination).map((shape) => [getShapeId(shape), shape]));
    const shapes = outlineShapes(destination);
    for (const item of shapes) setShapeText(byId.get(item.id)!, '');
    const title = shapes.find((item) => item.title)!;
    const body = shapes.find((item) => !item.title)!;
    const root = roots[index]!;
    targets.push(byId.get(title.id)!);
    ranges.push({ start: root.start, end: root.end });
    const bodyStart = Math.min(root.end + 1, text.length);
    const bodyEnd = roots[index + 1] ? roots[index + 1]!.start - 1 : text.length;
    if (root.end < text.length && bodyEnd >= bodyStart) {
      targets.push(byId.get(body.id)!);
      ranges.push({ start: bodyStart, end: bodyEnd });
    }
  });
  setShapeParagraphs(targets, { source, ranges });
  const first = roots[0]!;
  setShapeText(source, '', { range: { start: Math.max(0, first.start - 1), end: text.length } });
  return added;
}
