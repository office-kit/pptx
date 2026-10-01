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
  removeSlide,
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

/** Moving title text moves the boundary between adjacent slides' bodies. */
export function outlineTitleMove(pres: PresentationData, slide: SlideData, direction: -1 | 1) {
  const slides = getSlides(pres);
  const index = slides.indexOf(slide);
  if (index <= 0) return null;
  const previousShapes = getSlideShapes(slides[index - 1]!);
  const previousSlot = outlineShapes(slides[index - 1]!)
    .filter((item) => !item.title)
    .at(-1);
  const bodySlot = outlineShapes(slide).find((item) => !item.title);
  if (!previousSlot || !bodySlot) return null;
  const previous = previousShapes.find((shape) => getShapeId(shape) === previousSlot.id)!;
  const body = getSlideShapes(slide).find((shape) => getShapeId(shape) === bodySlot.id)!;
  const before = getShapeText(previous);
  const after = getShapeText(body);
  const source = direction === -1 ? previous : body;
  if (!(direction === -1 ? before : after)) return null;
  const paragraphs = getShapeParagraphElements(source);
  const moved = direction === -1 ? paragraphs.at(-1)! : paragraphs[0]!;
  const length = moved.reduce(
    (sum, element) => sum + (element.kind === 'br' ? 1 : element.text.length),
    0,
  );
  const separator = before && after ? 1 : 0;
  const total = before.length + separator + after.length;
  const boundary = direction === -1 ? before.length - length : before.length + separator + length;
  const ranges =
    direction === -1
      ? [
          { start: 0, end: Math.max(0, boundary - 1) },
          { start: boundary, end: total },
        ]
      : [
          { start: 0, end: boundary },
          { start: Math.min(total, boundary + 1), end: total },
        ];
  return {
    previous,
    body,
    sources: [...(before ? [previous] : []), ...(after ? [body] : [])],
    ranges,
  };
}

export function moveOutlineTitle(pres: PresentationData, slide: SlideData, direction: -1 | 1) {
  const move = outlineTitleMove(pres, slide, direction);
  if (!move) return false;
  // Concatenate before distributing so shared source/target handles retain their XML.
  setShapeParagraphs(move.previous, { sources: move.sources });
  setShapeParagraphs([move.previous, move.body], { source: move.previous, ranges: move.ranges });
  return true;
}

/** Mac PowerPoint demotes a slide title into the preceding slide's body. */
export function demoteOutlineTitle(
  pres: PresentationData,
  slide: SlideData,
): SlideShapeData | null {
  const slides = getSlides(pres);
  const index = slides.indexOf(slide);
  if (index <= 0) return null;
  const previous = slides[index - 1]!;
  const previousById = new Map(getSlideShapes(previous).map((shape) => [getShapeId(shape), shape]));
  const bodySlot = outlineShapes(previous).find((item) => !item.title);
  if (!bodySlot)
    throw new Error('Outline demotion requires a body placeholder on the previous slide');
  const shapes = getSlideShapes(slide);
  const byId = new Map(shapes.map((shape) => [getShapeId(shape), shape]));
  const outline = outlineShapes(slide);
  const ids = new Set(outline.map((item) => item.id));
  if (
    shapes.some(
      (shape) =>
        !ids.has(getShapeId(shape)) &&
        !['dt', 'ftr', 'sldNum'].includes(getShapePlaceholderType(shape) ?? ''),
    )
  )
    throw new Error('Outline demotion of slides with additional objects is not supported');
  const target = previousById.get(bodySlot.id)!;
  const sources = [
    ...(getShapeText(target) ? [target] : []),
    ...outline.filter((item) => item.title).map((item) => byId.get(item.id)!),
    ...outline
      .filter((item) => !item.title && getShapeText(byId.get(item.id)!))
      .map((item) => byId.get(item.id)!),
  ];
  setShapeParagraphs(target, { sources });
  removeSlide(pres, slide);
  return target;
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

/** Move only the selected paragraphs, as the native outline menu does. */
export function outlineParagraphMove(
  source: SlideShapeData,
  range: { start: number; end: number },
  direction: -1 | 1,
) {
  getParagraphLevel(source, range);
  let offset = 0;
  const paragraphs = getShapeParagraphElements(source).map((elements) => {
    const start = offset;
    const end =
      start +
      elements.reduce(
        (length, element) => length + (element.kind === 'br' ? 1 : element.text.length),
        0,
      );
    offset = end + 1;
    return { start, end };
  });
  const first = paragraphs.findIndex((paragraph) => paragraph.end >= range.start);
  let last = first;
  while (last + 1 < paragraphs.length && paragraphs[last + 1]!.start < range.end) last++;
  const adjacent = direction === -1 ? first - 1 : last + 1;
  if (first < 0 || adjacent < 0 || adjacent >= paragraphs.length) return null;
  const moving = paragraphs.slice(first, last + 1);
  const ranges = [...paragraphs];
  ranges.splice(first, moving.length);
  const destination = direction === -1 ? first - 1 : first + 1;
  ranges.splice(destination, 0, ...moving);
  const start =
    direction === -1
      ? paragraphs[adjacent]!.start
      : paragraphs[first]!.start + paragraphs[adjacent]!.end - paragraphs[adjacent]!.start + 1;
  return {
    ranges,
    selection: { start, end: start + moving[moving.length - 1]!.end - moving[0]!.start },
  };
}
