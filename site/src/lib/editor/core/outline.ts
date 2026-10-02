import {
  getShapeId,
  addSlideAt,
  addSlidePlaceholder,
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
  findShapeById,
  copyShape,
  removeShape,
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

/** Mac PowerPoint joins the final body paragraph suffix onto the surviving title. */
export function deleteOutlineTitleBodyRange(
  slide: SlideData,
  start: { id: number; offset: number },
  end: { id: number; offset: number },
): boolean {
  const shapes = outlineShapes(slide);
  const from = shapes.findIndex((item) => item.id === start.id && item.title);
  const to = shapes.findIndex((item) => item.id === end.id && !item.title);
  if (from < 0 || to <= from) return false;
  const title = findShapeById(slide, start.id)!;
  const body = findShapeById(slide, end.id)!;
  const titleLength = getShapeText(title).length;
  setShapeParagraphs(title, { sources: [title, body] });
  setShapeText(title, '', { range: { start: start.offset, end: titleLength + 1 + end.offset } });
  const text = getShapeText(title);
  // PowerPoint keeps all title paragraphs through the paragraph where deletion starts.
  let boundary = 0;
  for (const paragraph of getShapeParagraphElements(title)) {
    boundary += paragraph.reduce(
      (length, element) => length + (element.kind === 'br' ? 1 : element.text.length),
      0,
    );
    if (boundary >= start.offset) break;
    boundary += 1;
  }
  setShapeParagraphs([title, body], {
    source: title,
    ranges: [
      { start: 0, end: boundary },
      { start: Math.min(boundary + 1, text.length), end: text.length },
    ],
  });
  for (const item of shapes.slice(from + 1, to)) setShapeText(findShapeById(slide, item.id)!, '');
  return true;
}

/** Preserve PowerPoint's slide boundary when replacing an outline title range. */
export function splitOutlineTitleRange(
  pres: PresentationData,
  slide: SlideData,
  start: { id: number; offset: number },
  end: { id: number; offset: number; slide?: SlideData },
): number | null {
  const endSlide = end.slide ?? slide;
  const shapes = outlineShapes(slide);
  if (endSlide !== slide) {
    const slides = getSlides(pres);
    const index = slides.indexOf(slide) + 1;
    if (
      slides[index] !== endSlide ||
      !shapes.some((item) => item.id === start.id && item.title) ||
      !outlineShapes(endSlide).some((item) => item.id === end.id && item.title)
    )
      return null;
    // Enter across adjacent titles retains both slides and their title boundary.
    const title = findShapeById(slide, start.id)!;
    setShapeText(title, '', { range: { start: start.offset, end: getShapeText(title).length } });
    for (const item of shapes) {
      if (!item.title) setShapeText(findShapeById(slide, item.id)!, '');
    }
    setShapeText(findShapeById(endSlide, end.id)!, '', { range: { start: 0, end: end.offset } });
    return index;
  }
  const layout = getSlideLayout(slide);
  if (!layout) return null;
  const from = shapes.findIndex((item) => item.id === start.id && item.title);
  const to = shapes.findIndex((item) => item.id === end.id && !item.title);
  if (from < 0 || to <= from) return null;
  const source = findShapeById(slide, end.id)!;
  const text = getShapeText(source);
  // PowerPoint promotes the unselected end paragraph into the new title.
  // Subsequent body paragraphs stay body paragraphs on the new slide.
  const paragraphs = getShapeParagraphElements(source);
  let boundary = 0;
  for (const paragraph of paragraphs) {
    boundary += paragraph.reduce(
      (length, element) => length + (element.kind === 'br' ? 1 : element.text.length),
      0,
    );
    if (boundary >= end.offset) break;
    boundary++;
  }
  const index = getSlides(pres).indexOf(slide) + 1;
  const next = addSlideAt(pres, index, { layout });
  for (const item of outlineShapes(next)) {
    const shape = findShapeById(next, item.id)!;
    if (item.title)
      setShapeParagraphs(shape, { source, range: { start: end.offset, end: boundary } });
    else removeShape(shape);
  }
  for (const [position, item] of shapes.entries()) {
    if (item.title) continue;
    const body = findShapeById(slide, item.id)!;
    if (position > to) copyShape(next, body);
    else if (position === to && boundary < text.length) {
      const copied = copyShape(next, body);
      setShapeParagraphs(copied, {
        source: body,
        range: { start: boundary + 1, end: text.length },
      });
    }
    setShapeText(body, '');
  }
  const title = findShapeById(slide, start.id)!;
  setShapeText(title, '', { range: { start: start.offset, end: getShapeText(title).length } });
  return index;
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

/** Extra objects are discarded by title demotion after PowerPoint's confirmation. */
export function outlineDemotionNeedsConfirmation(slide: SlideData): boolean {
  const ids = new Set(outlineShapes(slide).map((item) => item.id));
  return getSlideShapes(slide).some(
    (shape) =>
      !ids.has(getShapeId(shape)) &&
      !['dt', 'ftr', 'sldNum'].includes(getShapePlaceholderType(shape) ?? ''),
  );
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
  const bodySlot = outlineShapes(previous).find((item) => !item.title);
  let target = bodySlot
    ? getSlideShapes(previous).find((shape) => getShapeId(shape) === bodySlot.id)
    : null;
  if (!target) {
    const layout = getSlideLayout(previous);
    const slot =
      layout &&
      getSlideLayoutPlaceholders(layout).find((item) =>
        ['body', 'obj', 'subTitle'].includes(item.type ?? 'obj'),
      );
    target = slot
      ? addSlidePlaceholder(
          previous,
          slot.type === 'obj' || slot.type === 'subTitle' ? slot.type : 'body',
        )
      : addSlidePlaceholder(previous, 'body', { source: 'master' });
  }
  if (!target)
    throw new Error('Outline demotion requires a body placeholder in the slide layout or master');
  const shapes = getSlideShapes(slide);
  const byId = new Map(shapes.map((shape) => [getShapeId(shape), shape]));
  const outline = outlineShapes(slide);
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
