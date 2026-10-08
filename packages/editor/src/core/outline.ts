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
  getSlidePartName,
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

export type OutlineLocation = { slide: SlideData; id: number; offset: number };

/** Every outline field in display order, keyed as the outline view keys its inputs. */
export function outlineSlots(pres: PresentationData) {
  return getSlides(pres).flatMap((slide, index) =>
    outlineShapes(slide).map((item) => ({
      ...item,
      slide,
      index,
      key: `${getSlidePartName(slide)}:${item.id}`,
    })),
  );
}

function paragraphSpans(shape: SlideShapeData): Array<{ start: number; end: number }> {
  let offset = 0;
  return getShapeParagraphElements(shape).map((elements) => {
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
}

/**
 * Deleting across a slide boundary treats the outline as one text: slides
 * whose titles fall inside the range are removed, the end paragraph's suffix
 * joins the start paragraph (which keeps its title or body level), and the
 * end slide's remaining body moves onto the start slide. Native comparison
 * covers title-to-title; the body-start and body-end cases follow the same
 * text model (see NATIVE_PARITY.md).
 */
export function deleteOutlineSlideRange(
  pres: PresentationData,
  start: OutlineLocation,
  end: OutlineLocation,
): boolean {
  const slides = getSlides(pres);
  const from = slides.indexOf(start.slide);
  const to = slides.indexOf(end.slide);
  const startShapes = outlineShapes(start.slide);
  const endShapes = outlineShapes(end.slide);
  const first = startShapes.findIndex((item) => item.id === start.id);
  const last = endShapes.findIndex((item) => item.id === end.id);
  if (from < 0 || to <= from || first < 0 || last < 0) return false;
  const startsInTitle = startShapes[first]!.title;
  const endsInTitle = endShapes[last]!.title;
  const head = findShapeById(start.slide, start.id)!;
  const tail = findShapeById(end.slide, end.id)!;
  const following = endShapes
    .slice(last + 1)
    .filter((item) => !item.title)
    .map((item) => findShapeById(end.slide, item.id)!);
  const length = getShapeText(head).length;
  setShapeParagraphs(head, { sources: [head, tail] });
  setShapeText(head, '', { range: { start: start.offset, end: length + 1 + end.offset } });
  for (const item of startShapes.slice(first + 1))
    setShapeText(findShapeById(start.slide, item.id)!, '');
  let body = startsInTitle ? null : head;
  if (startsInTitle && !endsInTitle) {
    // The joined paragraph stays in the title; later paragraphs of the end
    // body remain body text, as in deleteOutlineTitleBodyRange.
    const text = getShapeText(head);
    const boundary = paragraphSpans(head).find((span) => span.end >= start.offset)!.end;
    body = outlineBodyTarget(start.slide);
    setShapeParagraphs([head, body], {
      source: head,
      ranges: [
        { start: 0, end: boundary },
        { start: Math.min(boundary + 1, text.length), end: text.length },
      ],
    });
  }
  if (following.length) {
    body ??= outlineBodyTarget(start.slide);
    // The caret paragraph survives even when empty; an emptied body does not.
    const keep = body === head || getShapeText(body) !== '';
    setShapeParagraphs(body, { sources: [...(keep ? [body] : []), ...following] });
  }
  for (const removed of slides.slice(from + 1, to + 1)) removeSlide(pres, removed);
  return true;
}

/**
 * Delete any ordered outline range. Within one slide, separate bodies stay
 * separate shapes; only a title-to-body range joins text across fields.
 */
export function deleteOutlineRange(
  pres: PresentationData,
  start: OutlineLocation,
  end: OutlineLocation,
): boolean {
  if (start.slide !== end.slide) return deleteOutlineSlideRange(pres, start, end);
  const shapes = outlineShapes(start.slide);
  const from = shapes.findIndex((item) => item.id === start.id);
  const to = shapes.findIndex((item) => item.id === end.id);
  if (from < 0 || to < from) return false;
  const head = findShapeById(start.slide, start.id)!;
  if (from === to) {
    setShapeText(head, '', { range: { start: start.offset, end: end.offset } });
    return true;
  }
  if (shapes[from]!.title && !shapes[to]!.title)
    return deleteOutlineTitleBodyRange(start.slide, start, end);
  setShapeText(head, '', { range: { start: start.offset, end: getShapeText(head).length } });
  for (const item of shapes.slice(from + 1, to))
    setShapeText(findShapeById(start.slide, item.id)!, '');
  setShapeText(findShapeById(end.slide, end.id)!, '', { range: { start: 0, end: end.offset } });
  return true;
}

/**
 * Enter in an outline title starts a new slide after it with the same layout.
 * Mac PowerPoint moves the title suffix and the whole following body there.
 */
export function splitOutlineTitle(
  pres: PresentationData,
  slide: SlideData,
  source: SlideShapeData,
  range: { start: number; end: number },
): number | null {
  const layout = getSlideLayout(slide);
  if (!layout) return null;
  const length = getShapeText(source).length;
  const index = getSlides(pres).indexOf(slide) + 1;
  const next = addSlideAt(pres, index, { layout });
  const placeholders = outlineShapes(next);
  for (const item of placeholders) setShapeText(findShapeById(next, item.id)!, '');
  // Copy whole placeholders so paragraph levels, bullets and links survive.
  for (const item of placeholders) {
    if (!item.title) removeShape(findShapeById(next, item.id)!);
  }
  for (const item of outlineShapes(slide)) {
    if (item.title) continue;
    const body = findShapeById(slide, item.id)!;
    copyShape(next, body);
    setShapeText(body, '');
  }
  const heading = placeholders.find((item) => item.title);
  if (heading)
    setShapeParagraphs(findShapeById(next, heading.id)!, {
      source,
      range: { start: range.end, end: length },
    });
  setShapeText(source, '', { range: { start: range.start, end: length } });
  return index;
}

/** A paragraph and the deeper paragraphs after it, which a bullet click selects together. */
export function outlineParagraphBlock(
  shape: SlideShapeData,
  index: number,
): { first: number; last: number } {
  const count = getShapeParagraphElements(shape).length;
  const level = getParagraphLevel(shape, index);
  let last = index;
  while (last + 1 < count && getParagraphLevel(shape, last + 1) > level) last++;
  return { first: index, last };
}

/** UTF-16 range covering whole paragraphs `first` through `last`. */
export function outlineParagraphRange(
  shape: SlideShapeData,
  first: number,
  last: number,
): { start: number; end: number } {
  const spans = paragraphSpans(shape);
  return { start: spans[first]!.start, end: spans[last]!.end };
}

/**
 * Move whole body paragraphs, keeping their XML (levels, bullets, links), to a
 * paragraph boundary of any outline body, then shift their levels. A null
 * target id is the slide's first body, added from the layout when missing.
 * Returns the moved paragraphs' new location, or null when nothing changes.
 */
export function moveOutlineParagraphs(
  source: { slide: SlideData; id: number; first: number; last: number },
  target: { slide: SlideData; id: number | null; index: number },
  levelOffset = 0,
): { shape: SlideShapeData; first: number; last: number } | null {
  const shape = findShapeById(source.slide, source.id)!;
  const spans = paragraphSpans(shape);
  const moving = spans.slice(source.first, source.last + 1);
  const remaining = spans.filter((_, index) => index < source.first || index > source.last);
  const destination =
    target.id === null ? outlineBodyTarget(target.slide) : findShapeById(target.slide, target.id)!;
  let first: number;
  if (destination === shape) {
    if (target.index >= source.first && target.index <= source.last + 1) {
      if (!levelOffset) return null;
      first = source.first;
    } else {
      first = target.index > source.last ? target.index - moving.length : target.index;
      const ranges = [...remaining];
      ranges.splice(first, 0, ...moving);
      setShapeParagraphs(shape, { source: shape, ranges });
    }
  } else {
    const targetSpans = paragraphSpans(destination);
    if (!getShapeText(destination)) {
      // An empty body holds one empty paragraph; the moved paragraphs replace it.
      setShapeParagraphs(destination, { source: shape, ranges: moving });
      first = 0;
    } else {
      // Ranges copy from one source, so append the paragraphs and then reorder.
      setShapeParagraphs(destination, { sources: [destination, shape] });
      const appended = paragraphSpans(destination).slice(targetSpans.length);
      setShapeParagraphs(destination, {
        source: destination,
        ranges: [
          ...targetSpans.slice(0, target.index),
          ...appended.slice(source.first, source.last + 1),
          ...targetSpans.slice(target.index),
        ],
      });
      first = target.index;
    }
    if (remaining.length) setShapeParagraphs(shape, { source: shape, ranges: remaining });
    else setShapeText(shape, '');
  }
  const last = first + moving.length - 1;
  if (levelOffset)
    for (let index = first; index <= last; index++)
      setParagraphLevel(destination, index, { offset: levelOffset });
  return { shape: destination, first, last };
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

function outlineBodyTarget(slide: SlideData): SlideShapeData {
  const bodySlot = outlineShapes(slide).find((item) => !item.title);
  let target = bodySlot
    ? getSlideShapes(slide).find((shape) => getShapeId(shape) === bodySlot.id)
    : null;
  if (!target) {
    const layout = getSlideLayout(slide);
    const slot =
      layout &&
      getSlideLayoutPlaceholders(layout).find((item) =>
        ['body', 'obj', 'subTitle'].includes(item.type ?? 'obj'),
      );
    target = slot
      ? addSlidePlaceholder(
          slide,
          slot.type === 'obj' || slot.type === 'subTitle' ? slot.type : 'body',
        )
      : addSlidePlaceholder(slide, 'body', { source: 'master' });
  }
  if (!target)
    throw new Error('Outline editing requires a body placeholder in the slide layout or master');
  return target;
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
  const target = outlineBodyTarget(previous);
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
