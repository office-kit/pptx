import {
  getSlideLayout,
  getSlideLayoutPlaceholders,
  getSlideLayoutType,
  getSlideLayouts,
  getSlideMasterPartName,
  type PresentationData,
  type SlideData,
  type SlideLayoutData,
} from '@office-kit/pptx';

/**
 * Pick the layout used by PowerPoint's generic New Slide command.
 *
 * PowerPoint keeps the current layout when it is a normal content layout. A
 * title slide is the one exception: it advances to that master's title and
 * content layout so the new slide has an editable body placeholder.
 */
export function newSlideLayout(
  pres: PresentationData,
  current: SlideData | null,
): SlideLayoutData | null {
  const currentLayout = current ? getSlideLayout(current) : null;
  if (!current || !currentLayout) return null;
  if (getSlideLayoutType(currentLayout) !== 'title') return currentLayout;

  const layouts = getSlideLayouts(pres);
  const sameMasterLayouts = layouts.filter(
    (layout) => getSlideMasterPartName(layout) === getSlideMasterPartName(current),
  );
  const objectLayout = sameMasterLayouts.find((layout) => getSlideLayoutType(layout) === 'obj');
  if (objectLayout) return objectLayout;

  const contentLayouts = sameMasterLayouts.filter((layout) => {
    const type = getSlideLayoutType(layout);
    if (type === 'obj') return false;
    return getSlideLayoutPlaceholders(layout).some(
      (placeholder) =>
        placeholder.type === 'body' || (placeholder.type === null && placeholder.idx !== null),
    );
  });
  if (contentLayouts.length === 0) return currentLayout;

  return contentLayouts[0] ?? currentLayout;
}
