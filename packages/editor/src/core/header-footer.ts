// Insert ▸ Header & Footer: the date, slide number and footer the reference desktop app keeps
// in the `dt`, `sldNum` and `ftr` placeholders. A layout without the slot
// borrows the master's, as the reference desktop app does; a deck with neither (the DSL's
// default template) gets a plain box along the bottom edge, named so it is
// found again, as `slide-numbers.ts` does for the number.

import {
  addSlidePlaceholder,
  addSlideTextBox,
  emu,
  findSlidePlaceholder,
  getShapeName,
  getSlideShapes,
  getSlideSize,
  inches,
  setParagraphAlignment,
  setShapeTextFormat,
  getShapeParagraphElements,
  getShapeText,
  getSlideLayout,
  getSlideLayoutType,
  removeShape,
  setShapeText,
  setShapeTextField,
  type PresentationData,
  type SlideData,
} from '@office-kit/pptx';
import { hideSlideNumber, showSlideNumber, slideNumberShape } from './slide-numbers.ts';

export type HeaderFooterDate =
  | { readonly kind: 'auto'; readonly format: 'datetime1' | 'datetime2' }
  | { readonly kind: 'fixed'; readonly text: string };

export interface HeaderFooter {
  readonly date: HeaderFooterDate | null;
  readonly slideNumber: boolean;
  readonly footer: string | null;
  readonly hideOnTitle: boolean;
}

export const isTitleSlide = (slide: SlideData): boolean => {
  const layout = getSlideLayout(slide);
  return layout !== null && getSlideLayoutType(layout) === 'title';
};

export function readHeaderFooter(slide: SlideData): Omit<HeaderFooter, 'hideOnTitle'> {
  const dt = find(slide, 'dt');
  const field = dt
    ? getShapeParagraphElements(dt, 0).find((element) => element.kind === 'fld')
    : undefined;
  const date: HeaderFooterDate | null = !dt
    ? null
    : field?.kind === 'fld' && field.type?.startsWith('datetime')
      ? { kind: 'auto', format: field.type === 'datetime2' ? 'datetime2' : 'datetime1' }
      : { kind: 'fixed', text: getShapeText(dt) };
  const footer = find(slide, 'ftr');
  return {
    date,
    slideNumber: slideNumberShape(slide) !== null,
    footer: footer ? getShapeText(footer) : null,
  };
}

const FALLBACK_NAME = { dt: 'Date', ftr: 'Footer' } as const;
const FOOTER_PT = 12;

/** The slide's date or footer: its placeholder, or the box this module added. */
const find = (slide: SlideData, type: 'dt' | 'ftr') =>
  findSlidePlaceholder(slide, type) ??
  getSlideShapes(slide).find((shape) => getShapeName(shape) === FALLBACK_NAME[type]) ??
  null;

function slot(pres: PresentationData, slide: SlideData, type: 'dt' | 'ftr') {
  const placeholder =
    addSlidePlaceholder(slide, type) ?? addSlidePlaceholder(slide, type, { source: 'master' });
  if (placeholder) return placeholder;
  const size = getSlideSize(pres);
  if (!size) return null;
  const margin = inches(0.25);
  const h = inches(0.4);
  const w = type === 'dt' ? inches(2) : Math.round(size.width / 3);
  const x = type === 'dt' ? margin : Math.round((size.width - w) / 2);
  const box = addSlideTextBox(slide, {
    x: emu(x),
    y: emu(size.height - h - margin),
    w: emu(w),
    h,
    text: ' ',
    name: FALLBACK_NAME[type],
  });
  setShapeTextFormat(box, { size: FOOTER_PT });
  if (type === 'ftr') setParagraphAlignment(box, 0, 'ctr');
  return box;
}

/** Applies the settings to one slide. */
export function applyHeaderFooter(
  pres: PresentationData,
  slide: SlideData,
  settings: HeaderFooter,
): void {
  const hide = settings.hideOnTitle && isTitleSlide(slide);

  // An existing slot is updated in place, so a moved or restyled one stays put.
  const existingDate = find(slide, 'dt');
  if (existingDate && (!settings.date || hide)) removeShape(existingDate);
  if (settings.date && !hide) {
    const dt = existingDate ?? slot(pres, slide, 'dt');
    if (dt) {
      if (settings.date.kind === 'auto') setShapeTextField(dt, settings.date.format);
      else setShapeText(dt, settings.date.text);
    }
  }

  if (settings.slideNumber && !hide) showSlideNumber(pres, slide);
  else hideSlideNumber(slide);

  const existingFooter = find(slide, 'ftr');
  if (existingFooter && (settings.footer === null || hide)) removeShape(existingFooter);
  if (settings.footer !== null && !hide) {
    const ftr = existingFooter ?? slot(pres, slide, 'ftr');
    if (ftr) setShapeText(ftr, settings.footer);
  }
}
