// Home ▸ Section: the reference desktop app's slide sections (`p14:sectionLst`), edited as a
// list of start indices. A deck with sections has every slide in one; adding
// the first section mid-deck puts the slides before it in "Default Section",
// as the reference desktop app does.

import {
  getSlidePartName,
  getSlides,
  getSlideSections,
  setSlideSections,
  type PresentationData,
} from '@office-kit/pptx';

export interface SectionRange {
  readonly name: string;
  /** Index of the section's first slide; sections are contiguous. */
  readonly start: number;
}

export const DEFAULT_SECTION = 'Default Section';
export const UNTITLED_SECTION = 'Untitled Section';

/** The sections as start indices, in slide order (empty sections are dropped). */
export function sectionRanges(pres: PresentationData): SectionRange[] {
  const indexByPart = new Map(
    getSlides(pres).map((slide, index) => [getSlidePartName(slide), index]),
  );
  return getSlideSections(pres)
    .flatMap((section) => {
      const indices = section.slides
        .map((slide) => indexByPart.get(getSlidePartName(slide)))
        .filter((index) => index !== undefined);
      return indices.length ? [{ name: section.name, start: Math.min(...indices) }] : [];
    })
    .sort((a, b) => a.start - b.start);
}

function write(pres: PresentationData, ranges: readonly SectionRange[]): void {
  const slides = getSlides(pres);
  const sorted = [...ranges].sort((a, b) => a.start - b.start);
  setSlideSections(
    pres,
    sorted.map((range, i) => ({
      name: range.name,
      slides: slides.slice(range.start, sorted[i + 1]?.start ?? slides.length),
    })),
  );
}

/** Starts a section at `index`. */
export function addSection(pres: PresentationData, index: number, name = UNTITLED_SECTION): void {
  const ranges = sectionRanges(pres).filter((range) => range.start !== index);
  if (ranges.length === 0 && index > 0) ranges.push({ name: DEFAULT_SECTION, start: 0 });
  write(pres, [...ranges, { name, start: index }]);
}

/** The section the slide at `index` belongs to, or null in a deck without sections. */
export function sectionOf(ranges: readonly SectionRange[], index: number): SectionRange | null {
  return ranges.filter((range) => range.start <= index).at(-1) ?? null;
}

export function renameSection(pres: PresentationData, start: number, name: string): void {
  write(
    pres,
    sectionRanges(pres).map((range) => (range.start === start ? { ...range, name } : range)),
  );
}

/** Removes the section, not its slides: they join the section before it. */
export function removeSection(pres: PresentationData, start: number): void {
  const ranges = sectionRanges(pres).filter((range) => range.start !== start);
  // The first section must start at the first slide.
  if (ranges[0] && ranges[0].start !== 0) ranges[0] = { ...ranges[0], start: 0 };
  write(pres, ranges);
}

export function removeAllSections(pres: PresentationData): void {
  setSlideSections(pres, []);
}
