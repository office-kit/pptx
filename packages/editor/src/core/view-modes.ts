// The editor's views, named after Mac PowerPoint's View menu. Reading View and
// Slide Show are not modes of the editing window: PowerPoint opens them in a
// window of their own.
export type ViewMode =
  | 'normal'
  | 'outline'
  | 'sorter'
  | 'notesPage'
  | 'slideMaster'
  | 'handoutMaster'
  | 'notesMaster';

export function isMasterView(mode: ViewMode): boolean {
  return mode === 'slideMaster' || mode === 'handoutMaster' || mode === 'notesMaster';
}

/** Views that show a printed page (portrait notes or handout page) instead of a slide. */
export function isPageView(mode: ViewMode): boolean {
  return mode === 'notesPage' || mode === 'handoutMaster' || mode === 'notesMaster';
}

/** Views with the slide editor, where the Notes and Comments status bar buttons appear. */
export function isSlideEditingView(mode: ViewMode): boolean {
  return mode === 'normal' || mode === 'outline';
}
