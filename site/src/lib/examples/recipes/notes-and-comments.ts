// Add speaker notes and two review comments to a slide.

import { addSlideComment, cm, setSlideNotes, type SlideData } from '@office-kit/pptx';

declare const slide: SlideData;

setSlideNotes(slide, 'Open the comments pane to see the review thread.');

addSlideComment(slide, {
  author: { name: 'Reviewer A', initials: 'RA' },
  text: 'Tighten the headline.',
  position: { x: cm(2), y: cm(2) },
});
addSlideComment(slide, {
  author: { name: 'Reviewer B', initials: 'RB' },
  text: 'Numbers look strong.',
  position: { x: cm(10), y: cm(6) },
});
