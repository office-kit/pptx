// Brand a from-scratch deck's color scheme and typography — no template
// required. setPresentationTheme / setPresentationFonts patch only the
// slots you name; every other theme slot keeps its default.

import {
  createPresentation,
  setPresentationFonts,
  setPresentationTheme,
  type PresentationData,
} from 'pptx-kit';

const pres: PresentationData = createPresentation();

setPresentationTheme(pres, {
  name: 'Consulting Navy',
  dark1: '#0B1F3A',
  accent1: '#00A9E0',
  accent2: '#FDB913',
  accent3: '#6E7B8B',
});

setPresentationFonts(pres, {
  majorLatin: 'Georgia',
  minorLatin: 'Calibri',
});

// Every shape added from here on inherits the new scheme via the
// scheme-color tokens (accent1, tx1, ...) it was already using.
